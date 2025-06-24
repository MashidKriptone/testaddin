/* global Office */

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        initializeMSAL();
        Office.actions.associate("onMessageSendHandler", onMessageSendHandler);

        console.log('Add-in is running in the background.');
    }
});

// Email regex: validates general email format with 2-3 character domain extensions
const emailRegex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,3}$/;

// Regex patterns for additional checks
const regexPatterns = {
    email: /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,3}$/,
    body: /\b(confidential|prohibited|restricted)\b/i,
    attachmentName: /\.(exe|bat|sh)$/i,
    imei: /[^\d](\d{15}|\d{2}\-\d{6}\-\d{6}\-\d)[^\d]/,
    namesUSCensus: /[^\w]([A-Z][a-z]{1,12}(\s?,\s?|[\s]|\s([A-Z])\.\s)[A-Z][a-z]{1,12})[^\w]/,
    swiftBIC: /[^\w-]([A-Z]{6}[A-Z0-9]{2}([A-Z0-9]{3})?)[^\w-]/,
    timeZoneOffset: /[^\d]([-+](0[0-9]|1[0-2]):00|\+13:00|[+-]0[34]:30|\+0(5:(30|45)|9:30))[^\d]/,
    ibanAndorra: /[^\w](AD\d{2}((\s\d{4}){2}(\s[a-zA-Z0-9]){3}|\d{8}[a-zA-Z0-9]{12}))[^\w]/,
    taxFileNumber: /[^\w-.;&](\d{8,9})[^\w-.;&]/,
    ibanAustria: /[^\w](AT\d{2}(\s(\d{4}\s){3}\d{4}|\d{16}))[^\w]/,
    phoneAustria: /[^\d\-]((\+43[\s\-]|0)(\d{1,4}[\s\-]\d{3}[\s\-]\d{3}([\s\-]\d{3,6})?|\d{1,4}[\s\-]\d{3,12}))[^\d\-]/,
    ssnAustria: /[^\w-.;&](\d{4}[0-3]\d(0[1-9]|1[0-2])\d{2})[^\w-.;&]/,
    ibanAzerbaijan: /[^\w](AZ\d{2}(\s[A-Za-z0-9]{4}\s(\d{4}\s){4}\d{4}|[A-Za-z0-9]{4}\d{20}))[^\w]/,
};


async function onMessageSendHandler(eventArgs) {
    // Track execution time
    const startTime = Date.now();
    console.log('🚀 onMessageSendHandler started at:', new Date().toISOString());

    try {
        // 1. Initialize MSAL if not already done
        if (!isInitialized) {
            console.log('⚙️ Initializing MSAL...');
            initializeMSAL();
        }

        // 2. Authenticate and get access token
        console.log('🔐 Acquiring access token...');
        let token;
        try {
            token = await getAccessToken();
            console.log('✅ Access token acquired successfully');
        } catch (authError) {
            console.error('❌ Authentication failed:', authError);
            await showOutlookNotification("Authentication Required", "Please sign in to continue.");
            eventArgs.completed({ allowEvent: false });
            return;
        }

        // 3. Get the current mail item
        const item = Office.context.mailbox.item;
        console.log('📧 Processing mail item:', item.itemId);

        // 4. Retrieve all email details in parallel
        console.log('📋 Gathering email details...');
        let from, toRecipients, ccRecipients, bccRecipients, subject, body, attachments;
        try {
            [from, toRecipients, ccRecipients, bccRecipients, subject, body, attachments] = await Promise.all([
                getFromAsync(item).catch(e => { console.warn('⚠️ From address error:', e); return ''; }),
                getRecipientsAsync(item.to).catch(e => { console.warn('⚠️ To recipients error:', e); return ''; }),
                getRecipientsAsync(item.cc).catch(e => { console.warn('⚠️ CC recipients error:', e); return ''; }),
                getRecipientsAsync(item.bcc).catch(e => { console.warn('⚠️ BCC recipients error:', e); return ''; }),
                getSubjectAsync(item).catch(e => { console.warn('⚠️ Subject error:', e); return ''; }),
                getBodyAsync(item).catch(e => { console.warn('⚠️ Body error:', e); return ''; }),
                getAttachmentsAsync(item).catch(e => { console.warn('⚠️ Attachments error:', e); return []; })
            ]);
        } catch (error) {
            console.error('❌ Failed to gather email details:', error);
            await showOutlookNotification("Error", "Could not retrieve email details");
            eventArgs.completed({ allowEvent: false });
            return;
        }

        // Log gathered details (redacting sensitive info in production)
        console.log('ℹ️ Email details gathered:', {
            from: from ? `${from.substring(0, 3)}...@...${from.split('@')[1]?.substring(from.split('@')[1].length - 3)}` : 'empty',
            toCount: toRecipients ? toRecipients.split(',').length : 0,
            ccCount: ccRecipients ? ccRecipients.split(',').length : 0,
            bccCount: bccRecipients ? bccRecipients.split(',').length : 0,
            subjectLength: subject?.length || 0,
            bodyLength: body?.length || 0,
            attachmentCount: attachments?.length || 0
        });

        // 5. Validate we have at least one recipient
        if (!toRecipients && !ccRecipients && !bccRecipients) {
            console.warn('❌ No recipients found');
            await showOutlookNotification("Error", "Please add at least one recipient (To, CC, or BCC)");
            eventArgs.completed({ allowEvent: false });
            return;
        }

        // 6. Validate email addresses format
        console.log('🔍 Validating email addresses...');
        if (!validateEmailRecipients(toRecipients, ccRecipients, bccRecipients)) {
            await showOutlookNotification("Invalid Email", "One or more email addresses are invalid");
            eventArgs.completed({ allowEvent: false });
            return;
        }

        // 7. Fetch policy settings
        console.log('⚖️ Fetching policy settings...');
        let policy;
        try {
            policy = await fetchPolicyDomains(token, from);
            console.log('ℹ️ Policy settings:', {
                allowedDomains: policy.allowedDomains?.length || 0,
                blockedDomains: policy.blockedDomains?.length || 0,
                contentScanning: policy.contentScanning,
                attachmentPolicy: policy.attachmentPolicy,
                encryptOutgoing: policy.encryptOutgoingEmails
            });
        } catch (error) {
            console.error('❌ Failed to fetch policy:', error);
            // Fail open - allow sending if policy can't be fetched
            console.warn('⚠️ Allowing send due to policy fetch failure');
        }

        // 8. Apply domain restrictions if policy exists
        if (policy && (policy.allowedDomains?.length > 0 || policy.blockedDomains?.length > 0)) {
            console.log('🔒 Applying domain restrictions...');

            // Check if domain restrictions are enabled
            if (policy.useAllowedDomains || policy.blockedDomains.length > 0) {
                const domainCheckResult = checkDomainRestrictions(
                    toRecipients,
                    ccRecipients,
                    bccRecipients,
                    policy.allowedDomains,
                    policy.blockedDomains,
                    policy.useAllowedDomains
                );

                if (domainCheckResult.blocked) {
                    console.warn(`❌ Blocked domain detected: ${domainCheckResult.domain}`);
                    await showOutlookNotification(
                        "Blocked Domain",
                        `Cannot send to ${domainCheckResult.domain} per company policy`
                    );
                    eventArgs.completed({ allowEvent: false });
                    return;
                }
            }

            // Check if domain requires encryption
            const requiresEncryption = policy.alwaysEncryptDomains.some(domain =>
                getAllRecipients(toRecipients, ccRecipients, bccRecipients)
                    .some(email => email.trim().endsWith(`@${domain}`))
            );

            if (requiresEncryption) {
                policy.encryptOutgoingEmails = true;
                policy.encryptOutgoingAttachments = true;
            }
        }

        // 9. Content scanning if enabled
        if (policy?.contentScanning) {
            console.log('🔎 Scanning email content...');
            const contentScanResult = scanContent(
                body,
                subject,
                attachments,
                policy.customRegexPatterns,
                policy.sensitiveKeywords
            );
            if (contentScanResult.found) {
                console.warn(`❌ Restricted content found: ${contentScanResult.type}`);
                await showOutlookNotification(
                    "Restricted Content",
                    `Cannot send: Email contains restricted ${contentScanResult.type}`
                );
                eventArgs.completed({ allowEvent: false });
                return;
            }
        }

        // 10. Attachment policy checks
        if (policy?.attachmentPolicy && attachments?.length > 0) {
            console.log('📎 Checking attachments...');

            // Check attachment size
            const sizeExceeded = attachments.some(att =>
                att.size > (policy.maxAttachmentSizeMB * 1024 * 1024)
            );
            if (sizeExceeded) {
                await showOutlookNotification(
                    "Attachment Too Large",
                    `Attachments exceed maximum size of ${policy.maxAttachmentSizeMB}MB`
                );
                eventArgs.completed({ allowEvent: false });
                return;
            }

            // Check blocked attachment types
            const attachmentCheckResult = checkAttachments(
                attachments,
                policy.blockedAttachments,
                policy.allowedAttachments,
                policy.requirePasswordProtectedAttachments
            );
            if (attachmentCheckResult.blocked) {
                console.warn(`❌ Blocked attachment: ${attachmentCheckResult.filename}`);
                await showOutlookNotification(
                    "Restricted Attachment",
                    `Cannot send: ${attachmentCheckResult.reason}`
                );
                eventArgs.completed({ allowEvent: false });
                return;
            }
        }

        // 11. Prepare email data for API
        console.log('📦 Preparing email data for API...');
        let emailData;
        try {
            emailData = await prepareEmailData(from, toRecipients, ccRecipients, bccRecipients, subject, body, attachments);
            // Detailed payload logging
            console.group('📤 Email Data Payload');
            console.log('📝 Basic Info:', {
                id: emailData.id,
                from: emailData.fromEmailID,
                timestamp: emailData.timestamp,
                subject: emailData.emailSubject,
                bodyLength: emailData.emailBody.length
            });

            console.log('👥 Recipients:', {
                to: emailData.emailTo,
                cc: emailData.emailCc,
                bcc: emailData.emailBcc
            });

            console.log('📎 Attachments:', emailData.attachments.map(att => ({
                name: att.fileName,
                size: att.fileSize,
                type: att.fileType,
                dataPreview: att.fileData?.substring(0, 50) + '...'
            })));

            console.log('🔢 Full Payload Size:', JSON.stringify(emailData).length, 'bytes');
            console.groupEnd();
        } catch (error) {
            console.error("Error preparing email data:", error);
            await showOutlookNotification("Error", "Failed to prepare email for sending");
            eventArgs.completed({ allowEvent: false });
            return;
        }


        // Replace the current encryption handling block with this:
        if (policy?.encryptOutgoingEmails || policy?.encryptOutgoingAttachments) {
            console.log("🔐 Beginning encryption process...");

            try {
                const encryptedResult = await getEncryptedEmail(emailData, token);

                if (encryptedResult.encryptedAttachments?.length > 0) {
                    console.log("✅ Encryption successful, updating email");

                    // Use the instruction note as the new body content
                    const newBodyContent = encryptedResult.instructionNote ||
                        "<p>This email contains encrypted content. Please use the attached files to view the secure message.</p>";

                    await updateEmailWithEncryptedContent(
                        item,
                        encryptedResult.encryptedAttachments,
                        newBodyContent
                    );

                    eventArgs.completed({ allowEvent: true });
                    return;
                } else {
                    console.warn("⚠️ Encryption required but no encrypted content returned");
                    await showOutlookNotification(
                        "Encryption Required",
                        "This email requires encryption but the service is unavailable. Email not sent."
                    );
                    eventArgs.completed({ allowEvent: false });
                    return;
                }
            } catch (encryptionError) {
                console.error("❌ Encryption process failed:", encryptionError);
                await showOutlookNotification(
                    "Encryption Failed",
                    "This email requires encryption but the service failed. Email not sent."
                );
                eventArgs.completed({ allowEvent: false });
                return;
            }
        }


        // 13. If no encryption needed, just save the email data
        // Replace the current saveEmailData call with this:
        if (!policy?.encryptOutgoingEmails && !policy?.encryptOutgoingAttachments) {
            console.log('💾 Saving email data (non-encrypted path)...');
            try {
                const saveResult = await saveEmailData(emailData, token);
                if (!saveResult.success) {
                    console.error('❌ Failed to save email data:', saveResult.message);
                    await showOutlookNotification("Warning", "Email will be sent but audit logging failed: " + saveResult.message);
                }
            } catch (error) {
                console.error('❌ Failed to save email data:', error);
                // Fail open - allow sending even if saving fails
                console.warn('⚠️ Allowing send despite save failure');
            }
        }

        // 14. All checks passed - allow the email to send
        console.log('✅ All checks passed - allowing send');
        await showSuccessNotification("Email sent successfully");
        eventArgs.completed({ allowEvent: true });

    } catch (error) {
        console.error('❌ Unhandled error in onMessageSendHandler:', error);
        await showOutlookNotification(
            "Error",
            "An unexpected error occurred while sending the email. Please try again."
        );
        eventArgs.completed({ allowEvent: false });
    } finally {
        console.log(`⏱️ Handler completed in ${Date.now() - startTime}ms`);
    }
}

// Helper Functions

/**
 * Validates all recipient email addresses
 */
function validateEmailRecipients(to, cc, bcc) {
    const validate = (recipients) =>
        !recipients || recipients.split(',').every(email => emailRegex.test(email.trim()));

    return validate(to) && validate(cc) && validate(bcc);
}

function getAllRecipients(to, cc, bcc) {
    return [
        ...(to ? to.split(',').filter(e => e.trim()) : []),
        ...(cc ? cc.split(',').filter(e => e.trim()) : []),
        ...(bcc ? bcc.split(',').filter(e => e.trim()) : [])
    ];
}

// Update checkDomainRestrictions to handle allow/block lists
function checkDomainRestrictions(to, cc, bcc, allowedDomains, blockedDomains, useAllowList) {
    const allRecipients = getAllRecipients(to, cc, bcc);

    for (const email of allRecipients) {
        const domain = email.split('@')[1]?.toLowerCase();
        if (!domain) continue;

        // Check against block list first
        if (blockedDomains.includes(domain)) {
            return { blocked: true, domain, reason: 'blocked by policy' };
        }

        // Check against allow list if enabled
        if (useAllowList && !allowedDomains.includes(domain)) {
            return { blocked: true, domain, reason: 'not in allowed domains' };
        }
    }

    return { blocked: false };
}


function scanContent(body, subject, attachments, customPatterns = [], keywords = []) {
    const textToScan = `${subject} ${body}`.toLowerCase();

    // Check custom regex patterns from policy
    for (const pattern of customPatterns) {
        try {
            const regex = new RegExp(pattern, 'i');
            if (regex.test(textToScan)) {
                return { found: true, type: 'custom policy violation' };
            }
        } catch (e) {
            console.warn('Invalid regex pattern:', pattern);
        }
    }

    // Check sensitive keywords
    for (const keyword of keywords) {
        if (textToScan.includes(keyword.toLowerCase())) {
            return { found: true, type: `sensitive keyword: ${keyword}` };
        }
    }

    // Check attachment names
    for (const attachment of attachments || []) {
        if (regexPatterns.attachmentName.test(attachment.name)) {
            return { found: true, type: 'attachment: ' + attachment.name };
        }
    }

    return { found: false };
}


/**
 * Updates the email with encrypted content
 */
async function updateEmailWithEncryptedContent(item, encryptedAttachments, newBodyContent) {
    try {
        // 1. First update the body with the instruction note
        await new Promise((resolve, reject) => {
            item.body.setAsync(
                newBodyContent,
                {
                    coercionType: Office.CoercionType.Html,
                    asyncContext: null
                },
                (result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        resolve();
                    } else {
                        reject(new Error(`Failed to update body: ${result.error.message}`));
                    }
                }
            );
        });

        // 2. Remove existing attachments
        if (encryptedAttachments && encryptedAttachments.length > 0) {
            const currentAttachments = await new Promise(resolve => {
                item.getAttachmentsAsync(resolve);
            });

            if (currentAttachments.value?.length > 0) {
                await Promise.all(currentAttachments.value.map(att =>
                    new Promise(resolve => {
                        item.removeAttachmentAsync(att.id, resolve);
                    })
                ));
            }


            // 3. Add new encrypted attachments
            for (const attachment of encryptedAttachments) {
                if (attachment.fileData) {
                    // Ensure we have clean base64 data (remove data URI prefix if present)
                    const cleanBase64 = attachment.fileData.replace(/^data:[^;]+;base64,/, '');

                    await new Promise((resolve, reject) => {
                        item.addFileAttachmentFromBase64Async(
                            cleanBase64,
                            attachment.fileName || "encrypted-file.ksf",
                            { isInline: false },
                            (result) => {
                                if (result.status === Office.AsyncResultStatus.Succeeded) {
                                    resolve();
                                } else {
                                    reject(new Error(`Failed to add attachment: ${result.error.message}`));
                                }
                            }
                        );
                    });
                }
            }
        }

        console.log("✅ Email processed with encryption settings");
    } catch (error) {
        console.error("❌ Failed to update email:", error);
        throw error;
    }
}
// Get user details from Microsoft Graph
async function getUserDetails(accessToken) {
    try {
        const response = await fetch('https://graph.microsoft.com/v1.0/me', {
            headers: {
                'Authorization': `Bearer ${accessToken}`
            }
        });

        if (!response.ok) {
            throw new Error(`Graph API request failed with status ${response.status}`);
        }

        return await response.json();
    } catch (error) {
        console.error("Error fetching user details:", error);
        throw error;
    }
}

// Update the fetchPolicyDomains function to match the new response structure
async function fetchPolicyDomains(token, from) {
    try {
        console.log('🔍 Fetching policy from API...');
        const response = await fetch(`https://kntrolemail.kriptone.com:6677/api/Policy/GetPolicyByEmailAsync/${from}`, {
            method: 'GET',
            headers: {
                'Content-Type': 'application/json',
                'Accept': 'application/json',
                'Authorization': `Bearer ${token}`,
                "X-Tenant-ID": "kriptone.com"
            },
        });

        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(`Policy API responded with status ${response.status}: ${errorText}`);
        }

        const responseData = await response.json();

        // Check if the response contains the data object
        if (!responseData.success || !responseData.data) {
            throw new Error('Invalid policy API response structure');
        }

        const policy = responseData.data;
        console.log("🔹 Policy Response:", policy);

        // Map the API response to our expected format
        return {
            // Basic policy info
            policyName: policy.policyName || 'Default Policy',
            isEnabled: policy.isEnabled !== false, // Default to true if not specified
            enableIRM: policy.enableIRM === true,
            enableLogging: policy.enableLogging === true,

            // Domain policy
            allowedDomains: policy.domainPolicy?.allowedDomains || [],
            blockedDomains: policy.domainPolicy?.blockedDomains || [],
            alwaysEncryptDomains: policy.domainPolicy?.alwaysEncryptDomains || [],
            useAllowedDomains: policy.domainPolicy?.useAllowedDomains === true,

            // Attachment policy
            attachmentPolicy: policy.attachmentPolicy?.useAllowedAttachments === true,
            allowedAttachments: policy.attachmentPolicy?.allowedAttachments || [],
            blockedAttachments: policy.attachmentPolicy?.blockedAttachments || [],
            maxAttachmentSizeMB: policy.attachmentPolicy?.maxAttachmentSizeMB || 10,
            encryptOutgoingAttachments: policy.attachmentPolicy?.encryptOutgoingAttachments === true,
            requirePasswordProtectedAttachments: policy.attachmentPolicy?.requirePasswordProtectedAttachments === true,

            // Regex/content scanning
            contentScanning: policy.regexPolicy?.contentScanning === true,
            customRegexPatterns: policy.regexPolicy?.customRegexPatterns || [],
            sensitiveKeywords: policy.regexPolicy?.sensitiveKeywords || [],

            // Encryption policy
            encryptOutgoingEmails: policy.encryptionPolicy?.encryptOutgoingEmails === true,
            enableEncryption: policy.encryptionPolicy?.enableEncryption === true,

            // IRM policy
            irmPolicy: policy.irmPolicy ? {
                expiryDate: policy.irmPolicy.expiryDate,
                maxOpenCount: policy.irmPolicy.maxOpenCount,
                maxFailedAttempts: policy.irmPolicy.maxFailedAttempts,
                lockOnFailure: policy.irmPolicy.lockOnFailure === true,
                blockCopy: policy.irmPolicy.blockCopy === true,
                blockPrint: policy.irmPolicy.blockPrint === true,
                blockSaveAs: policy.irmPolicy.blockSaveAs === true,
                blockEdit: policy.irmPolicy.blockEdit === true,
                blockScreenCapture: policy.irmPolicy.blockScreenCapture === true,
                allowedUsers: policy.irmPolicy.allowedUsers || [],
                allowedLocations: policy.irmPolicy.allowedLocations || [],
                blockedLocations: policy.irmPolicy.blockedLocations || []
            } : null
        };

    } catch (error) {
        console.error("❌ Error fetching policy:", error);
        return getDefaultPolicy();
    }
}

// Default policy to use when API fails or returns empty
function checkAttachments(attachments, blockedTypes = [], allowedTypes = [], requirePassword = false) {
    for (const attachment of attachments) {
        const ext = `.${attachment.name.split('.').pop().toLowerCase()}`;

        // Check blocked extensions
        if (blockedTypes.includes(ext)) {
            return {
                blocked: true,
                filename: attachment.name,
                reason: `Attachment type ${ext} is blocked by policy`
            };
        }

        // Check if using allow list and attachment not in it
        if (allowedTypes.length > 0 && !allowedTypes.includes(ext)) {
            return {
                blocked: true,
                filename: attachment.name,
                reason: `Attachment type ${ext} is not allowed by policy`
            };
        }

        // Check for password protection requirement
        if (requirePassword && !attachment.isPasswordProtected) {
            return {
                blocked: true,
                filename: attachment.name,
                reason: 'Attachment must be password protected'
            };
        }
    }
    return { blocked: false };
}

// Update the getDefaultPolicy function
function getDefaultPolicy() {
    return {
        policyName: 'Default Security Policy',
        isEnabled: true,
        enableIRM: false,
        enableLogging: true,

        // Domain policy defaults
        allowedDomains: [],
        blockedDomains: [],
        alwaysEncryptDomains: [],
        useAllowedDomains: false,

        // Attachment policy defaults
        attachmentPolicy: true,
        allowedAttachments: ['.pdf', '.docx', '.xlsx', '.pptx', '.jpg', '.png'],
        blockedAttachments: ['.exe', '.bat', '.sh', '.dll', '.msi'],
        maxAttachmentSizeMB: 10,
        encryptOutgoingAttachments: false,
        requirePasswordProtectedAttachments: false,

        // Content scanning defaults
        contentScanning: true,
        customRegexPatterns: [],
        sensitiveKeywords: ['confidential', 'proprietary', 'secret'],

        // Encryption defaults
        encryptOutgoingEmails: false,
        enableEncryption: false,

        // IRM defaults
        irmPolicy: null
    };
}
async function getEncryptedEmail(emailDataDto, token) {
    try {
        console.log("📤 Sending email data to encryption API");
        const response = await fetch("https://kntrolemail.kriptone.com:6677/api/Email", {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                "Authorization": `Bearer ${token}`,
                "X-Tenant-ID": "kriptone.com"
            },
            body: JSON.stringify(emailDataDto)
        });

        // First check response status
        if (!response.ok) {
            // Read the error response once and store it
            const errorResponse = await response.text();
            console.error("🔴 API Error Response:", {
                status: response.status,
                errorResponse
            });
            throw new Error(`Encryption failed with status ${response.status}`);
        }

        // Now read the successful response
        const responseData = await response.json();
        console.log("🔹 Raw API Response:", JSON.stringify(responseData, null, 2));
        // Validate response structure
        if (!responseData.encryptedAttachments || responseData.encryptedAttachments.length === 0) {
            throw new Error("API response missing encrypted attachments");
        }

        return {
            encryptedAttachments: responseData.encryptedAttachments || [],
            instructionNote: responseData.instructionNote,
            encryptedEmailBody: responseData.encryptedEmailBody
        };

    } catch (error) {
        console.error("❌ Full encryption error details:", {
            error: error.message,
            stack: error.stack,
            requestPayload: {
                ...emailDataDto,
                attachments: emailDataDto.attachments.map(att => ({
                    fileName: att.fileName,
                    size: att.fileSize,
                    type: att.fileType,
                    dataLength: att.fileData?.length || 0
                }))
            }
        });
        throw error;
    }
}
async function saveEmailData(emailData, token) {
    try {
        console.log("📤 Attempting to save email data...");
        const response = await fetch('https://kntrolemail.kriptone.com:6677/api/Email', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Accept': 'application/json',
                'Authorization': `Bearer ${token}`,
                "X-Tenant-ID": "kriptone.com"
            },
            body: JSON.stringify(emailData),
        });

        // First check if we got any response
        if (!response) {
            throw new Error("No response received from server");
        }

        // Try to parse JSON response
        let json;
        try {
            json = await response.json();
            console.log("🔹 API Response:", json);
        } catch (parseError) {
            // If JSON parsing fails, get the text response
            const textResponse = await response.text();
            throw new Error(`Invalid JSON response: ${textResponse}`);
        }

        // Check if response indicates success
        if (!response.ok) {
            throw new Error(json.message || `Server returned status ${response.status}`);
        }

        return {
            success: true,
            message: json.message || "Email data saved successfully"
        };
    } catch (error) {
        console.error("❌ Detailed save error:", {
            error: error.message,
            requestPayload: {
                ...emailData,
                attachments: emailData.attachments.map(att => ({
                    fileName: att.fileName,
                    size: att.fileSize,
                    type: att.fileType,
                    dataLength: att.fileData?.length || 0
                }))
            }
        });
        return {
            success: false,
            message: error.message || "Failed to save email data"
        };
    }
}

// Helper functions
function validateEmailAddresses(recipients) {
    return recipients ? recipients.split(',').every(email => emailRegex.test(email.trim())) : true;
}

function isDomainBlocked(recipients, blockedDomains) {
    if (!blockedDomains || blockedDomains.length === 0) return false;

    const recipientArray = recipients ? recipients.split(',').map(email => email.trim()) : [];
    const blockedEmail = recipientArray.find(email => blockedDomains.includes(email.split('@')[1]));

    if (blockedEmail) {
        console.log(`🔴 Blocked Email Found: ${blockedEmail}`);
        return true;
    }
    return false;
}
async function prepareEmailData(from, to, cc, bcc, subject, body, attachments) {
    const emailId = generateUUID();
    console.log("📥 Preparing email data with ID:", emailId);

    // Process attachments with strict validation
    const attachmentPayloads = [];
    for (const attachment of attachments || []) {
        if (!attachment || !attachment.id) {
            console.warn("⚠️ Skipping invalid attachment:", attachment);
            continue;
        }

        try {
            const base64Data = await getAttachmentBase64Fallback(attachment);

            if (!base64Data) {
                throw new Error("Attachment returned empty data");
            }

            attachmentPayloads.push({
                id: generateUUID(),
                fileName: attachment.name || 'Unknown',
                fileSize: attachment.size || 0,
                fileType: attachment.attachmentType || 'application/octet-stream',
                uploadTime: new Date().toISOString(),
                fileData: base64Data
            });
        } catch (error) {
            console.error(`❌ Failed to process attachment ${attachment.name}:`, error);
            throw new Error(`Could not process attachment ${attachment.name}`);
        }
    }

    // Validate we have at least one valid recipient
    const allRecipients = [
        ...(to ? to.split(',').map(e => e.trim()).filter(e => e) : []),
        ...(cc ? cc.split(',').map(e => e.trim()).filter(e => e) : []),
        ...(bcc ? bcc.split(',').map(e => e.trim()).filter(e => e) : [])
    ];

    if (allRecipients.length === 0) {
        throw new Error("No valid recipients specified");
    }

    return {
        id: emailId,
        fromEmailID: from || "",
        emailTo: to ? to.split(',').map(e => e.trim()).filter(e => e) : [],
        emailCc: cc ? cc.split(',').map(e => e.trim()).filter(e => e) : [],
        emailBcc: bcc ? bcc.split(',').map(e => e.trim()).filter(e => e) : [],
        emailSubject: subject || "(No Subject)",
        emailBody: body || "",
        timestamp: new Date().toISOString(),
        attachments: attachmentPayloads
    };
}

// Fallback attachment method
async function getAttachmentBase64Fallback(attachment) {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.getAttachmentContentAsync(
            attachment.id,
            (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    if (result.value.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
                        resolve(result.value.content);
                    } else {
                        reject(new Error("Attachment content not in Base64 format"));
                    }
                } else {
                    reject(result.error);
                }
            }
        );
    });
}

function generateUUID() {
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
        const r = Math.random() * 16 | 0, v = c === 'x' ? r : (r & 0x3 | 0x8);
        return v.toString(16);
    });
}


// Async functions to retrieve email details
function getFromAsync(item) {
    return new Promise((resolve, reject) => {
        item.from.getAsync(result => result.status === Office.AsyncResultStatus.Succeeded ? resolve(result.value.emailAddress) : reject(result.error));
    });
}

function getRecipientsAsync(recipients) {
    return new Promise((resolve, reject) => {
        recipients.getAsync(result => result.status === Office.AsyncResultStatus.Succeeded ? resolve(result.value.map(r => r.emailAddress).join(", ")) : reject(result.error));
    });
}

function getSubjectAsync(item) {
    return new Promise((resolve, reject) => {
        item.subject.getAsync(result => result.status === Office.AsyncResultStatus.Succeeded ? resolve(result.value) : reject(result.error));
    });
}

function getBodyAsync(item) {
    return new Promise((resolve, reject) => {
        item.body.getAsync(Office.CoercionType.Text, result => result.status === Office.AsyncResultStatus.Succeeded ? resolve(result.value) : reject(result.error));
    });
}

function getAttachmentsAsync(item) {
    return new Promise((resolve, reject) => {
        item.getAttachmentsAsync(result => result.status === Office.AsyncResultStatus.Succeeded ? resolve(result.value) : reject(result.error));
    });
}


async function showOutlookNotification(title, message) {
    Office.context.mailbox.item.notificationMessages.addAsync("error", {
        type: "errorMessage",
        message: `${title}: ${message}`,
    });
}
async function showSuccessNotification(message) {
    try {
        await Office.context.mailbox.item.notificationMessages.addAsync("success", {
            type: "informationalMessage",
            message: `Success: ${message}`
        });
        console.log(`✅ Success notification shown: ${message}`);
    } catch (error) {
        console.error("❌ Failed to show success notification:", error);
    }
}

// MSAL Configuration
const msalConfig = {
    auth: {
        clientId: "7b7b9a2e-eff4-4af2-9e37-b0df0821b144",
        authority: "https://login.microsoftonline.com/common",
        redirectUri: "https://mashidkriptone.github.io/testaddin/redirect.html"
    },
    cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: true
    }
};

// MSAL instance and state
let msalInstance;
let isInitialized = false;
let authInProgress = false;

// Initialize MSAL
function initializeMSAL() {
    if (isInitialized) return;

    msalInstance = new msal.PublicClientApplication(msalConfig);
    isInitialized = true;

    // Handle redirect response if any
    msalInstance.handleRedirectPromise()
        .then(handleAuthResponse)
        .catch(error => {
            console.error("Redirect handling error:", error);
        });
    // Set up button event listeners
    document.getElementById("signInButton").addEventListener("click", signIn);
    document.getElementById("signOutButton").addEventListener("click", signOut);

    // Update UI based on current auth state
    updateUI();
}

// Handle authentication response
function handleAuthResponse(response) {
    if (response) {
        console.log("Authentication successful:", response);
        return response;
    }
    return null;
}

// Get access token with proper error handling
async function getAccessToken() {
    if (!isInitialized) {
        initializeMSAL();
    }

    try {
        // Check for existing accounts
        const accounts = msalInstance.getAllAccounts();
        if (accounts.length === 0) {
            return await loginAndGetToken();
        }

        // Try silent token acquisition
        const silentRequest = {
            scopes: ["User.Read", "Mail.Send"],
            account: accounts[0]
        };

        try {
            const response = await msalInstance.acquireTokenSilent(silentRequest);
            return response.accessToken;
        } catch (silentError) {
            console.log("Silent token failed, trying popup:", silentError);
            return await loginAndGetToken();
        }
    } catch (error) {
        console.error("Error in getAccessToken:", error);
        throw error;
    }
}

// Perform interactive login
async function loginAndGetToken() {
    if (authInProgress) {
        throw new Error("Authentication already in progress");
    }

    authInProgress = true;
    try {
        const loginRequest = {
            scopes: ["User.Read", "Mail.Send"],
            prompt: "select_account"
        };

        // First login
        const loginResponse = await msalInstance.loginPopup(loginRequest);

        // Then get token
        const tokenRequest = {
            scopes: ["User.Read", "Mail.Send"],
            account: loginResponse.account
        };

        const tokenResponse = await msalInstance.acquireTokenSilent(tokenRequest);
        return tokenResponse.accessToken;
    } finally {
        authInProgress = false;
    }
}

// Sign in function
async function signIn() {
    try {
        const loginRequest = {
            scopes: ["User.Read", "Mail.Send"],
            prompt: "select_account"
        };

        const loginResponse = await msalInstance.loginPopup(loginRequest);
        console.log("Login successful:", loginResponse);
        updateUI();
    } catch (error) {
        console.error("Login error:", error);
        showOutlookNotification("Login Failed", "Please try signing in again.");
    }
}

// Sign out function
async function signOut() {
    try {
        const accounts = msalInstance.getAllAccounts();
        if (accounts.length > 0) {
            const logoutRequest = {
                account: accounts[0],
                postLogoutRedirectUri: window.location.origin
            };
            await msalInstance.logoutPopup(logoutRequest);
        }
        console.log("Logout successful");
        updateUI();
    } catch (error) {
        console.error("Logout error:", error);
        showOutlookNotification("Logout Failed", "Please try signing out again.");
    }
}

function updateUI() {
    const accounts = msalInstance?.getAllAccounts() || [];
    const isSignedIn = accounts.length > 0;

    document.getElementById("signInButton").style.display = isSignedIn ? "none" : "block";
    document.getElementById("signOutButton").style.display = isSignedIn ? "block" : "none";

    if (isSignedIn) {
        console.log("User is signed in as:", accounts[0].username);
    } else {
        console.log("User is signed out");
    }
}

async function fetchEmails(token) {
    try {
        const response = await fetch('https://graph.microsoft.com/v1.0/me/messages?$top=10', {
            headers: {
                'Authorization': `Bearer ${token}`
            }
        });

        if (!response.ok) {
            throw new Error(`Failed to fetch emails: ${response.statusText}`);
        }

        const emails = await response.json();
        console.log("🔹 Recent emails:", emails);
        return emails;
    } catch (error) {
        console.error("Error fetching emails:", error);
        throw error;
    }
}

