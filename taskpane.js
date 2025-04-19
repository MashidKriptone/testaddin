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
            policy = await fetchPolicyDomains(token);
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
            const domainCheckResult = checkDomainRestrictions(
                toRecipients, 
                ccRecipients, 
                bccRecipients, 
                policy.allowedDomains, 
                policy.blockedDomains
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

        // 9. Content scanning if enabled
        if (policy?.contentScanning) {
            console.log('🔎 Scanning email content...');
            const contentScanResult = scanContent(body, subject, attachments);
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
            const attachmentCheckResult = checkAttachments(attachments, policy.blockedAttachments);
            if (attachmentCheckResult.blocked) {
                console.warn(`❌ Blocked attachment: ${attachmentCheckResult.filename}`);
                await showOutlookNotification(
                    "Restricted Attachment", 
                    `Cannot send: Attachment "${attachmentCheckResult.filename}" is restricted`
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
            console.log("ℹ️ Prepared email data structure:", {
                id: emailData.id,
                from: emailData.fromEmailID,
                toCount: emailData.emailTo.length,
                ccCount: emailData.emailCc.length,
                bccCount: emailData.emailBcc.length,
                subjectLength: emailData.emailSubject.length,
                bodyLength: emailData.emailBody.length,
                attachmentCount: emailData.attachments.length
            });
        } catch (error) {
            console.error("Error preparing email data:", error);
            await showOutlookNotification("Error", "Failed to prepare email for sending");
            eventArgs.completed({ allowEvent: false });
            return;
        }

        // Handle encryption if required
        if (policy?.encryptOutgoingEmails || policy?.encryptOutgoingAttachments) {
            console.log("🔐 Beginning encryption process...");
            
            try {
                const encryptedResult = await getEncryptedEmail(emailData, token);
                
                if (!encryptedResult) {
                    throw new Error("No response from encryption service");
                }

                console.log("ℹ️ Encryption result received");
                await updateEmailWithEncryptedContent(item, encryptedResult);
                
                console.log("✅ Email prepared with encryption");
                eventArgs.completed({ allowEvent: true });
                return;
            } catch (error) {
                console.error("❌ Encryption failed:", error);
                await showOutlookNotification(
                    "Encryption Error", 
                    "Failed to encrypt email. Please try again or contact support."
                );
                eventArgs.completed({ allowEvent: false });
                return;
            }
        }

        // 13. If no encryption needed, just save the email data
        console.log('💾 Saving email data...');
        try {
            const saveResult = await saveEmailData(emailData, token);
            if (!saveResult.success) {
                throw new Error(saveResult.message || 'Unknown error saving email');
            }
            console.log('✅ Email data saved successfully');
        } catch (error) {
            console.error('❌ Failed to save email data:', error);
            // Fail open - allow sending even if saving fails
            console.warn('⚠️ Allowing send despite save failure');
        }

        // 14. All checks passed - allow the email to send
        console.log('✅ All checks passed - allowing send');
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

/**
 * Checks domain restrictions against policy
 */
function checkDomainRestrictions(to, cc, bcc, allowedDomains, blockedDomains) {
    const allRecipients = [
        ...(to ? to.split(',') : []),
        ...(cc ? cc.split(',') : []),
        ...(bcc ? bcc.split(',') : [])
    ].map(e => e.trim());

    for (const email of allRecipients) {
        const domain = email.split('@')[1];
        
        // Check blocked domains first
        if (blockedDomains?.includes(domain)) {
            return { blocked: true, domain };
        }
        
        // If allowed domains exist, enforce whitelist
        if (allowedDomains?.length > 0 && !allowedDomains.includes(domain)) {
            return { blocked: true, domain };
        }
    }
    
    return { blocked: false };
}

/**
 * Scans email content for restricted patterns
 */
function scanContent(body, subject, attachments) {
    const textToScan = `${subject} ${body}`.toLowerCase();
    
    // Check each regex pattern
    for (const [type, pattern] of Object.entries(regexPatterns)) {
        if (pattern.test(textToScan)) {
            return { found: true, type };
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
 * Checks attachments against blocked types
 */
function checkAttachments(attachments, blockedTypes = []) {
    for (const attachment of attachments) {
        const ext = attachment.name.split('.').pop().toLowerCase();
        if (blockedTypes.includes(ext)) {
            return { blocked: true, filename: attachment.name };
        }
    }
    return { blocked: false };
}

/**
 * Updates the email with encrypted content
 */
async function updateEmailWithEncryptedContent(item, encryptedResult) {
    // Update body with instructions
    await new Promise((resolve, reject) => {
        item.body.setAsync(
                encryptedResult.instructionNote,
            { coercionType: Office.CoercionType.Html },
            (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    resolve();
                } else {
                        reject(new Error(`Failed to update body: ${result.error.message}`));
                }
            }
        );
    });

    // Remove existing attachments
    const attachments = await new Promise((resolve) => {
        item.getAttachmentsAsync(resolve);
    });

    if (attachments.value?.length > 0) {
        await Promise.all(attachments.value.map(att => 
            new Promise((resolve) => {
                item.removeAttachmentAsync(att.id, resolve);
            })
        ));
    }

    // Add encrypted attachment
    await new Promise((resolve, reject) => {
        item.addFileAttachmentFromBase64Async(
            encryptedResult.encryptedFile,
            encryptedResult.fileName || "secure-message.ksf",
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

// Fetch policy domains from the backend
async function fetchPolicyDomains(token) {
    try {
        const response = await fetch('https://kntrolemail.kriptone.com:6677/api/Policy', {
            method: 'GET',
            headers: { 'Content-Type': 'application/json', 'Accept': 'application/json', 'Authorization': `Bearer ${token}`, "X-Tenant-ID": "kriptone.com", },
        });

        if (!response.ok) throw new Error('Failed to fetch policy data: ' + response.statusText);

        const json = await response.json();
        console.log("🔹 Raw API Response:", JSON.stringify(json, null, 2));
        const policy = json[0];

        return {
            allowedDomains: policy.allowedDomains || [],
            blockedDomains: policy.blockedDomains || [],
            contentScanning: policy.contentScanning,
            attachmentPolicy: policy.attachmentPolicy,
            blockedAttachments: policy.blockedAttachments || [],
            encryptOutgoingEmails: policy.encryptOutgoingEmails || false,
            encryptOutgoingAttachments: policy.encryptOutgoingAttachments || false
        };
    } catch (error) {
        console.error("❌ Error fetching policy:", error);
        return { allowedDomains: [], blockedDomains: [], contentScanning: false, attachmentPolicy: false, blockedAttachments: [] };
    }
}

async function getEncryptedEmail(emailDataDto, token) {
    try {
        console.log("📤 Sending email data to encryption API", {
            emailId: emailDataDto.id,
            from: emailDataDto.fromEmailID,
            recipientCount: emailDataDto.emailTo.length + emailDataDto.emailCc.length + emailDataDto.emailBcc.length,
            attachmentCount: emailDataDto.attachments.length
        });

        const response = await fetch("https://kntrolemail.kriptone.com:6677/api/Email", {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                "Authorization": `Bearer ${token}`,
                "X-Tenant-ID": "kriptone.com"
            },
            body: JSON.stringify({
                ...emailDataDto,
                // Ensure we're not sending excessively large attachments
                attachments: emailDataDto.attachments.map(att => ({
                    ...att,
                    fileData: att.fileData.length > 1000000 ? "[LARGE_FILE_TRUNCATED]" : att.fileData
                }))
            })
        });

        if (!response.ok) {
            let errorDetails;
            try {
                errorDetails = await response.json();
            } catch (e) {
                errorDetails = await response.text();
            }

            console.error("🔴 API Error Response:", {
                status: response.status,
                errorDetails
            });

            throw new Error(`Email encryption failed: ${response.status} - ${JSON.stringify(errorDetails)}`);
        }

        return await response.json();
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
async function saveEmailData(emailData,token) {
    try {
        const response = await fetch('https://kntrolemail.kriptone.com:6677/api/Email', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json', 'Accept': 'application/json','Authorization': `Bearer ${token}`,"X-Tenant-ID": "kriptone.com", },
            body: JSON.stringify(emailData),
        });

        const json = await response.json();

        return {
            success: response.ok && json.success,
            message: json.message || "Unknown error",
        };
    } catch (error) {
        console.error("❌ Error saving email data:", error);
        return {
            success: false,
            message: "Unable to connect to the server. Please try again later.",
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
    for (const attachment of attachments) {
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
                fileData: base64Data,
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
        fromEmailID: from,
        emailTo: to ? to.split(',').map(e => e.trim()).filter(e => e) : [],
        emailCc: cc ? cc.split(',').map(e => e.trim()).filter(e => e) : [],
        emailBcc: bcc ? bcc.split(',').map(e => e.trim()).filter(e => e) : [],
        emailSubject: subject || "(No Subject)",
        emailBody: body || "",
        timestamp: new Date().toISOString(),
        attachments: attachmentPayloads,
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
  
  async function fetchAttachmentBase64UsingGraph(itemId, attachmentId) {
    try {
        const token = await getAccessToken();
        const response = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${itemId}/attachments/${attachmentId}/$value`, {
            headers: {
                'Authorization': `Bearer ${token}`
            }
        });

        if (!response.ok) {
            throw new Error(`Graph API request failed with status ${response.status}`);
        }

        const blob = await response.blob();
        return await new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onloadend = () => {
                const base64String = reader.result.split(',')[1];
                resolve(base64String);
            };
            reader.onerror = reject;
            reader.readAsDataURL(blob);
        });
    } catch (error) {
        console.error("Graph API attachment fetch failed:", error);
        throw error;
    }
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

function showOutlookNotification(title, message) {
    Office.context.mailbox.item.notificationMessages.addAsync("error", {
        type: "errorMessage",
        message: `${title}: ${message}`,
    });
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

// function formatTokenResponse(response) {
//     return {
//         access_token: response.accessToken,
//         id_token: response.idToken,
//         expires_in: Math.floor((response.expiresOn.getTime() - Date.now()) / 1000),
//         token_type: "Bearer",
//         scope: response.scopes.join(" "),
//         account: {
//             username: response.account.username,
//             name: response.account.name
//         }
//     };
// }

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