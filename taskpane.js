/* global Office, msal */

Office.onReady((info) => {
    console.log("Office ready");
    console.log('Office Host:', Office.context.host);
    console.log('Office Version:', Office.context.diagnostics.hostVersion);
    console.log('Office.addin:', Office.addin);


    if (info.host === Office.HostType.Outlook) {
        // First associate handlers
        Office.actions.associate("onNewMessageCompose", onNewMessageCompose);
        Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
        // initializeMSAL();
        initializeUI();
        registerIRMFunctions();
    }
});
async function onNewMessageCompose(event) {
    console.log("🚀 Launch Event: onNewMessageCompose triggered");

    try {
        if (Office.addin && Office.addin.showAsTaskpane) {
            await Office.addin.showAsTaskpane();
            console.log("✅ Taskpane opened using Office.addin");
        } else {
            console.warn("⚠️ Office.addin not available, fallback used");
            Office.context.ui.displayDialogAsync(
                'https://mashidkriptone.github.io/testaddin/taskpane.html',
                {
                    height: 74,
                    width: 26,
                    left: 0,
                    top: 0,
                    promptBeforeOpen: false
                },
                (result) => {
                    console.log("Fallback dialog opened");
                }
            );
        }
    } catch (err) {
        console.error("❌ Error opening taskpane:", err);
    } finally {
        event.completed();
    }
}



// IRM Settings
let irmSettings = {
    blockCopy: false,
    blockPrint: false,
    blockSaveAs: false,
    blockEdit: false,
    blockScreenCapture: false,
    lockOnFailure: false,
    sendAccessAck: false,
    maxOpenCount: null,
    expireOn: null,
    maxFailAttempts: 5,
    recipientRestrictions: "none",
    policyApplied: false
};

// Current policy from server
let currentPolicy = null;



// Initialize UI elements
function initializeUI() {

    // IRM Control Event Listeners
    document.getElementById("blockCopyCheckbox").addEventListener("change", () => toggleIRMControl('blockCopy'));
    document.getElementById("blockPrintCheckbox").addEventListener("change", () => toggleIRMControl('blockPrint'));
    document.getElementById("blockSaveAsCheckbox").addEventListener("change", () => toggleIRMControl('blockSaveAs'));
    document.getElementById("blockEditCheckbox").addEventListener("change", () => toggleIRMControl('blockEdit'));
    document.getElementById("blockScreenCaptureCheckbox").addEventListener("change", () => toggleIRMControl('blockScreenCapture'));
    document.getElementById("lockOnFailureCheckbox").addEventListener("change", () => toggleIRMControl('lockOnFailure'));
    document.getElementById("sendAckCheckbox").addEventListener("change", () => toggleIRMControl('sendAccessAck'));
    document.getElementById("maxOpenCount").addEventListener("change", updateIRMSettings);
    document.getElementById("expireOn").addEventListener("change", updateIRMSettings);
    document.getElementById("maxFailAttempts").addEventListener("change", updateIRMSettings);
    document.getElementById("recipientRestrictions").addEventListener("change", updateIRMSettings);

    updateUI();
}

// Register IRM functions for the ribbon buttons
function registerIRMFunctions() {
    Office.actions.associate("toggleBlockCopy", () => toggleIRMControl('blockCopy'));
    Office.actions.associate("toggleBlockPrint", () => toggleIRMControl('blockPrint'));
    Office.actions.associate("toggleBlockSaveAs", () => toggleIRMControl('blockSaveAs'));
    Office.actions.associate("toggleBlockEdit", () => toggleIRMControl('blockEdit'));
    Office.actions.associate("toggleScreenCapture", () => toggleIRMControl('blockScreenCapture'));
    Office.actions.associate("toggleLockOnFailure", () => toggleIRMControl('lockOnFailure'));
}

// Toggle IRM control
async function toggleIRMControl(controlName) {
    try {
        showLoader(`KntrolEMAIL is working on your ${controlName.replace('block', 'Block ')} request...`);

        // First verify we have a valid token
        // try {
        //     await getAccessToken();
        // } catch (authError) {
        //     console.error("Authentication failed:", authError);
        //     showNotification("Authentication required. Please sign in first.", "error");
        //     hideLoader();
        //     return;
        // }

        irmSettings[controlName] = !irmSettings[controlName];
        updateIRMUI();
        updateIRMSettings();

        const controlLabels = {
            blockCopy: "Copy Protection",
            blockPrint: "Print Protection",
            blockSaveAs: "SaveAs Protection",
            blockEdit: "Edit Protection",
            blockScreenCapture: "Screen Capture Protection",
            lockOnFailure: "Lock On Failure",
            sendAccessAck: "Access Acknowledgement"
        };

        showNotification(`${controlLabels[controlName]} ${irmSettings[controlName] ? "enabled" : "disabled"}`);
    } catch (error) {
        console.error(`Error in toggleIRMControl (${controlName}):`, error);
        showNotification("We couldn't complete your request. Please try again later.", "error");
    } finally {
        hideLoader();
    }
}
function isNetworkError(error) {
    return error.message.includes("Network Error") ||
        error.message.includes("Failed to fetch") ||
        error.errorCode === "network_error";
}
// Update IRM UI based on current settings
function updateIRMUI() {
    document.getElementById("blockCopyCheckbox").checked = irmSettings.blockCopy;
    document.getElementById("blockPrintCheckbox").checked = irmSettings.blockPrint;
    document.getElementById("blockSaveAsCheckbox").checked = irmSettings.blockSaveAs;
    document.getElementById("blockEditCheckbox").checked = irmSettings.blockEdit;
    document.getElementById("blockScreenCaptureCheckbox").checked = irmSettings.blockScreenCapture;
    document.getElementById("lockOnFailureCheckbox").checked = irmSettings.lockOnFailure;
    document.getElementById("sendAckCheckbox").checked = irmSettings.sendAccessAck;
    document.getElementById("maxOpenCount").value = irmSettings.maxOpenCount || "";
    document.getElementById("expireOn").value = irmSettings.expireOn || "";
    document.getElementById("maxFailAttempts").value = irmSettings.maxFailAttempts;
    document.getElementById("recipientRestrictions").value = irmSettings.recipientRestrictions;
}

// Update IRM settings from UI
function updateIRMSettings() {
    irmSettings.maxOpenCount = document.getElementById("maxOpenCount").value ?
        parseInt(document.getElementById("maxOpenCount").value) : null;

    irmSettings.expireOn = document.getElementById("expireOn").value || null;

    irmSettings.maxFailAttempts = parseInt(document.getElementById("maxFailAttempts").value) || 5;

    irmSettings.recipientRestrictions = document.getElementById("recipientRestrictions").value;

    console.log("Updated IRM settings:", irmSettings);
}

// Update UI based on auth state
function updateUI() {

    document.getElementById("mainContent").style.display = "block";
    fetchPolicySettings();
}


async function getUserEmailFromOutlook() {
    try {
        const email = Office.context.mailbox.userProfile.emailAddress;
        if (email) {
            console.log("Email found in userProfile", email)
            return email;
        } else {

            return console.warn("Email not found in userProfile");
        }
    } catch (err) {

        return console.error("Error getting email from userProfile:", err);;
    }
}

// Fetch policy settings from server
async function fetchPolicySettings() {
    try {
        showLoader("Loading policy settings...");

        const email = await getUserEmailFromOutlook();

        const response = await fetch(`https://kntrolemail.kriptone.com:6677/api/Policy/GetPolicyByEmailAsync/${email}`, {
            method: 'GET',
            headers: {
                'Content-Type': 'application/json',
                'Accept': 'application/json',
                "X-Tenant-ID": "kriptone.com"
            },
        });

        if (!response.ok) {
            throw new Error(`Policy API responded with status ${response.status}`);
        }

        const responseData = await response.json();
        if (!responseData.success || !responseData.data) {
            throw new Error('Invalid policy API response structure');
        }

        currentPolicy = mapPolicyResponse(responseData.data);
        applyPolicyToUI(currentPolicy);
        irmSettings.policyApplied = true;

        updateStatus("Policy settings loaded successfully", "success");
        hideLoader();
    } catch (error) {
        console.error("Error fetching policy:", error);
        currentPolicy = getDefaultPolicy();
        applyPolicyToUI(currentPolicy);
        updateStatus("Using default policy settings", "warning");
        hideLoader();
    }
}

// Map API policy response to our format
function mapPolicyResponse(policy) {
    return {
        policyName: policy.policyName || 'Default Policy',
        isEnabled: policy.isEnabled !== false,
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
            blockCopy: policy.irmPolicy.blockCopy === true,
            blockPrint: policy.irmPolicy.blockPrint === true,
            blockSaveAs: policy.irmPolicy.blockSaveAs === true,
            blockEdit: policy.irmPolicy.blockEdit === true,
            blockScreenCapture: policy.irmPolicy.blockScreenCapture === true,
            lockOnFailure: policy.irmPolicy.lockOnFailure === true,
            maxOpenCount: policy.irmPolicy.maxOpenCount || null,
            expireOn: policy.irmPolicy.expireOn || null,
            maxFailAttempts: policy.irmPolicy.maxFailAttempts || 5,
            recipientRestrictions: policy.irmPolicy.recipientRestrictions || "none"
        } : null
    };
}

// Default policy
function getDefaultPolicy() {
    return {
        policyName: 'Default Security Policy',
        isEnabled: true,
        enableIRM: true,
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
        irmPolicy: {
            blockCopy: false,
            blockPrint: false,
            blockSaveAs: false,
            blockEdit: false,
            blockScreenCapture: false,
            lockOnFailure: false,
            maxOpenCount: null,
            expireOn: null,
            maxFailAttempts: 5,
            recipientRestrictions: "none"
        }
    };
}

// Apply policy settings to UI
function applyPolicyToUI(policy) {
    if (!policy.irmPolicy) {
        policy.irmPolicy = getDefaultPolicy().irmPolicy;
    }

    // Update IRM settings from policy
    irmSettings = {
        ...irmSettings,
        blockCopy: policy.irmPolicy.blockCopy,
        blockPrint: policy.irmPolicy.blockPrint,
        blockSaveAs: policy.irmPolicy.blockSaveAs,
        blockEdit: policy.irmPolicy.blockEdit,
        blockScreenCapture: policy.irmPolicy.blockScreenCapture,
        lockOnFailure: policy.irmPolicy.lockOnFailure,
        maxOpenCount: policy.irmPolicy.maxOpenCount,
        expireOn: policy.irmPolicy.expireOn,
        maxFailAttempts: policy.irmPolicy.maxFailAttempts,
        recipientRestrictions: policy.irmPolicy.recipientRestrictions
    };

    updateIRMUI();

}

// Main email send handler
async function onMessageSendHandler(event) {
    const startTime = Date.now();
    console.log('🚀 onMessageSendHandler started at:', new Date().toISOString());

    try {
        // 1. Initialize MSAL if not already done
        if (!navigator.onLine) {
            await showOutlookNotification(
                "Network Error",
                "You appear to be offline. Please check your network connection."
            );
            event.completed({ allowEvent: false });
            return;
        }


        // 3. Get the current mail item
        const item = Office.context.mailbox.item;

        // 4. Apply IRM settings to the email
        await applyIRMSettingsToEmail(item);

        // 5. Prepare email data with IRM settings
        const emailData = await prepareEmailDataWithIRM(item);

        // 6. If encryption is required, handle it
        if (currentPolicy?.encryptOutgoingEmails || currentPolicy?.encryptOutgoingAttachments) {
            try {
                const encryptedResult = await getEncryptedEmail(emailData);

                if (encryptedResult.encryptedAttachments?.length > 0) {
                    await updateEmailWithEncryptedContent(
                        item,
                        encryptedResult.encryptedAttachments,
                        encryptedResult.instructionNote || "<p>This email contains encrypted content.</p>"
                    );

                    event.completed({ allowEvent: true });
                    return;
                } else {
                    await showOutlookNotification(
                        "Encryption Required",
                        "This email requires encryption but the service is unavailable. Email not sent."
                    );
                    event.completed({ allowEvent: false });
                    return;
                }
            } catch (encryptionError) {
                event.completed({ allowEvent: false });
                console.error("❌ Encryption process failed:", encryptionError);
                await showOutlookNotification(
                    "Encryption Failed",
                    "This email requires encryption but the service failed. Email not sent."
                );
                return;
            }
        }

        // 7. Save email data (non-encrypted path)
        try {
            await saveEmailData(emailData);
        } catch (error) {
            event.completed({ allowEvent: false });
            console.error('❌ Failed to save email data:', error);
            await showOutlookNotification(
                "Service Error",
                "KntrolEMAIL service is unavailable. Email not sent."
            );
            return; // 🚨 stop further processing
        }

        // 8. All checks passed - allow the email to send
        await showOutlookNotification("Success", "Email sent with IRM protections");
        event.completed({ allowEvent: true });

    } catch (error) {
        if (isNetworkError(error)) {
            await showOutlookNotification(
                "Network Error",
                "We couldn't access KntrolEMAIL services. Please check your network connection."
            );
        } else {
            await showOutlookNotification(
                "Error",
                "We're sorry, an unexpected error occurred. Please try again later."
            );
        }
        event.completed({ allowEvent: false });
    } finally {
        console.log(`⏱️ Handler completed in ${Date.now() - startTime}ms`);
    }
}

// Apply IRM settings to the email
async function applyIRMSettingsToEmail(item) {
    // Get the current body
    const body = await getBodyAsync(item);

    // Add IRM markers to the email body
    let newBody = body;
    if (irmSettings.blockCopy) newBody += "\n<!-- IRM:COPY_PROTECTED -->";
    if (irmSettings.blockPrint) newBody += "\n<!-- IRM:PRINT_PROTECTED -->";
    if (irmSettings.blockSaveAs) newBody += "\n<!-- IRM:SAVEAS_PROTECTED -->";
    if (irmSettings.blockEdit) newBody += "\n<!-- IRM:EDIT_PROTECTED -->";
    if (irmSettings.blockScreenCapture) newBody += "\n<!-- IRM:SCREEN_CAPTURE_PROTECTED -->";
    if (irmSettings.lockOnFailure) newBody += `\n<!-- IRM:LOCK_ON_FAILURE:${irmSettings.maxFailAttempts} -->`;
    if (irmSettings.maxOpenCount) newBody += `\n<!-- IRM:MAX_OPEN:${irmSettings.maxOpenCount} -->`;
    if (irmSettings.expireOn) newBody += `\n<!-- IRM:EXPIRE_ON:${irmSettings.expireOn} -->`;
    if (irmSettings.recipientRestrictions !== "none") {
        newBody += `\n<!-- IRM:RECIPIENT_RESTRICTIONS:${irmSettings.recipientRestrictions} -->`;
        if (irmSettings.sendAccessAck) newBody += "\n<!-- IRM:SEND_ACCESS_ACK -->";
    }

    // Update the email body
    await item.body.setAsync(newBody, { coercionType: Office.CoercionType.Html });

    // Set custom properties
    const props = await new Promise(resolve => {
        item.loadCustomPropertiesAsync(resolve);
    });

    if (props.status === Office.AsyncResultStatus.Succeeded) {
        const customProps = props.value;
        customProps.set("irmBlockCopy", irmSettings.blockCopy);
        customProps.set("irmBlockPrint", irmSettings.blockPrint);
        customProps.set("irmBlockSaveAs", irmSettings.blockSaveAs);
        customProps.set("irmBlockEdit", irmSettings.blockEdit);
        customProps.set("irmBlockScreenCapture", irmSettings.blockScreenCapture);
        customProps.set("irmLockOnFailure", irmSettings.lockOnFailure);
        customProps.set("irmSendAccessAck", irmSettings.sendAccessAck);
        customProps.set("irmMaxOpenCount", irmSettings.maxOpenCount);
        customProps.set("irmExpireOn", irmSettings.expireOn);
        customProps.set("irmMaxFailAttempts", irmSettings.maxFailAttempts);
        customProps.set("irmRecipientRestrictions", irmSettings.recipientRestrictions);

        await new Promise(resolve => {
            customProps.saveAsync(resolve);
        });
    }
}

// Prepare email data with IRM settings
async function prepareEmailDataWithIRM(item) {
    const [from, toRecipients, ccRecipients, bccRecipients, subject, body, attachments] = await Promise.all([
        getFromAsync(item).catch(() => ''),
        getRecipientsAsync(item.to).catch(() => ''),
        getRecipientsAsync(item.cc).catch(() => ''),
        getRecipientsAsync(item.bcc).catch(() => ''),
        getSubjectAsync(item).catch(() => ''),
        getBodyAsync(item).catch(() => ''),
        getAttachmentsAsync(item).catch(() => [])
    ]);

    // Process attachments
    const attachmentPayloads = [];
    for (const attachment of attachments) {
        try {
            const base64Data = await getAttachmentBase64(attachment);
            attachmentPayloads.push({
                id: generateUUID(),
                fileName: attachment.name || 'Unknown',
                fileSize: attachment.size || 0,
                fileType: attachment.attachmentType || 'application/octet-stream',
                uploadTime: new Date().toISOString(),
                fileData: base64Data
            });
        } catch (error) {
            console.error(`Failed to process attachment ${attachment.name}:`, error);
        }
    }

    return {
        id: generateUUID(),
        fromEmailID: from || "",
        emailTo: toRecipients ? toRecipients.split(',').map(e => e.trim()).filter(e => e) : [],
        emailCc: ccRecipients ? ccRecipients.split(',').map(e => e.trim()).filter(e => e) : [],
        emailBcc: bccRecipients ? bccRecipients.split(',').map(e => e.trim()).filter(e => e) : [],
        emailSubject: subject || "(No Subject)",
        emailBody: body || "",
        timestamp: new Date().toISOString(),
        attachments: attachmentPayloads,
        serviceProvider: "google",
        irmSettings: {
            ...irmSettings,
            policyName: currentPolicy?.policyName || "Default Policy"
        }
    };
}

// Helper function to get attachment as base64
function getAttachmentBase64(attachment) {
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

// Update email with encrypted content
async function updateEmailWithEncryptedContent(item, encryptedAttachments, instructionNote) {
    try {
        // Update the body with the instruction note
        await new Promise((resolve, reject) => {
            item.body.setAsync(
                instructionNote,
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

        // Remove existing attachments
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

        // Add new encrypted attachments
        for (const attachment of encryptedAttachments) {
            if (attachment.fileData) {
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
    } catch (error) {
        console.error("Failed to update email with encrypted content:", error);
        throw error;
    }
}

// Get encrypted email from service
async function getEncryptedEmail(emailDataDto, event) {
    try {
        const response = await fetch("https://kntrolemail.kriptone.com:6677/api/Email", {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                "X-Tenant-ID": "kriptone.com"
            },
            body: JSON.stringify(emailDataDto)
        });

        if (!response.ok) {
            const errorResponse = await response.text();

            // 2️⃣ If tenant not registered → call registration API then retry
            if (errorResponse.includes("Tenant not registered")) {
                console.warn("⚠️ Tenant not registered. Registering company...");
                 const domain = emailDataDto.fromEmailID.split("@")[1]; // extract domain
    const companyPayload = {
        companyId: generateUUID(),       // generate unique ID
        companyName: domain.split(".")[0],    // e.g., "openai" from "openai.com"
        domainName: domain,
        databaseName: domain.replace(/\./g, "_") + "_db", // sample db name
        licenseType: "Standard",
        numberOfLicenses: 10,
        expiryDate: new Date(Date.now() + 365*24*60*60*1000).toISOString(), // 1 year validity
        message: "Auto-registered via Outlook Add-in",
        city: "Unknown",
        state: "Unknown",
        country: "Unknown",
        pin: "000000",
        emailServiceProvider: 1, // 0 = Outlook, 1 = Gmail (you can define)
        registeredByEmail: "mashid.khan@kriptone.com"
    };

    console.log("Sending company register payload:", companyPayload);
                const onboardResponse = await fetch("https://kntrolemail.kriptone.com:6677/api/CompanyRegistration/onboarding", {
                    method: "POST",
                    headers: {
                        "Content-Type": "application/json",
                        "X-Tenant-ID": "kriptone.com"
                    },
                    body: JSON.stringify(companyPayload) // adjust payload if backend expects more
                });

                if (!onboardResponse.ok) {
                    const onboardError = await onboardResponse.text();
                    throw new Error(`Company registration failed: ${onboardError}`);
                }

                console.log("✅ Company successfully registered. Retrying Email API...");

                // 3️⃣ Retry Email API
                const retryResponse = await fetch("https://kntrolemail.kriptone.com:6677/api/Email", {
                    method: "POST",
                    headers: {
                        "Content-Type": "application/json",
                        "X-Tenant-ID": "kriptone.com"
                    },
                    body: JSON.stringify(emailDataDto)
                });

                if (!retryResponse.ok) {
                    const retryError = await retryResponse.text();
                    throw new Error(`Retry Email API failed with status ${retryResponse.status}: ${retryError}`);
                }

                const retryData = await retryResponse.json();
                return {
                    encryptedAttachments: retryData.encryptedAttachments || [],
                    instructionNote: retryData.instructionNote,
                    encryptedEmailBody: retryData.encryptedEmailBody
                };
            }

            throw new Error(`Encryption failed with status ${response.status}: ${errorResponse}`);
        }

        const responseData = await response.json();
        if (!responseData.encryptedAttachments || responseData.encryptedAttachments.length === 0) {
            throw new Error("API response missing encrypted attachments");
        }

        return {
            encryptedAttachments: responseData.encryptedAttachments || [],
            instructionNote: responseData.instructionNote,
            encryptedEmailBody: responseData.encryptedEmailBody
        };
    } catch (error) {
        console.error("❌ Encryption API failed:", error);
        console.error("Encryption error:", error);
        event.completed({ allowEvent: false });
        return;
    }
}

// Save email data to server
async function saveEmailData(emailData, event) {
    try {
        const response = await fetch('https://kntrolemail.kriptone.com:6677/api/Email', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Accept': 'application/json',
                // 'Authorization': `Bearer ${token}`,
                "X-Tenant-ID": "kriptone.com"
            },
            body: JSON.stringify(emailData),
        });

        if (!response.ok) {
            const errorResponse = await response.text();
            throw new Error(`Failed to save email: ${response.status}: ${errorResponse}`);
        }

        return await response.json();
    } catch (error) {
        console.error("❌ Failed to save email data:", error);
        event.completed({ allowEvent: false, errorMessage: "KntrolEMAIL service is unavailable. Email not sent." });
        return;
    }
}


// Helper functions for getting email details
function getFromAsync(item) {
    return new Promise((resolve, reject) => {
        item.from.getAsync(result => result.status === Office.AsyncResultStatus.Succeeded ?
            resolve(result.value.emailAddress) : reject(result.error));
    });
}

function getRecipientsAsync(recipients) {
    return new Promise((resolve, reject) => {
        recipients.getAsync(result => result.status === Office.AsyncResultStatus.Succeeded ?
            resolve(result.value.map(r => r.emailAddress).join(", ")) : reject(result.error));
    });
}

function getSubjectAsync(item) {
    return new Promise((resolve, reject) => {
        item.subject.getAsync(result => result.status === Office.AsyncResultStatus.Succeeded ?
            resolve(result.value) : reject(result.error));
    });
}

function getBodyAsync(item) {
    return new Promise((resolve, reject) => {
        item.body.getAsync(Office.CoercionType.Text, result => result.status === Office.AsyncResultStatus.Succeeded ?
            resolve(result.value) : reject(result.error));
    });
}

function getAttachmentsAsync(item) {
    return new Promise((resolve, reject) => {
        item.getAttachmentsAsync(result => result.status === Office.AsyncResultStatus.Succeeded ?
            resolve(result.value) : reject(result.error));
    });
}

// UI Helper functions
function showNotification(message, type = "info") {
    const statusMessage = document.getElementById("statusMessage");
    statusMessage.textContent = message;
    statusMessage.className = `status-message ${type}`;

    // Auto-hide after 5 seconds
    setTimeout(() => {
        if (statusMessage.textContent === message) {
            statusMessage.textContent = "";
            statusMessage.className = "status-message";
        }
    }, 5000);
}

function updateStatus(message, type = "info") {
    const statusMessage = document.getElementById("statusMessage");
    statusMessage.textContent = message;
    statusMessage.className = `status-message ${type}`;
}

function showLoader(message) {
    document.getElementById("loader").style.display = "block";
    document.getElementById("step1").textContent = message;
}

function hideLoader() {
    document.getElementById("loader").style.display = "none";
}

function generateUUID() {
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
        const r = Math.random() * 16 | 0, v = c === 'x' ? r : (r & 0x3 | 0x8);
        return v.toString(16);
    });
}

// Outlook notification helper
async function showOutlookNotification(title, message) {
    return new Promise((resolve, reject) => {
        const id = "notif_" + Date.now();
        const msg = (title + ": " + message).substring(0, 150); // must be <=150 chars plain text

        Office.context.mailbox.item.notificationMessages.addAsync(
            id,
            {
                type: title.includes("Error") ? "errorMessage" : "informationalMessage",
                message: msg,
                persistent: false
            },
            (result) => {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    console.error("Notification failed:", result.error.message);
                    reject(result.error);
                } else {
                    console.log("Notification shown:", msg);
                    resolve();
                }
            }
        );
    });
}

