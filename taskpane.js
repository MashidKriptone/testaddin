/* global Office, msal */

Office.onReady((info) => {
    console.log("Office ready");

    if (info.host === Office.HostType.Outlook) {
        // First associate handlers
        Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
        Office.actions.associate("onNewMessageCompose", onNewMessageCompose);

        // Then initialize other components
        initializeMSAL();
        initializeUI();
        registerIRMFunctions();
    }
});


function onNewMessageCompose(event) {
    try {
        console.log("New message compose event triggered");
        Office.context.ui.displayTaskPane();
    } catch (error) {
        console.error("Error opening task pane:", error);
    } finally {
        event.completed();
    }
}

// MSAL Configuration
const msalConfig = {
    auth: {
        clientId: "7b7b9a2e-eff4-4af2-9e37-b0df0821b144",
        authority: "https://login.microsoftonline.com/common",
        redirectUri: "https://mashidkriptone.github.io/testaddin/redirect.html", // Use dynamic origin
    },
    cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: true,
        secureCookies: true
    },
    system: {
        loggerOptions: {
            loggerCallback: (level, message, containsPii) => {
                if (containsPii) return;
                console.log(`MSAL ${level}: ${message}`);
            },
            logLevel: msal.LogLevel.Verbose
        }
    }
};


// MSAL instance and state
let msalInstance;
let isInitialized = false;
let authInProgress = false;

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
            updateStatus("Authentication error. Please try signing in again.", "error");
        });
}

// Handle authentication response
function handleAuthResponse(response) {
    if (response) {
        console.log("Authentication successful:", response);
        updateUI();
        fetchPolicySettings();
        return response;
    }
    return null;
}

// Initialize UI elements
function initializeUI() {
    document.getElementById("signInButton").addEventListener("click", signIn);
    // document.getElementById("signOutButton").addEventListener("click", signOut);

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
        try {
            await getAccessToken();
        } catch (authError) {
            console.error("Authentication failed:", authError);
            showNotification("Authentication required. Please sign in first.", "error");
            hideLoader();
            return;
        }

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
    const accounts = msalInstance?.getAllAccounts() || [];
    const isSignedIn = accounts.length > 0;

    document.getElementById("signInButton").style.display = isSignedIn ? "none" : "block";
    // document.getElementById("signOutButton").style.display = isSignedIn ? "block" : "none";
    document.getElementById("mainContent").style.display = isSignedIn ? "block" : "none";

    if (isSignedIn) {
        console.log("User is signed in as:", accounts[0].username);
        updateStatus(`Signed in as ${accounts[0].username}`, "success");
        fetchPolicySettings();
    } else {
        console.log("User is signed out");
        updateStatus("Please sign in to use KntrolEMAIL", "info");
    }
}

// Sign in function
async function signIn() {
    try {
        showLoader("Signing in...");
        const loginRequest = {
            scopes: ["User.Read", "Mail.Send"],
            prompt: "select_account"
        };

        const loginResponse = await msalInstance.loginPopup(loginRequest);
        console.log("Login successful:", loginResponse);
        updateUI();
        hideLoader();
    } catch (error) {
        console.error("Login error:", error);
        updateStatus("Login failed. Please try again.", "error");
        hideLoader();
    }
}

// Sign out function
// async function signOut() {
//     try {
//         showLoader("Signing out...");
//         const accounts = msalInstance.getAllAccounts();
//         if (accounts.length > 0) {
//             const logoutRequest = {
//                 account: accounts[0],
//                 postLogoutRedirectUri: window.location.origin
//             };
//             await msalInstance.logoutPopup(logoutRequest);
//         }
//         console.log("Logout successful");
//         updateUI();
//         hideLoader();
//     } catch (error) {
//         console.error("Logout error:", error);
//         updateStatus("Logout failed. Please try again.", "error");
//         hideLoader();
//     }
// }

// Get access token
async function getAccessToken() {
    if (!isInitialized) {
        initializeMSAL();
    }

    try {
        const accounts = msalInstance.getAllAccounts();
        if (accounts.length === 0) {
            return await loginAndGetToken();
        }

        const silentRequest = {
            scopes: ["User.Read", "Mail.Send"],
            account: accounts[0],
            forceRefresh: false // Only force refresh when needed
        };

        try {
            const response = await msalInstance.acquireTokenSilent(silentRequest);
            return response.accessToken;
        } catch (silentError) {
            console.log("Silent token acquisition failed, trying popup:", silentError);

            // Add specific error handling
            if (silentError instanceof msal.InteractionRequiredAuthError) {
                return await loginAndGetToken();
            }

            throw silentError;
        }
    } catch (error) {
        console.error("Error in getAccessToken:", error);

        // Show user-friendly error message
        if (error.errorCode === "network_error") {
            showNotification("Network error. Please check your connection and try again.", "error");
        } else if (error.errorCode === "login_required") {
            showNotification("Session expired. Please sign in again.", "warning");
            await signOut();
        } else {
            showNotification("Authentication failed. Please try again.", "error");
        }

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

        const loginResponse = await msalInstance.loginPopup(loginRequest);
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

// Fetch policy settings from server
async function fetchPolicySettings() {
    try {
        showLoader("Loading policy settings...");
        const token = await getAccessToken();
        const user = await getUserDetails(token);
        const email = user.mail || user.userPrincipalName;

        const response = await fetch(`https://kntrolemail.kriptone.com:6677/api/Policy/GetPolicyByEmailAsync/${email}`, {
            method: 'GET',
            headers: {
                'Content-Type': 'application/json',
                'Accept': 'application/json',
                'Authorization': `Bearer ${token}`,
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

    // Update policy status display
    // const policyStatus = document.getElementById("policyStatus");
    // policyStatus.innerHTML = `
    //     <strong>Active Policy:</strong> ${policy.policyName}<br>
    //     <strong>IRM Enabled:</strong> ${policy.enableIRM ? "Yes" : "No"}<br>
    //     <strong>Last Updated:</strong> ${new Date().toLocaleString()}
    // `;
}

// Main email send handler
async function onMessageSendHandler(eventArgs) {
    const startTime = Date.now();
    console.log('🚀 onMessageSendHandler started at:', new Date().toISOString());

    try {
        // 1. Initialize MSAL if not already done
        if (!navigator.onLine) {
            await showOutlookNotification(
                "Network Error",
                "You appear to be offline. Please check your network connection."
            );
            eventArgs.completed({ allowEvent: false });
            return;
        }
        if (!isInitialized) {
            initializeMSAL();
        }

        // 2. Authenticate and get access token
        let token;
        try {
            token = await getAccessToken();
        } catch (authError) {
            console.error('❌ Authentication failed:', authError);
            await showOutlookNotification("Authentication Required", "Please sign in to continue.");
            eventArgs.completed({ allowEvent: false });
            return;
        }

        // 3. Get the current mail item
        const item = Office.context.mailbox.item;

        // 4. Apply IRM settings to the email
        await applyIRMSettingsToEmail(item);

        // 5. Prepare email data with IRM settings
        const emailData = await prepareEmailDataWithIRM(item, token);

        // 6. If encryption is required, handle it
        if (currentPolicy?.encryptOutgoingEmails || currentPolicy?.encryptOutgoingAttachments) {
            try {
                const encryptedResult = await getEncryptedEmail(emailData, token);

                if (encryptedResult.encryptedAttachments?.length > 0) {
                    await updateEmailWithEncryptedContent(
                        item,
                        encryptedResult.encryptedAttachments,
                        encryptedResult.instructionNote || "<p>This email contains encrypted content.</p>"
                    );

                    eventArgs.completed({ allowEvent: true });
                    return;
                } else {
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

        // 7. Save email data (non-encrypted path)
        try {
            await saveEmailData(emailData, token);
        } catch (error) {
            console.error('❌ Failed to save email data:', error);
            await showOutlookNotification("Warning", "Email will be sent but audit logging failed");
        }

        // 8. All checks passed - allow the email to send
        await showOutlookNotification("Success", "Email sent with IRM protections");
        eventArgs.completed({ allowEvent: true });

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
        eventArgs.completed({ allowEvent: false });
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
async function prepareEmailDataWithIRM(item, token) {
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
async function getEncryptedEmail(emailDataDto, token) {
    try {
        const response = await fetch("https://kntrolemail.kriptone.com:6677/api/Email", {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                "Authorization": `Bearer ${token}`,
                "X-Tenant-ID": "kriptone.com"
            },
            body: JSON.stringify(emailDataDto)
        });

        if (!response.ok) {
            const errorResponse = await response.text();
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
        console.error("Encryption error:", error);
        throw error;
    }
}

// Save email data to server
async function saveEmailData(emailData, token) {
    try {
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

        if (!response.ok) {
            const errorResponse = await response.text();
            throw new Error(`Failed to save email: ${response.status}: ${errorResponse}`);
        }

        return await response.json();
    } catch (error) {
        console.error("Error saving email data:", error);
        throw error;
    }
}

// Get user details from Microsoft Graph
async function getUserDetails(accessToken) {
    const response = await fetch('https://graph.microsoft.com/v1.0/me', {
        headers: {
            'Authorization': `Bearer ${accessToken}`
        }
    });

    if (!response.ok) {
        throw new Error(`Graph API request failed with status ${response.status}`);
    }

    return await response.json();
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
    return new Promise((resolve) => {
        // Format the message better
        const formattedMessage = `
            <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;">
                <h3 style="color: #0078d4; margin-bottom: 8px;">${title}</h3>
                <p style="margin-top: 0;">${message}</p>
                ${title.includes("Error") ? '<p style="font-size: smaller;">Please try again or contact support if the problem persists.</p>' : ''}
            </div>
        `;

        Office.context.mailbox.item.notificationMessages.addAsync("notification", {
            type: title.includes("Error") ? "errorMessage" : "informationalMessage",
            message: formattedMessage,
            icon: "icon1",
            persistent: false
        }, resolve);
    });
}
async function refreshTokenIfNeeded() {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length === 0) return false;

    const silentRequest = {
        scopes: ["User.Read", "Mail.Send"],
        account: accounts[0],
        forceRefresh: true
    };

    try {
        const response = await msalInstance.acquireTokenSilent(silentRequest);
        return true;
    } catch (error) {
        console.error("Token refresh failed:", error);
        return false;
    }
}