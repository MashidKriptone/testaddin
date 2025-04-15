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

// Event handler for the ItemSend event
async function onMessageSendHandler(eventArgs) {
    try {
       if (!isInitialized) {
            initializeMSAL();
        }

        // Get token first
        let token;
        try {
            token = await getAccessToken();
            console.log("Access token:", token);
        } catch (authError) {
            console.error("Authentication failed:", authError);
            showOutlookNotification("Authentication Required", "Please sign in to continue.");
            eventArgs.completed({ allowEvent: false });
            return;
        }
        // Authenticate and get token
        const item = Office.context.mailbox.item;

        // Retrieve email details
        const from = await getFromAsync(item);
        const toRecipients = await getRecipientsAsync(item.to);
        const ccRecipients = await getRecipientsAsync(item.cc);
        const bccRecipients = await getRecipientsAsync(item.bcc);
        const subject = await getSubjectAsync(item);
        const body = await getBodyAsync(item);
        const attachments = await getAttachmentsAsync(item);

        console.log("🔹 Email Details:", { from, toRecipients, ccRecipients, bccRecipients, subject, body, attachments });

        // Fetch policy domains
        const { allowedDomains, blockedDomains, contentScanning, attachmentPolicy, blockedAttachments,encryptOutgoingEmails, encryptOutgoingAttachments } = await fetchPolicyDomains(token);

        console.log("🔹 Policy Check:", { allowedDomains, blockedDomains });


            // **1️⃣ If no domain restrictions exist, allow email to send**
        const noDomainRestrictions = allowedDomains.length === 0 && blockedDomains.length === 0;
        if (noDomainRestrictions) {
            console.log("✅ No domain restrictions. Proceeding...");
        } else {
            // **2️⃣ Check if the email contains blocked domains**
            if (isDomainBlocked(toRecipients, blockedDomains) || 
                isDomainBlocked(ccRecipients, blockedDomains) || 
                isDomainBlocked(bccRecipients, blockedDomains)) {
                showOutlookNotification("Blocked Domain", "KntrolEMAIL detected a blocked domain policy and prevented the email from being sent.");
                eventArgs.completed({ allowEvent: false });
                return;
            }
        }

        if (!toRecipients && !ccRecipients && !bccRecipients) {
            console.warn("❌ No recipients found. Email is not sent.");
            Office.context.mailbox.item.notificationMessages.addAsync("error", {
                type: "errorMessage",
                message: "Please add at least one recipient (To, CC, or BCC).",
            });
            eventArgs.completed({ allowEvent: false });
            return;
        }
        // **3️⃣ Validate email addresses**
        if (!validateEmailAddresses(toRecipients) || 
            !validateEmailAddresses(ccRecipients) || 
            !validateEmailAddresses(bccRecipients)) {
            showOutlookNotification("Invalid Email", "One or more email addresses are invalid.");
            eventArgs.completed({ allowEvent: false });
            return;
        }

        if (contentScanning) {
            for (const pattern in regexPatterns) {
                if (regexPatterns[pattern].test(body)) {
                    showOutlookNotification("Restricted Content", `Your email contains restricted data: ${pattern}.`);
                    eventArgs.completed({ allowEvent: false });
                    return;
                }
            }
        }

        if (attachmentPolicy) {
            for (const attachment of attachments) {
                const ext = attachment.name.substring(attachment.name.lastIndexOf('.') + 1).toLowerCase();
if (blockedAttachments.includes(ext)) {
                    showOutlookNotification("Restricted Attachment", `Attachment \"${attachment.name}\" is restricted.`);
                    eventArgs.completed({ allowEvent: false });
                    return;
                }
            }
        }
        console.log("✅ Passed all policy checks. Saving email data...");

        // **6️⃣ Save email data to API before sending**
        const emailData = prepareEmailData(from, toRecipients, ccRecipients, bccRecipients, subject, body, attachments);
        if (encryptOutgoingEmails || encryptOutgoingAttachments) {
            const encrypted = await getEncryptedEmail(emailData ,token);
            if (encryptOutgoingEmails && encrypted.EncryptedBody) {
                await item.body.setAsync(encrypted.EncryptedBody, { coercionType: Office.CoercionType.Html });
            }
            // Note: Office.js doesn't support modifying attachments during send
        }

            const saveSuccess = await saveEmailData(emailData,token);
            if (saveSuccess.success) {
                console.log("✅ Email data saved. Ensuring email is sent.");
                eventArgs.completed({ allowEvent: true });
            } else {
                console.warn("❌ Email saving failed:", saveSuccess.message);
                showOutlookNotification("Error", saveSuccess.message || "Email saving failed due to a backend error.");
                eventArgs.completed({ allowEvent: false });
            }
        
        
        console.log("✅ Email Passed Validation. Fetching Microsoft Graph Emails...");
        const oldEmail = await fetchEmails(); 
        console.log("Previous Emails ",oldEmail);
        
       

    } catch (error) {
        console.error('❌ Error during send event:', error);
        showOutlookNotification("Error", "An unexpected error occurred while sending the email.");
        eventArgs.completed({ allowEvent: false });
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

// Fetch policy domains from the backend
async function fetchPolicyDomains(token) {
    try {
        const response = await fetch('https://kntrolemail.kriptone.com:6677/api/Policy', {
            method: 'GET',
            headers: { 'Content-Type': 'application/json', 'Accept': 'application/json','Authorization': `Bearer ${token}`,"X-Tenant-ID": "kriptone.com", },
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
            encryptOutgoingEmails:policy.encryptOutgoingEmails || false,
            encryptOutgoingAttachments:policy.encryptOutgoingAttachments || false
        };
    } catch (error) {
        console.error("❌ Error fetching policy:", error);
        return { allowedDomains: [], blockedDomains: [], contentScanning: false, attachmentPolicy: false, blockedAttachments: [] };
    }
}

async function getEncryptedEmail(emailDataDto,token) {
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
        throw new Error("Encryption request failed: " + response.status);
    }

    return await response.json();
}

  function interceptGmailAndEncryptWithPolicy(encryptOutgoingEmails,encryptOutgoingAttachments) {
                const sendButton = document.querySelector("div[role='button'][data-tooltip^='Send']");
                if (!sendButton) return;
            
                sendButton.addEventListener("click", async (e) => {
                    e.preventDefault();
            
                    const bodyElement = document.querySelector("div[aria-label='Message Body']");
                    const toField = document.querySelector("textarea[name='to']");
                    const ccField = document.querySelector("textarea[name='cc']");
                    const bccField = document.querySelector("textarea[name='bcc']");
                    const subjectField = document.querySelector("input[name='subjectbox']");
                    const attachmentBox = document.querySelector(".a1.aaA.aMZ");
            
                    // Gather data
                    const rawBody = bodyElement.innerHTML;
                    const rawAttachments = await getGmailAttachmentData(); // should return: [{ FileName, FileSize, FileType, FileData }]
            
                    const emailDataDto = {
                        FromEmailID: "current_user_email@domain.com", // Replace with actual user
                        EmailTo: toField?.value?.split(",").map(e => e.trim()) || [],
                        EmailCc: ccField?.value?.split(",").map(e => e.trim()) || [],
                        EmailBcc: bccField?.value?.split(",").map(e => e.trim()) || [],
                        EmailSubject: subjectField?.value || "",
                        EmailBody: encryptOutgoingEmails ? rawBody : "", // Empty if not encrypting
                        Attachments: encryptOutgoingAttachments ? rawAttachments : [], // Empty if not encrypting
                    };
            
                    try {
                        const encrypted = await getEncryptedEmail(emailDataDto);
            
                        // Apply encrypted body if policy requires
                        if (encryptOutgoingEmails && encrypted.EncryptedBody) {
                            bodyElement.innerHTML = encrypted.EncryptedBody;
                        }
            
                        // Show encrypted attachment message if policy requires
                        if (encryptOutgoingAttachments && encrypted.Attachments?.length > 0 && attachmentBox) {
                            attachmentBox.innerHTML = "<p>[Encrypted attachments included - use KntrolFILE Reader]</p>";
                        }
            
                        // Now actually trigger send
                        sendButton.click();
            
                    } catch (error) {
                        console.error("❌ Gmail Encryption Error:", error);
                        alert("Encryption failed. Please try again.");
                    }
                });
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

function prepareEmailData(from, to, cc, bcc, subject, body, attachments) {
    let emailId = generateUUID();
    return {
        Id: emailId,
        FromEmailID: from,
        Attachments: attachments.map(attachment => ({
            Id: generateUUID(),
            FileName: attachment.name,
            FileType: attachment.attachmentType,
            FileSize: attachment.size,
            UploadTime: new Date().toISOString(),
        })),
        EmailBcc: bcc ? bcc.split(',').map(email => email.trim()) : [],
        EmailCc: cc ? cc.split(',').map(email => email.trim()) : [],
        EmailBody: body,
        EmailSubject: subject,
        EmailTo: to ? to.split(',').map(email => email.trim()) : [],
        Timestamp: new Date().toISOString(),
    };
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

async function fetchEmails() {
    try {
        const token = await getAccessToken();
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