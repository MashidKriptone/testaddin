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
        const token = getAccessToken();
        console.log("access token ------------",token)

        // Fetch policy domains
        const { allowedDomains, blockedDomains, contentScanning, attachmentPolicy, blockedAttachments } = await fetchPolicyDomains();

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
                if (blockedAttachments.includes(attachment.name.split('.').pop())) {
                    showOutlookNotification("Restricted Attachment", `Attachment \"${attachment.name}\" is restricted.`);
                    eventArgs.completed({ allowEvent: false });
                    return;
                }
            }
        }
        console.log("✅ Passed all policy checks. Saving email data...");

        // **6️⃣ Save email data to API before sending**
        const emailData = prepareEmailData(from, toRecipients, ccRecipients, bccRecipients, subject, body, attachments);
        const saveSuccess = await saveEmailData(emailData);

        console.log("✅ Email Passed Validation. Fetching Microsoft Graph Emails...");
        await fetchEmails(); 
        
        if (saveSuccess.success) {
            console.log("✅ Email data saved. Ensuring email is sent.");
            eventArgs.completed({ allowEvent: true });
        } else {
            console.warn("❌ Email saving failed:", saveSuccess.message);
            showOutlookNotification("Error", saveSuccess.message || "Email saving failed due to a backend error.");
            eventArgs.completed({ allowEvent: false });
        }

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
async function fetchPolicyDomains() {
    try {
        const response = await fetch('https://kntrolemail.kriptone.com:6677/api/Policy', {
            method: 'GET',
            headers: { 'Content-Type': 'application/json', 'Accept': 'application/json' },
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
        };
    } catch (error) {
        console.error("❌ Error fetching policy:", error);
        return { allowedDomains: [], blockedDomains: [], contentScanning: false, attachmentPolicy: false, blockedAttachments: [] };
    }
}

async function saveEmailData(emailData) {
    try {
        const response = await fetch('https://kntrolemail.kriptone.com:6677/api/Email', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json', 'Accept': 'application/json' },
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
let msalInstance;
let currentAccount = null;

function initializeMSAL() {
    const msalConfig = {
        auth: {
            clientId: "7b7b9a2e-eff4-4af2-9e37-b0df0821b144", // 🔁 Replace with your real Client ID
            authority: "https://login.microsoftonline.com/common",
            redirectUri: "https://outlook.live.com/" // 🔁 Update this to match your Azure registration
        }
    };

    msalInstance = new msal.PublicClientApplication(msalConfig);

    const loginRequest = {
        scopes: ["User.Read", "Mail.Send"]
    };

    document.getElementById("signInButton").addEventListener("click", async () => {
        try {
            const loginResponse = await msalInstance.loginPopup(loginRequest);
            console.log("Login successful:", loginResponse);
            currentAccount = loginResponse.account;
            updateUI(true);
        } catch (error) {
            console.error("Login error:", error);
        }
    });

    document.getElementById("signOutButton").addEventListener("click", async () => {
        try {
            const logoutRequest = {
                account: currentAccount
            };
            await msalInstance.logoutPopup(logoutRequest);
            console.log("Logout successful");
            updateUI(false);
        } catch (error) {
            console.error("Logout error:", error);
        }
    });
}

function updateUI(isSignedIn) {
    document.getElementById("signInButton").style.display = isSignedIn ? "none" : "inline-block";
    document.getElementById("signOutButton").style.display = isSignedIn ? "inline-block" : "none";
}

function formatTokenResponse(response) {
    return {
        access_token: response.accessToken,
        id_token: response.idToken,
        expires_in: Math.floor((response.expiresOn.getTime() - Date.now()) / 1000),
        token_type: "Bearer",
        scope: response.scopes.join(" "),
        account: {
            username: response.account.username,
            name: response.account.name
        }
    };
}

async function getAccessToken() {
    const accounts = msalInstance.getAllAccounts();

    if (accounts.length === 0) {
        throw new Error("No signed-in account found");
    }

    const silentRequest = {
        scopes: ["User.Read", "Mail.Send"],
        account: accounts[0]
    };

    try {
        const response = await msalInstance.acquireTokenSilent(silentRequest);
        return response.accessToken;
    } catch (error) {
        console.warn("Silent token acquisition failed, falling back to popup.", error);
        const response = await msalInstance.acquireTokenPopup(silentRequest);
        return response.accessToken;
    }
}

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