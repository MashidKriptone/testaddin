/* global Office */

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        Office.context.mailbox.item.addHandlerAsync(Office.EventType.ItemSend, validateAndTrackEmail);
        console.log("Add-in is running.");
    }
});

// Email regex: validates general email format with 2-3 character domain extensions
const emailRegex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,3}$/;

// Regex patterns for additional checks
const regexPatterns = {
    body: /\b(confidential|prohibited|restricted)\b/i, // Example sensitive keywords in the body
    attachmentName: /\.(exe|bat|sh)$/i, // Example restricted file extensions
};

// Event handler for the ItemSend event
async function validateAndTrackEmail(eventArgs) {
    try {
        const item = Office.context.mailbox.item;

        // Retrieve email details
        const from = await getFromAsync(item);
        const toRecipients = await getRecipientsAsync(item.to);
        const ccRecipients = await getRecipientsAsync(item.cc);
        const bccRecipients = await getRecipientsAsync(item.bcc);
        const subject = await getSubjectAsync(item);
        const body = await getBodyAsync(item);
        const attachments = await getAttachmentsAsync(item);

        // Fetch policy domains
        const { allowedDomains, blockedDomains } = await fetchPolicyDomains();

        // Allow email if allowedDomains or blockedDomains is empty
        if (
            (blockedDomains.length > 0 && isDomainBlocked(toRecipients, blockedDomains)) ||
            isDomainBlocked(ccRecipients, blockedDomains) ||
            isDomainBlocked(bccRecipients, blockedDomains)
        ) {
            Office.context.mailbox.item.notificationMessages.addAsync("error", {
                type: "errorMessage",
                message: "This email cannot be sent as it contains domains that violate the policy.",
            });
            eventArgs.completed({ allowEvent: false });
            return;
        }

        // Validate email addresses
        if (!validateEmailAddresses(toRecipients) ||
            !validateEmailAddresses(ccRecipients) ||
            !validateEmailAddresses(bccRecipients)) {
            Office.context.mailbox.item.notificationMessages.addAsync("error", {
                type: "errorMessage",
                message: "One or more email addresses are invalid.",
            });
            eventArgs.completed({ allowEvent: false });
            return;
        }

        // Validate body content using regex
        if (regexPatterns.body.test(body)) {
            Office.context.mailbox.item.notificationMessages.addAsync("error", {
                type: "errorMessage",
                message: "The email contains prohibited content in the body.",
            });
            eventArgs.completed({ allowEvent: false });
            return;
        }

        // Validate attachments
        for (const attachment of attachments) {
            if (regexPatterns.attachmentName.test(attachment.name)) {
                Office.context.mailbox.item.notificationMessages.addAsync("error", {
                    type: "errorMessage",
                    message: `The attachment "${attachment.name}" has a restricted file type.`,
                });
                eventArgs.completed({ allowEvent: false });
                return;
            }
        }

        // Prepare email data for saving
        const emailData = prepareEmailData(from, toRecipients, ccRecipients, bccRecipients, subject, body, attachments);

        // Save email data to the backend server
        await saveEmailData(emailData);

        // Allow the email to be sent
        eventArgs.completed();
    } catch (error) {
        console.error('Error during send event:', error);
        eventArgs.completed({ allowEvent: false });
    }
}

// Helper function to fetch policy domains from the backend
async function fetchPolicyDomains() {
    try {
        const response = await fetch('https://kntrolemail.kriptone.com:6677/api/Admin/policies', {
            method: 'GET',
            headers: {
                'Content-Type': 'application/json',
                'Accept': 'application/json',
            },
        });

        if (!response.ok) {
            throw new Error('Failed to fetch policy domains: ' + response.statusText);
        }

        const policies = await response.json();
        const allowedDomains = policies[0]?.AllowedDomains || [];
        const blockedDomains = policies[0]?.BlockedDomains || [];

        return { allowedDomains, blockedDomains };
    } catch (error) {
        console.error('Error fetching policy domains:', error);
        return { allowedDomains: [], blockedDomains: [] }; // Default to empty arrays
    }
}

// Helper function to validate email addresses
function validateEmailAddresses(recipients) {
    if (!recipients) return true;

    const emailArray = recipients.split(',').map(email => email.trim());
    for (let email of emailArray) {
        if (!emailRegex.test(email)) {
            console.log(`Invalid email address: ${email}`);
            return false;
        }
    }
    return true;
}

// Helper function to check if domains are blocked
function isDomainBlocked(recipients, blockedDomains) {
    if (!blockedDomains || blockedDomains.length === 0) return false; // Allow by default if no blocked domains

    const recipientArray = recipients ? recipients.split(',').map(email => email.trim()) : [];

    for (let recipient of recipientArray) {
        const domain = recipient.split('@')[1]; // Extract the domain from the email

        if (blockedDomains.includes(domain)) {
            console.log(`Domain ${domain} is blocked.`);
            return true;
        }
    }
    return false;
}

// Helper function to prepare email data
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

// Helper function to save email data to the backend
async function saveEmailData(emailData) {
    const response = await fetch('https://kntrolemail.kriptone.com:6677/api/Email', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'Accept': 'application/json',
        },
        body: JSON.stringify(emailData),
    });

    if (!response.ok) {
        throw new Error('Failed to save email data: ' + response.statusText);
    }

    console.log('Email data saved successfully.');
}

// Helper function to generate a UUID
function generateUUID() {
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
        const r = Math.random() * 16 | 0, v = c === 'x' ? r : (r & 0x3 | 0x8);
        return v.toString(16);
    });
}

// Async functions to retrieve email details
function getFromAsync(item) {
    return new Promise((resolve, reject) => {
        item.from.getAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const fromEmail = result.value.emailAddress || result.value.address;
                resolve(fromEmail);
            } else {
                reject('Error retrieving from address: ' + result.error.message);
            }
        });
    });
}

function getRecipientsAsync(recipients) {
    return new Promise((resolve, reject) => {
        recipients.getAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const emails = result.value.map(recipient => recipient.emailAddress || recipient.address).join(", ");
                resolve(emails);
            } else {
                reject('Error retrieving recipients: ' + result.error.message);
            }
        });
    });
}

function getSubjectAsync(item) {
    return new Promise((resolve, reject) => {
        item.subject.getAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve(result.value);
            } else {
                reject('Error retrieving subject: ' + result.error.message);
            }
        });
    });
}

function getBodyAsync(item) {
    return new Promise((resolve, reject) => {
        item.body.getAsync(Office.CoercionType.Text, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve(result.value);
            } else {
                reject('Error retrieving body: ' + result.error.message);
            }
        });
    });
}

function getAttachmentsAsync(item) {
    return new Promise((resolve, reject) => {
        item.getAttachmentsAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const attachments = result.value;
                resolve(attachments || []);
            } else {
                reject('Error retrieving attachments: ' + result.error.message);
            }
        });
    });
}
