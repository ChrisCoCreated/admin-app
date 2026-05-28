function cleanText(value) {
  return String(value ?? "").trim();
}

function parseEmailList(value) {
  return Array.from(new Set(cleanText(value)
    .split(/[,\n;]/)
    .map((entry) => cleanText(entry).toLowerCase())
    .filter(Boolean)));
}

async function sendGraphMail(graphClient, options = {}) {
  const fromEmail = cleanText(options.fromEmail);
  const recipients = Array.from(new Set((options.to || []).map((entry) => cleanText(entry).toLowerCase()).filter(Boolean)));
  const subject = cleanText(options.subject);
  const html = cleanText(options.html);

  if (!fromEmail) {
    throw new Error("Missing ENQUIRY_REMINDER_FROM_EMAIL.");
  }
  if (!recipients.length) {
    throw new Error("At least one email recipient is required.");
  }
  if (!subject || !html) {
    throw new Error("Email subject and HTML body are required.");
  }

  await graphClient.fetchJson(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(fromEmail)}/sendMail`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      message: {
        subject,
        body: {
          contentType: "HTML",
          content: html,
        },
        toRecipients: recipients.map((address) => ({
          emailAddress: { address },
        })),
      },
      saveToSentItems: false,
    }),
  });

  return {
    from: fromEmail,
    to: recipients,
    subject,
  };
}

module.exports = {
  parseEmailList,
  sendGraphMail,
};
