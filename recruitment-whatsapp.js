function cleanText(value) {
  return String(value || "").trim();
}

export function normalizeRecruitmentPhoneForActions(phoneNumber) {
  const raw = cleanText(phoneNumber);
  if (!raw) {
    return "";
  }
  let digits = raw.replace(/[^\d+]/g, "");
  if (digits.startsWith("00")) {
    digits = `+${digits.slice(2)}`;
  }
  if (!digits.startsWith("+")) {
    const numeric = digits.replace(/\D/g, "");
    if (numeric.startsWith("0")) {
      digits = `+44${numeric.slice(1)}`;
    } else if (numeric.startsWith("44")) {
      digits = `+${numeric}`;
    } else {
      digits = `+${numeric}`;
    }
  }
  return digits.replace(/(?!^\+)\D/g, "");
}

export function getRecruitmentWhatsAppMessage(candidateName) {
  const firstName = cleanText(candidateName).split(/\s+/)[0] || "there";
  return `Hi ${firstName}, thanks for applying to Thrive Homecare. Are you available today for a quick initial chat? Chris`;
}

export function getRecruitmentWhatsAppUrl(phoneNumber, candidateName) {
  const normalized = normalizeRecruitmentPhoneForActions(phoneNumber).replace(/\D/g, "");
  if (!normalized) {
    return "";
  }
  const url = new URL("https://web.whatsapp.com/send");
  url.searchParams.set("phone", normalized);
  url.searchParams.set("text", getRecruitmentWhatsAppMessage(candidateName));
  return url.toString();
}
