function cleanText(value) {
  return String(value || "").trim();
}

export const INTRO_WHATSAPP_SERVICES = {
  thrive: {
    label: "Thrive Homecare",
    url: "https://www.thrivehomecare.co.uk/",
  },
  mentalCapacity: {
    label: "Plan with Care / Mental Capacity",
    url: "https://www.planwithcare.co.uk/mental-capacity",
  },
};

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

export function getIntroWhatsAppMessage(contactName, serviceKey = "thrive") {
  const firstName = cleanText(contactName).split(/\s+/)[0] || "there";
  const service = INTRO_WHATSAPP_SERVICES[serviceKey] || INTRO_WHATSAPP_SERVICES.thrive;
  if (service === INTRO_WHATSAPP_SERVICES.mentalCapacity) {
    return `Hi ${firstName}, it was good to meet you earlier. As mentioned, Plan with Care supports families and professionals with mental capacity assessments and related planning. Here's the link in case useful: ${service.url}`;
  }
  return `Hi ${firstName}, it was good to meet you earlier. As mentioned, Thrive Homecare provides companionship and care at home, focused on helping people stay independent, connected and living well. Here's the link in case useful: ${service.url}`;
}

export function getWhatsAppUrlForMessage(phoneNumber, message) {
  const normalized = normalizeRecruitmentPhoneForActions(phoneNumber).replace(/\D/g, "");
  if (!normalized) {
    return "";
  }
  const url = new URL("https://web.whatsapp.com/send");
  url.searchParams.set("phone", normalized);
  url.searchParams.set("text", cleanText(message));
  return url.toString();
}

export function getRecruitmentWhatsAppUrl(phoneNumber, candidateName) {
  return getWhatsAppUrlForMessage(phoneNumber, getRecruitmentWhatsAppMessage(candidateName));
}
