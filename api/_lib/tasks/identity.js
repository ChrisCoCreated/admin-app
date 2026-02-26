function resolveUserUpn(claims) {
  const value = String(claims?.preferred_username || claims?.upn || claims?.email || "")
    .trim()
    .toLowerCase();

  if (!value) {
    const error = new Error("Could not resolve user UPN from token claims.");
    error.status = 401;
    error.code = "TOKEN_UPN_MISSING";
    throw error;
  }

  return value;
}

module.exports = {
  resolveUserUpn,
};
