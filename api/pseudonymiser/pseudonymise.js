const pseudonymiserHandler = require("./index");

module.exports = async function pseudonymiseHandler(req, res) {
  req.query = { ...(req.query || {}), action: "pseudonymise" };
  await pseudonymiserHandler(req, res);
};
