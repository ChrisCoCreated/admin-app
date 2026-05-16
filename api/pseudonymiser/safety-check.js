const pseudonymiserHandler = require("./index");

module.exports = async function safetyCheckHandler(req, res) {
  req.query = { ...(req.query || {}), action: "safety-check" };
  await pseudonymiserHandler(req, res);
};
