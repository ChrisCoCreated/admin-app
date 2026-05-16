const pseudonymiserHandler = require("./index");

module.exports = async function pseudonymiserHealthHandler(req, res) {
  req.query = { ...(req.query || {}), action: "health" };
  await pseudonymiserHandler(req, res);
};
