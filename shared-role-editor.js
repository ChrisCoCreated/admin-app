// Shared management UI: keep identical in Admin and Associates.
export function createRoleEditor({ request, ids }) {
  const node = key => document.getElementById(ids[key]);
  let users = [], roles = {}, descriptions = {}, writesEnabled = false, busy = false, session = 0;
  const status = (text, error = false) => {
    node("status").textContent = text;
    node("status").classList.toggle("error", error);
  };
  function updateControls() {
    for (const control of node("form").elements) control.disabled = busy || !writesEnabled;
    const user = users.find(user => user.email === node("email").value.trim().toLowerCase());
    const bootstrap = node("options").querySelector('input[value="superadmin"]');
    if (bootstrap && user?.protected) bootstrap.disabled = true;
  }
  function resetEditor() {
    // Associates has a button named "reset", which shadows form.reset in browsers.
    HTMLFormElement.prototype.reset.call(node("form"));
    node("email").readOnly = false;
    updateControls();
  }
  function edit(user) {
    resetEditor();
    node("email").value = user.email;
    node("email").readOnly = true;
    for (const input of node("options").querySelectorAll("input")) input.checked = user.roles.includes(input.value);
    updateControls();
    node("form").scrollIntoView({ behavior: "smooth", block: "center" });
  }
  function render() {
    node("users").replaceChildren();
    const query = node("search").value.trim().toLowerCase();
    for (const user of users.filter(user => [user.email, ...(user.aliases || [])].some(email => email.includes(query)))) {
      const row = document.createElement(node("users").tagName === "TBODY" ? "tr" : "div");
      row.className = "assignment-row";
      const table = row.tagName === "TR";
      const email = document.createElement(table ? "td" : "div");
      const summary = document.createElement(table ? "td" : "p");
      email.textContent = user.email;
      if (user.aliases?.length) {
        const aliases = document.createElement("small");
        aliases.textContent = `Also signs in as ${user.aliases.join(", ")}`;
        aliases.style.display = "block";
        email.appendChild(aliases);
      }
      summary.textContent = user.roles.map(role => roles[role]).join(" / ") || "No assigned roles — Admin blocked; Associates Care access";
      const action = document.createElement(table ? "td" : "div");
      const button = document.createElement("button");
      button.type = "button";
      button.textContent = "Edit";
      button.disabled = busy || !writesEnabled;
      button.addEventListener("click", () => edit(user));
      action.appendChild(button);
      row.append(email, summary, action);
      node("users").appendChild(row);
    }
    if (!node("users").children.length) {
      if (node("users").tagName === "TBODY") {
        const cell = node("users").insertRow().insertCell(); cell.colSpan = 3; cell.textContent = "No matching assignments.";
      } else node("users").textContent = "No matching assignments.";
    }
  }
  async function load() {
    const version = session;
    const data = await request();
    if (version !== session) return;
    users = data.users; roles = data.roles; descriptions = data.descriptions; writesEnabled = data.writesEnabled === true;
    node("options").replaceChildren();
    for (const [value, label] of Object.entries(roles)) {
      const option = document.createElement("label");
      const input = document.createElement("input");
      input.type = "checkbox"; input.value = value;
      const copy = document.createElement("span");
      const title = document.createElement("strong"); title.textContent = label;
      const detail = document.createElement("small"); detail.textContent = descriptions[value] || "";
      copy.append(title, detail); option.append(input, copy); node("options").append(option);
    }
    resetEditor(); render(); node("content").hidden = false;
    status(writesEnabled ? "Roles apply to both apps. Changes take effect on the next request; refresh pages to update menus." : "Role editing is paused during rollout.");
  }
  node("search").addEventListener("input", render);
  node("reset").addEventListener("click", resetEditor);
  node("form").addEventListener("submit", async event => {
    event.preventDefault();
    if (busy || !writesEnabled) return;
    const version = session;
    const email = node("email").value.trim().toLowerCase();
    const selected = [...node("options").querySelectorAll("input:checked")].map(input => input.value);
    busy = true; updateControls(); render(); status("Saving roles...");
    let saved = false;
    try {
      await request("PUT", { email, roles: selected }); saved = true;
      if (version !== session) return;
      await load(); status(`Roles saved for ${email} in both apps.`);
    } catch (error) {
      if (version !== session) return;
      if ([401, 403].includes(error.status)) node("content").hidden = true;
      status(saved ? `Roles saved for ${email}, but the list could not refresh. Reload to view it.` : error.message, true);
    } finally {
      if (version === session) { busy = false; updateControls(); render(); }
    }
  });
  return { load, status, clear() { session++; busy = false; users = []; node("content").hidden = true; node("users").replaceChildren(); resetEditor(); status("Sign in to manage roles."); } };
}
