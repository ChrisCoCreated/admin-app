const MSAL_SOURCES = [
  "https://alcdn.msauth.net/browser/2.39.0/js/msal-browser.min.js",
  "https://cdn.jsdelivr.net/npm/@azure/msal-browser@2.39.0/lib/msal-browser.min.js",
  "https://unpkg.com/@azure/msal-browser@2.39.0/lib/msal-browser.min.js",
];

function loadScript(src) {
  return new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = src;
    script.async = true;
    script.onload = () => resolve();
    script.onerror = () => reject(new Error(`Failed to load ${src}`));
    document.head.appendChild(script);
  });
}

async function ensureMsalLoaded() {
  if (window.msal?.PublicClientApplication) {
    return;
  }

  for (const src of MSAL_SOURCES) {
    try {
      await loadScript(src);
      if (window.msal?.PublicClientApplication) {
        return;
      }
    } catch {
      // Try next source.
    }
  }

  throw new Error("Could not load Microsoft sign-in library. Check your network/firewall and refresh.");
}

export function createAuthController(options) {
  const {
    tenantId,
    clientId,
    authCard,
    mainContainer,
    onSignedIn,
    onSignedOut,
  } = options;

  let msalInstance = null;
  let account = null;
  let signOutWrap = null;

  function shouldPreferRedirectAuth() {
    const ua = navigator.userAgent || "";
    const isAndroid = /Android/i.test(ua);
    const isIos = /iPhone|iPad|iPod/i.test(ua);
    const isInAppBrowser = /; wv\)|\bwv\b|FBAN|FBAV|Instagram|Line\/|LinkedInApp|Teams/i.test(ua);
    return isAndroid || isIos || isInAppBrowser;
  }

  function ensureSignOutControl() {
    if (signOutWrap || !mainContainer) {
      return;
    }

    signOutWrap = document.createElement("div");
    signOutWrap.className = "signout-row";
    signOutWrap.innerHTML = '<button id="signOutBtn" type="button" class="signout-btn">Sign out</button>';
    mainContainer.appendChild(signOutWrap);

    const signOutBtn = signOutWrap.querySelector("#signOutBtn");
    signOutBtn.addEventListener("click", () => {
      signOut().catch((error) => {
        console.error(error);
      });
    });
  }

  function setAuthUi(isSignedIn) {
    if (authCard) {
      authCard.hidden = isSignedIn;
    }

    if (isSignedIn) {
      ensureSignOutControl();
      if (signOutWrap) {
        signOutWrap.hidden = false;
      }
    } else if (signOutWrap) {
      signOutWrap.hidden = true;
    }
  }

  async function init() {
    if (window.location.protocol !== "http:" && window.location.protocol !== "https:") {
      throw new Error("Open this page from a web server URL (for example http://localhost:8081), not file://.");
    }

    await ensureMsalLoaded();

    if (!msalInstance) {
      msalInstance = new window.msal.PublicClientApplication({
        auth: {
          clientId,
          authority: `https://login.microsoftonline.com/${tenantId}`,
          redirectUri: window.location.origin,
        },
        cache: {
          cacheLocation: "localStorage",
          storeAuthStateInCookie: true,
        },
      });
      await msalInstance.initialize();
    }

    const redirectResult = await msalInstance.handleRedirectPromise();
    if (redirectResult?.account) {
      account = redirectResult.account;
      msalInstance.setActiveAccount(account);
    }

    if (!account) {
      const active = msalInstance.getActiveAccount();
      const all = msalInstance.getAllAccounts();
      account = active || all[0] || null;
      if (account) {
        msalInstance.setActiveAccount(account);
      }
    }

    return account;
  }

  async function signIn({ scopes = ["openid", "profile"], prompt = "select_account" } = {}) {
    await init();

    if (!account) {
      if (shouldPreferRedirectAuth()) {
        await msalInstance.loginRedirect({ scopes, prompt });
        return null;
      }

      try {
        const loginResult = await msalInstance.loginPopup({ scopes, prompt });
        account = loginResult.account;
        msalInstance.setActiveAccount(account);
      } catch (error) {
        const code = String(error?.errorCode || error?.code || "");
        const shouldFallbackToRedirect =
          code.includes("popup") ||
          code.includes("monitor_window_timeout") ||
          code.includes("block");

        if (!shouldFallbackToRedirect) {
          throw error;
        }

        await msalInstance.loginRedirect({ scopes, prompt });
        return null;
      }
    }

    setAuthUi(true);
    if (typeof onSignedIn === "function") {
      onSignedIn(account);
    }

    return account;
  }

  async function restoreSession() {
    await init();

    if (account) {
      setAuthUi(true);
      if (typeof onSignedIn === "function") {
        onSignedIn(account);
      }
    } else {
      setAuthUi(false);
    }

    return account;
  }

  async function acquireToken(scopes) {
    await init();

    if (!account) {
      throw new Error("Sign in first.");
    }

    const request = { account, scopes };
    try {
      const tokenResponse = await msalInstance.acquireTokenSilent(request);
      return tokenResponse.accessToken;
    } catch {
      const tokenResponse = await msalInstance.acquireTokenPopup(request);
      return tokenResponse.accessToken;
    }
  }

  async function signOut() {
    await init();
    if (!account) {
      setAuthUi(false);
      return;
    }

    try {
      if (shouldPreferRedirectAuth()) {
        await msalInstance.logoutRedirect({
          account,
          postLogoutRedirectUri: window.location.href,
        });
        return;
      }

      await msalInstance.logoutPopup({
        account,
        postLogoutRedirectUri: window.location.href,
      });
    } finally {
      account = null;
      setAuthUi(false);
      if (typeof onSignedOut === "function") {
        onSignedOut();
      }
    }
  }

  return {
    signIn,
    restoreSession,
    acquireToken,
    signOut,
    getAccount: () => account,
    setAuthUi,
  };
}
