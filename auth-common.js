const MSAL_SOURCES = [
  "https://alcdn.msauth.net/browser/2.39.0/js/msal-browser.min.js",
  "https://cdn.jsdelivr.net/npm/@azure/msal-browser@2.39.0/lib/msal-browser.min.js",
  "https://unpkg.com/@azure/msal-browser@2.39.0/lib/msal-browser.min.js",
];
const FORCE_ACCOUNT_SELECTION_KEY = "thrive.auth.forceAccountSelection";
const AUTO_RESTORE_SUPPRESSED_KEY = "thrive.auth.autoRestoreSuppressed";

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

  function canUseSessionStorage() {
    return typeof window !== "undefined" && !!window.sessionStorage;
  }

  function setForceAccountSelection(enabled) {
    if (!canUseSessionStorage()) {
      return;
    }
    if (enabled) {
      window.sessionStorage.setItem(FORCE_ACCOUNT_SELECTION_KEY, "true");
      return;
    }
    window.sessionStorage.removeItem(FORCE_ACCOUNT_SELECTION_KEY);
  }

  function isAutoRestoreSuppressed() {
    if (!canUseSessionStorage()) {
      return false;
    }
    return window.sessionStorage.getItem(AUTO_RESTORE_SUPPRESSED_KEY) === "true";
  }

  function setAutoRestoreSuppressed(enabled) {
    if (!canUseSessionStorage()) {
      return;
    }
    if (enabled) {
      window.sessionStorage.setItem(AUTO_RESTORE_SUPPRESSED_KEY, "true");
      return;
    }
    window.sessionStorage.removeItem(AUTO_RESTORE_SUPPRESSED_KEY);
  }

  function getCachedAccountCandidate() {
    if (!msalInstance) {
      return null;
    }
    const active = msalInstance.getActiveAccount();
    const all = msalInstance.getAllAccounts();
    return active || all[0] || null;
  }

  function consumeForceAccountSelection() {
    if (!canUseSessionStorage()) {
      return false;
    }
    const enabled = window.sessionStorage.getItem(FORCE_ACCOUNT_SELECTION_KEY) === "true";
    if (enabled) {
      window.sessionStorage.removeItem(FORCE_ACCOUNT_SELECTION_KEY);
    }
    return enabled;
  }

  function getErrorCode(error) {
    return String(error?.errorCode || error?.code || "").toLowerCase();
  }

  function shouldFallbackToRedirectOnInteractiveError(error) {
    const code = getErrorCode(error);
    return (
      shouldPreferRedirectAuth() ||
      code.includes("popup") ||
      code.includes("monitor_window_timeout") ||
      code.includes("block") ||
      code.includes("interaction_required") ||
      code.includes("login_required") ||
      code.includes("consent_required")
    );
  }

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

    const shouldForceAccountSelection = consumeForceAccountSelection();
    const suppressAutoRestore = isAutoRestoreSuppressed();
    const redirectResult = await msalInstance.handleRedirectPromise();
    if (redirectResult?.account) {
      account = redirectResult.account;
      msalInstance.setActiveAccount(account);
      setAutoRestoreSuppressed(false);
      console.info("[Auth] Completed interactive sign-in redirect.", {
        username: redirectResult.account?.username || "",
      });
    }

    if (!account && !shouldForceAccountSelection && !suppressAutoRestore) {
      const active = msalInstance.getActiveAccount();
      const all = msalInstance.getAllAccounts();
      account = active || all[0] || null;
      if (account) {
        msalInstance.setActiveAccount(account);
        console.info("[Auth] Restored cached Microsoft account.", {
          username: account.username || "",
          source: active ? "active_account" : "accounts_cache",
          cachedAccounts: all.length,
        });
      }
    } else if (!account && suppressAutoRestore) {
      const cachedAccount = getCachedAccountCandidate();
      console.info("[Auth] Skipping cached account restore because the user signed out.", {
        username: cachedAccount?.username || "",
      });
    } else if (!account && shouldForceAccountSelection) {
      console.info("[Auth] Skipping cached account restore to force account selection.");
    }

    return account;
  }

  async function signIn({
    scopes = ["openid", "profile"],
    prompt = "select_account",
    forcePrompt = false,
  } = {}) {
    await init();

    if (forcePrompt) {
      account = null;
      if (msalInstance) {
        msalInstance.setActiveAccount(null);
      }
    }

    if (!account || forcePrompt) {
      setAutoRestoreSuppressed(false);
      if (shouldPreferRedirectAuth()) {
        setForceAccountSelection(forcePrompt);
        await msalInstance.loginRedirect({ scopes, prompt });
        return null;
      }

      try {
        const loginResult = await msalInstance.loginPopup({ scopes, prompt });
        account = loginResult.account;
        msalInstance.setActiveAccount(account);
        setAutoRestoreSuppressed(false);
      } catch (error) {
        if (!shouldFallbackToRedirectOnInteractiveError(error)) {
          throw error;
        }

        setForceAccountSelection(forcePrompt);
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
    } catch (silentError) {
      try {
        const tokenResponse = await msalInstance.acquireTokenPopup(request);
        return tokenResponse.accessToken;
      } catch (interactiveError) {
        if (!shouldFallbackToRedirectOnInteractiveError(interactiveError)) {
          throw interactiveError;
        }

        await msalInstance.acquireTokenRedirect(request);
        throw silentError;
      }
    }
  }

  async function acquireSharePointToken(siteHost) {
    return acquireToken([`${siteHost}/AllSites.Write`]);
  }

  async function signOut(options = {}) {
    const redirectUri =
      typeof options.redirectUri === "string" && options.redirectUri.trim()
        ? options.redirectUri.trim()
        : window.location.href;
    const forceAccountSelection = options.forceAccountSelection === true;

    await init();
    setForceAccountSelection(forceAccountSelection);
    setAutoRestoreSuppressed(true);
    if (!account) {
      setAuthUi(false);
      return;
    }

    try {
      if (shouldPreferRedirectAuth()) {
        await msalInstance.logoutRedirect({
          account,
          postLogoutRedirectUri: redirectUri,
        });
        return;
      }

      await msalInstance.logoutPopup({
        account,
        postLogoutRedirectUri: redirectUri,
      });
    } finally {
      account = null;
      if (msalInstance) {
        msalInstance.setActiveAccount(null);
      }
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
    acquireSharePointToken,
    signOut,
    getAccount: () => account,
    getCachedAccount: () => getCachedAccountCandidate(),
    setAuthUi,
    requestAccountSelection: () => setForceAccountSelection(true),
  };
}
