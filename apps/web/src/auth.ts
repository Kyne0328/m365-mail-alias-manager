import {
  PublicClientApplication,
  type AccountInfo,
  type AuthenticationResult,
  type Configuration
} from "@azure/msal-browser";

const clientId = import.meta.env.VITE_ENTRA_CLIENT_ID;
export const apiScope =
  import.meta.env.VITE_API_SCOPE || (clientId ? `api://${clientId}/access_as_user` : "");

const config: Configuration = {
  auth: {
    clientId,
    authority: "https://login.microsoftonline.com/organizations",
    redirectUri: window.location.origin,
    postLogoutRedirectUri: window.location.origin
  },
  cache: {
    cacheLocation: "sessionStorage"
  }
};

export const msal = new PublicClientApplication(config);

export async function initializeAuth() {
  await msal.initialize();
  const redirectResult = await msal.handleRedirectPromise();

  if (redirectResult?.account) {
    msal.setActiveAccount(redirectResult.account);
    return;
  }

  const existing = msal.getAllAccounts()[0];
  if (existing) {
    msal.setActiveAccount(existing);
  }
}

export function getActiveAccount(): AccountInfo | null {
  return msal.getActiveAccount() ?? msal.getAllAccounts()[0] ?? null;
}

export async function signIn() {
  await msal.loginRedirect({
    scopes: [apiScope],
    prompt: "select_account"
  });
}

export async function signOut() {
  await msal.logoutRedirect({
    account: getActiveAccount() ?? undefined
  });
}

export async function acquireApiToken(): Promise<AuthenticationResult> {
  const account = getActiveAccount();
  if (!account) {
    throw new Error("Sign in is required.");
  }

  return msal.acquireTokenSilent({
    account,
    scopes: [apiScope]
  });
}
