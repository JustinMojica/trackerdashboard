import {
  InteractionRequiredAuthError,
  PublicClientApplication,
  type AccountInfo,
  type Configuration,
} from "@azure/msal-browser";

export type MicrosoftAuthConfig = {
  clientId: string;
  tenantId: string;
};

export type MicrosoftAuthAccount = {
  homeAccountId: string;
  name: string;
  username: string;
};

export type MicrosoftGraphAuthSession = {
  accessToken: string;
  account: MicrosoftAuthAccount;
};

const graphScopes = ["User.Read", "Sites.ReadWrite.All"];
const defaultTenant = "organizations";
let cachedApplication: PublicClientApplication | null = null;
let cachedApplicationKey = "";

export function sanitizeMicrosoftAuthConfig(
  config: MicrosoftAuthConfig,
): MicrosoftAuthConfig {
  return {
    clientId: config.clientId.trim(),
    tenantId: config.tenantId.trim(),
  };
}

export function hasMicrosoftAuthConfig(config: MicrosoftAuthConfig) {
  return Boolean(sanitizeMicrosoftAuthConfig(config).clientId);
}

export function microsoftAuthScopeLabel() {
  return graphScopes.join(", ");
}

export async function signInWithMicrosoft(
  config: MicrosoftAuthConfig,
): Promise<MicrosoftGraphAuthSession> {
  const application = await microsoftApplication(config);
  const result = await application.loginPopup({
    prompt: "select_account",
    scopes: graphScopes,
  });
  if (!result.account) {
    throw new Error("Microsoft sign-in did not return an account.");
  }
  application.setActiveAccount(result.account);
  return {
    accessToken: result.accessToken,
    account: toMicrosoftAuthAccount(result.account),
  };
}

export async function restoreMicrosoftGraphSession(
  config: MicrosoftAuthConfig,
): Promise<MicrosoftGraphAuthSession | null> {
  const application = await microsoftApplication(config);
  const account = application.getActiveAccount() ?? application.getAllAccounts()[0];
  if (!account) return null;
  application.setActiveAccount(account);
  try {
    const result = await application.acquireTokenSilent({
      account,
      scopes: graphScopes,
    });
    return {
      accessToken: result.accessToken,
      account: toMicrosoftAuthAccount(account),
    };
  } catch {
    return {
      accessToken: "",
      account: toMicrosoftAuthAccount(account),
    };
  }
}

export async function refreshMicrosoftGraphToken(
  config: MicrosoftAuthConfig,
): Promise<MicrosoftGraphAuthSession> {
  const application = await microsoftApplication(config);
  const account = application.getActiveAccount() ?? application.getAllAccounts()[0];
  if (!account) {
    return signInWithMicrosoft(config);
  }
  application.setActiveAccount(account);
  try {
    const result = await application.acquireTokenSilent({
      account,
      scopes: graphScopes,
    });
    return {
      accessToken: result.accessToken,
      account: toMicrosoftAuthAccount(account),
    };
  } catch (error) {
    if (error instanceof InteractionRequiredAuthError) {
      const result = await application.acquireTokenPopup({
        account,
        scopes: graphScopes,
      });
      return {
        accessToken: result.accessToken,
        account: toMicrosoftAuthAccount(result.account ?? account),
      };
    }
    throw error;
  }
}

export async function signOutOfMicrosoft(config: MicrosoftAuthConfig) {
  const application = await microsoftApplication(config);
  const account = application.getActiveAccount() ?? application.getAllAccounts()[0];
  if (!account) return;
  await application.logoutPopup({
    account,
    mainWindowRedirectUri: window.location.href,
  });
}

async function microsoftApplication(config: MicrosoftAuthConfig) {
  const cleanConfig = sanitizeMicrosoftAuthConfig(config);
  if (!cleanConfig.clientId) {
    throw new Error("Microsoft Entra application client ID is required.");
  }
  const key = `${cleanConfig.clientId}|${cleanConfig.tenantId || defaultTenant}`;
  if (!cachedApplication || cachedApplicationKey !== key) {
    cachedApplication = new PublicClientApplication(
      microsoftApplicationConfig(cleanConfig),
    );
    cachedApplicationKey = key;
    await cachedApplication.initialize();
  }
  return cachedApplication;
}

function microsoftApplicationConfig(config: MicrosoftAuthConfig): Configuration {
  const tenant = config.tenantId || defaultTenant;
  return {
    auth: {
      clientId: config.clientId,
      authority: `https://login.microsoftonline.com/${tenant}`,
      redirectUri: window.location.origin,
    },
    cache: {
      cacheLocation: "localStorage",
    },
  };
}

function toMicrosoftAuthAccount(account: AccountInfo): MicrosoftAuthAccount {
  return {
    homeAccountId: account.homeAccountId,
    name: account.name ?? account.username,
    username: account.username,
  };
}
