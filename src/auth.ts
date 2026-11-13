import { PublicClientApplication, type AccountInfo } from "@azure/msal-browser";
import { authentication } from "@microsoft/teams-js";
import { CLIENT_ID, pageUrl } from "./settings";
import { initHost } from "./teams";

/** Enough to list the user's teams and channels; user-consentable in most tenants. */
export const BASIC_SCOPES = ["User.Read", "Team.ReadBasic.All", "Channel.ReadBasic.All"];
/** Reading channel messages; needs admin consent in the external tenant. */
export const MESSAGE_SCOPES = ["ChannelMessage.Read.All"];

/** What auth-end.html hands back to the page that started sign-in. */
export interface TokenHandoff {
  accessToken: string;
  expiresOn: string; // ISO
  scopes: string[];
  tenantId: string;
  username: string;
}

export class SignInRequiredError extends Error {
  constructor() {
    super("Sign in to the external tenant to continue.");
  }
}

const msalByTenant = new Map<string, Promise<PublicClientApplication>>();

export function authEndUrl(): string {
  return pageUrl("auth-end.html");
}

/** One MSAL instance per external tenant, authority pinned to that tenant. */
export function getMsal(tenant: string): Promise<PublicClientApplication> {
  const key = tenant.toLowerCase();
  let p = msalByTenant.get(key);
  if (!p) {
    p = (async () => {
      const pca = new PublicClientApplication({
        auth: {
          clientId: CLIENT_ID,
          authority: `https://login.microsoftonline.com/${encodeURIComponent(tenant)}`,
          redirectUri: authEndUrl(),
        },
        cache: { cacheLocation: "localStorage" },
      });
      await pca.initialize();
      return pca;
    })();
    msalByTenant.set(key, p);
  }
  return p;
}

// ---- token cache shared between MSAL and hand-offs from the auth window ----

const tokenKey = (tenant: string) => `peekaboo.tokens.${tenant.toLowerCase()}`;
const accountKey = (tenant: string) => `peekaboo.account.${tenant.toLowerCase()}`;

function readHandoffs(tenant: string): TokenHandoff[] {
  try {
    return JSON.parse(sessionStorage.getItem(tokenKey(tenant)) ?? "[]") as TokenHandoff[];
  } catch {
    return [];
  }
}

function storeHandoff(tenant: string, h: TokenHandoff): void {
  const others = readHandoffs(tenant).filter((t) => t.scopes.join() !== h.scopes.join());
  try {
    sessionStorage.setItem(tokenKey(tenant), JSON.stringify([...others, h]));
    localStorage.setItem(accountKey(tenant), JSON.stringify({ tenantId: h.tenantId, username: h.username }));
  } catch {
    /* storage blocked */
  }
}

function covers(granted: string[], wanted: string[]): boolean {
  const g = new Set(granted.map((s) => s.toLowerCase().replace(/^https:\/\/graph\.microsoft\.com\//, "")));
  return wanted.every((w) => g.has(w.toLowerCase()));
}

function knownAccount(tenant: string): { tenantId: string; username: string } | null {
  try {
    return JSON.parse(localStorage.getItem(accountKey(tenant)) ?? "null");
  } catch {
    return null;
  }
}

export function signedInUser(tenant: string): string | null {
  return knownAccount(tenant)?.username ?? null;
}

function pickAccount(pca: PublicClientApplication, tenant: string): AccountInfo | null {
  const known = knownAccount(tenant);
  const all = pca.getAllAccounts();
  return (
    (known && all.find((a) => a.tenantId === known.tenantId && a.username === known.username)) ||
    (known && all.find((a) => a.tenantId === known.tenantId)) ||
    all.find((a) => a.tenantId.toLowerCase() === tenant.toLowerCase()) ||
    null
  );
}

/** Silent token for the external tenant, or SignInRequiredError. */
export async function getToken(tenant: string, scopes: string[]): Promise<string> {
  const fresh = readHandoffs(tenant).find(
    (h) => covers(h.scopes, scopes) && Date.parse(h.expiresOn) - Date.now() > 2 * 60_000,
  );
  if (fresh) return fresh.accessToken;

  const pca = await getMsal(tenant);
  const account = pickAccount(pca, tenant);
  if (!account) throw new SignInRequiredError();
  try {
    const r = await pca.acquireTokenSilent({ scopes, account });
    return r.accessToken;
  } catch {
    throw new SignInRequiredError();
  }
}

/**
 * Interactive sign-in against the external tenant.
 * Teams: host-managed auth window (iframes can't show the Entra sign-in page).
 * Browser: a plain popup running the same auth-start/auth-end pages.
 */
export async function signIn(tenant: string, scopes: string[]): Promise<TokenHandoff> {
  const url = pageUrl("auth-start.html", new URLSearchParams({ tenant, scopes: scopes.join(" ") }));
  const { inTeams } = await initHost();
  const raw = inTeams
    ? await authentication.authenticate({ url, width: 600, height: 640 })
    : await browserPopup(url);
  const handoff = JSON.parse(raw) as TokenHandoff;
  storeHandoff(tenant, handoff);
  return handoff;
}

function browserPopup(url: string): Promise<string> {
  return new Promise((resolve, reject) => {
    const w = window.open(url, "peekaboo-auth", "width=600,height=640");
    if (!w) return reject(new Error("Popup blocked. Allow popups for this site and try again."));
    const onMsg = (e: MessageEvent) => {
      if (e.origin !== location.origin || e.data?.source !== "peekaboo-auth") return;
      cleanup();
      if (e.data.ok) resolve(e.data.result as string);
      else reject(new Error(e.data.error as string));
    };
    const timer = setInterval(() => {
      if (w.closed) {
        cleanup();
        reject(new Error("Sign-in window was closed."));
      }
    }, 500);
    const cleanup = () => {
      clearInterval(timer);
      window.removeEventListener("message", onMsg);
    };
    window.addEventListener("message", onMsg);
  });
}

export function adminConsentUrl(tenant: string): string {
  const p = new URLSearchParams({
    client_id: CLIENT_ID,
    scope: "https://graph.microsoft.com/.default",
    redirect_uri: authEndUrl(),
  });
  return `https://login.microsoftonline.com/${encodeURIComponent(tenant)}/v2.0/adminconsent?${p}`;
}

export function signOut(tenant: string): void {
  try {
    sessionStorage.removeItem(tokenKey(tenant));
    localStorage.removeItem(accountKey(tenant));
  } catch {
    /* ignore */
  }
  void getMsal(tenant).then((pca) => pca.clearCache());
}
