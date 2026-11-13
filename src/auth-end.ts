import { authentication } from "@microsoft/teams-js";
import { getMsal, type TokenHandoff } from "./auth";
import { initHost } from "./teams";

const status = document.getElementById("status")!;

/** Present when Teams opened sign-in in the external browser. */
function externalReturn(): { method: string; authId: string } | null {
  try {
    return JSON.parse(sessionStorage.getItem("peekaboo.authReturn") ?? "null");
  } catch {
    return null;
  }
}

async function finish(): Promise<string> {
  const p = new URLSearchParams(location.search);
  // Landing here from the admin consent link, not from a sign-in.
  if (p.has("admin_consent")) {
    status.textContent =
      p.get("admin_consent") === "True"
        ? "Admin consent granted. You can close this window and sign in from the app."
        : `Admin consent was not granted: ${p.get("error_description") ?? "unknown error"}`;
    return "";
  }

  const tenant = sessionStorage.getItem("peekaboo.authTenant");
  if (!tenant) throw new Error("Sign-in state was lost. Close this tab and try again.");
  const pca = await getMsal(tenant);
  const r = await pca.handleRedirectPromise();
  if (!r) throw new Error("No sign-in response found.");
  sessionStorage.removeItem("peekaboo.authTenant");

  const handoff: TokenHandoff = {
    accessToken: r.accessToken,
    expiresOn: (r.expiresOn ?? new Date(Date.now() + 50 * 60_000)).toISOString(),
    scopes: r.scopes,
    tenantId: r.tenantId,
    username: r.account.username,
  };
  return JSON.stringify(handoff);
}

async function report(ok: boolean, payload: string) {
  if (!payload && ok) return; // admin consent page: nothing to hand back
  const result = ok ? payload : JSON.stringify({ error: payload });

  // External browser (Teams desktop/mobile): hand the result back to Teams through its auth-callback deep link.
  const ext = externalReturn();
  if (ext && ext.method === "deeplink") {
    sessionStorage.removeItem("peekaboo.authReturn");
    const q = new URLSearchParams({ authId: ext.authId, method: "deeplink", result });
    location.href = `msteams://teams.microsoft.com/l/auth-callback?${q}`;
    status.textContent = ok
      ? "Signed in. Returning to Teams… You can close this tab."
      : `Sign-in failed: ${payload}. Returning to Teams…`;
    return;
  }

  const { inTeams } = await initHost();
  if (inTeams) {
    authentication.notifySuccess(result);
    return;
  }
  if (window.opener) {
    window.opener.postMessage({ source: "peekaboo-auth", ok: true, result }, location.origin);
    window.close();
    status.textContent = "Signed in. You can close this tab.";
    return;
  }
  status.textContent = ok ? "Signed in. You can close this tab." : `Sign-in failed: ${payload}`;
}

finish()
  .then((r) => report(true, r))
  .catch((e: unknown) => {
    const msg = e instanceof Error ? e.message : String(e);
    status.textContent = `Sign-in failed: ${msg}`;
    return report(false, msg);
  });
