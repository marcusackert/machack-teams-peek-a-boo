import { getMsal } from "./auth";
import { CLIENT_ID } from "./settings";

const status = document.getElementById("status")!;

async function start() {
  const p = new URLSearchParams(location.search);
  const tenant = p.get("tenant");
  const scopes = (p.get("scopes") ?? "").split(" ").filter(Boolean);
  if (!tenant || !scopes.length) throw new Error("Missing tenant or scopes.");
  if (!CLIENT_ID) throw new Error("VITE_CLIENT_ID is not configured.");

  // auth-end runs in this same window after the redirect and needs to know which tenant to finish.
  sessionStorage.setItem("peekaboo.authTenant", tenant);
  const pca = await getMsal(tenant);
  await pca.loginRedirect({ scopes, prompt: "select_account" });
}

start().catch((e: unknown) => {
  status.textContent = `Could not start sign-in: ${e instanceof Error ? e.message : String(e)}`;
});
