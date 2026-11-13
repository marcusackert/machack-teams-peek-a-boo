import "./styles.css";
import { pages } from "@microsoft/teams-js";
import { BASIC_SCOPES, SignInRequiredError, adminConsentUrl, getToken, signIn, signOut, signedInUser } from "./auth";
import { graphGet, listChannels, listJoinedTeams, type Channel, type Page, type Team, type TokenFn } from "./graph";
import { h } from "./render";
import { CLIENT_ID, DEFAULT_TENANT, loadSelection, pageUrl, saveSelection, selectionToParams, type ChannelSelection } from "./settings";
import { initHost, notifyLoaded } from "./teams";

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
const tenantInput = $<HTMLInputElement>("tenant");
const signInBtn = $<HTMLButtonElement>("signin");
const signOutBtn = $<HTMLButtonElement>("signout");
const whoEl = $("who");
const teamSel = $<HTMLSelectElement>("team");
const channelSel = $<HTMLSelectElement>("channel");
const saveBtn = $<HTMLButtonElement>("save");
const msgEl = $("msg");

let teams: Team[] = [];
let channels: Channel[] = [];
let tenantName: string | undefined;
let isTabConfig = false;
const current = loadSelection();
/** Team/channel to preselect once the lists load (from the saved selection or a pasted link). */
let preselect: { teamId: string; channelId: string } | null = current;

/** Parses "Get link to channel" URLs: /l/channel/{channelId}/{name}?groupId={teamId}&tenantId={tenantId} */
export function parseChannelLink(text: string): { tenant: string; teamId: string; channelId: string } | null {
  try {
    const u = new URL(text.trim());
    const m = u.pathname.match(/\/l\/channel\/([^/]+)/);
    const teamId = u.searchParams.get("groupId");
    const tenant = u.searchParams.get("tenantId");
    if (!m || !teamId || !tenant) return null;
    return { tenant, teamId, channelId: decodeURIComponent(m[1]!) };
  } catch {
    return null;
  }
}

const tenant = () => tenantInput.value.trim();
const token: TokenFn = () => getToken(tenant(), BASIC_SCOPES);

async function main() {
  const host = await initHost();
  if (!host.inTeams) $("apphead").hidden = false;
  isTabConfig = host.context?.page.frameContext === "settings";
  if (isTabConfig) {
    saveBtn.hidden = true;
    pages.config.registerOnSaveHandler((evt) => {
      const sel = selection();
      if (!sel) return evt.notifyFailure("Pick a team and channel first.");
      const params = selectionToParams(sel);
      pages.config
        .setConfig({
          entityId: `peekaboo-${sel.channelId}`,
          contentUrl: pageUrl("index.html", params),
          websiteUrl: pageUrl("index.html", params),
          suggestedDisplayName: sel.channelName ?? "Peek-a-boo",
        })
        .then(() => evt.notifySuccess(), (e: Error) => evt.notifyFailure(e.message));
    });
  }
  notifyLoaded();

  if (!CLIENT_ID) return note("VITE_CLIENT_ID is not configured. See README.", true);
  tenantInput.value = current?.tenant ?? DEFAULT_TENANT;
  signInBtn.addEventListener("click", () => void interactiveSignIn());
  signOutBtn.addEventListener("click", () => {
    signOut(tenant());
    resetPickers();
    updateWho();
  });
  tenantInput.addEventListener("change", () => {
    resetPickers();
    updateWho();
    if (tenant()) void loadTeams(false);
  });
  $<HTMLInputElement>("link").addEventListener("input", (e) => {
    const parsed = parseChannelLink((e.target as HTMLInputElement).value);
    if (!parsed) return note("That doesn't look like a Teams channel link.", true);
    preselect = parsed;
    const changed = tenant().toLowerCase() !== parsed.tenant.toLowerCase();
    tenantInput.value = parsed.tenant;
    note("Link recognized. Sign in to load the channel.");
    if (changed) resetPickers();
    updateWho();
    void loadTeams(false);
  });
  teamSel.addEventListener("change", () => void loadChannelsFor(teamSel.value));
  channelSel.addEventListener("change", validate);
  saveBtn.addEventListener("click", save);
  const copyConsent = $<HTMLButtonElement>("copy-consent");
  copyConsent.addEventListener("click", async () => {
    if (!tenant()) return note("Enter the tenant first.", true);
    const url = adminConsentUrl(tenant());
    await navigator.clipboard.writeText(url).catch(() => prompt("Copy this link:", url));
    copyConsent.textContent = "Copied ✓";
  });

  updateWho();
  if (tenant()) await loadTeams(false);
}

async function interactiveSignIn() {
  if (!tenant()) return note("Enter the other tenant's domain (e.g. contoso.com) or tenant ID first.", true);
  signInBtn.disabled = true;
  note("Waiting for sign-in…");
  try {
    await signIn(tenant(), BASIC_SCOPES);
    note("");
    await loadTeams(true);
  } catch (e) {
    note(`Sign-in failed: ${e instanceof Error ? e.message : String(e)}`, true);
  } finally {
    signInBtn.disabled = false;
    updateWho();
  }
}

async function loadTeams(afterSignIn: boolean) {
  try {
    note("Loading teams…");
    teams = (await listJoinedTeams(token)).sort((a, b) => a.displayName.localeCompare(b.displayName));
    tenantName = await graphGet<Page<{ displayName: string }>>(token, "/organization?$select=displayName")
      .then((p) => p.value[0]?.displayName)
      .catch(() => undefined);
    fill(teamSel, teams, preselect?.teamId, "Select a team");
    if (preselect && teamSel.value !== preselect.teamId) note("You are not a member of the team in that link.", true);
    teamSel.disabled = false;
    if (!teams.length) note("You are not a member of any team in this tenant.", true);
    else if (!preselect || teamSel.value === preselect.teamId) note("");
    if (teamSel.value) await loadChannelsFor(teamSel.value);
    updateWho();
  } catch (e) {
    if (e instanceof SignInRequiredError) {
      note(afterSignIn ? "Signed in, but no token was available. Try again." : "Sign in to list your teams.", afterSignIn);
    } else {
      note(`Could not list teams: ${e instanceof Error ? e.message : String(e)}`, true);
    }
  }
}

async function loadChannelsFor(teamId: string) {
  channelSel.disabled = true;
  fill(channelSel, [], undefined, "Loading…");
  validate();
  if (!teamId) return;
  try {
    channels = (await listChannels(token, teamId)).sort((a, b) =>
      a.displayName === "General" ? -1 : b.displayName === "General" ? 1 : a.displayName.localeCompare(b.displayName),
    );
    fill(channelSel, channels, preselect?.teamId === teamId ? preselect.channelId : undefined, "Select a channel");
    channelSel.disabled = false;
  } catch (e) {
    note(`Could not list channels: ${e instanceof Error ? e.message : String(e)}`, true);
  }
  validate();
}

function fill(sel: HTMLSelectElement, items: { id: string; displayName: string }[], selected: string | undefined, placeholder: string) {
  sel.replaceChildren(h("option", { value: "" }, placeholder), ...items.map((i) => h("option", { value: i.id }, i.displayName)));
  if (selected && items.some((i) => i.id === selected)) sel.value = selected;
}

function resetPickers() {
  teams = [];
  channels = [];
  fill(teamSel, [], undefined, "Sign in first");
  fill(channelSel, [], undefined, "Pick a team first");
  teamSel.disabled = channelSel.disabled = true;
  validate();
}

function selection(): ChannelSelection | null {
  const team = teams.find((t) => t.id === teamSel.value);
  const channel = channels.find((c) => c.id === channelSel.value);
  if (!tenant() || !team || !channel) return null;
  return { tenant: tenant(), tenantName, teamId: team.id, teamName: team.displayName, channelId: channel.id, channelName: channel.displayName };
}

function validate() {
  const ok = !!selection();
  saveBtn.disabled = !ok;
  if (isTabConfig) pages.config.setValidityState(ok);
}

function save() {
  const sel = selection();
  if (!sel) return;
  saveSelection(sel);
  location.href = pageUrl("index.html");
}

function updateWho() {
  const who = tenant() ? signedInUser(tenant()) : null;
  whoEl.textContent = who ? `Signed in as ${who}${tenantName ? ` · ${tenantName}` : ""}` : "Not signed in";
  signOutBtn.hidden = !who;
  signInBtn.textContent = who ? "Switch account" : "Sign in";
}

function note(text: string, error = false) {
  msgEl.textContent = text;
  msgEl.classList.toggle("error", error);
}

void main();
