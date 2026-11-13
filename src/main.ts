import "./styles.css";
import { BASIC_SCOPES, MESSAGE_SCOPES, SignInRequiredError, adminConsentUrl, getToken, signIn, signedInUser } from "./auth";
import { GraphError, fetchBlobUrl, getChannel, listMessages, type ChatMessage, type TokenFn } from "./graph";
import { dayLabel, h, lastActivity, renderPost } from "./render";
import { CLIENT_ID, loadSelection, pageUrl, type ChannelSelection } from "./settings";
import { initHost, notifyLoaded } from "./teams";
import { demoMessages, demoSelection, isDemo } from "./demo";

const POLL_MS = 30_000;
const SCOPES = [...BASIC_SCOPES, ...MESSAGE_SCOPES];

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
const feed = $("feed");
const stateBox = $("state");
const liveEl = $("live");
const liveText = $("live-text");
const olderBtn = $<HTMLButtonElement>("older");
const openLink = $<HTMLAnchorElement>("open-in-teams");

const messages = new Map<string, ChatMessage>();
const expanded = new Set<string>();
const blobCache = new Map<string, Promise<string>>();
let olderLink: string | undefined;
let lastRefresh = 0;
let lastSeen = 0;
let pollTimer: number | undefined;
let busy = false;

async function main() {
  const host = await initHost();
  if (!host.inTeams) $("apphead").hidden = false;
  notifyLoaded();

  const settingsHref = pageUrl("config.html", new URLSearchParams(location.search));
  for (const a of document.querySelectorAll<HTMLAnchorElement>("a[data-settings]")) a.href = settingsHref;

  if (!CLIENT_ID && !isDemo) return showState("⚙", "App not configured", "Set VITE_CLIENT_ID in .env and rebuild. See README.");

  const sel = isDemo ? demoSelection : loadSelection();
  if (!sel) {
    return showState("👁", "No channel selected", "Pick a team and channel in the other tenant to start peeking.", [
      linkButton("Choose a channel", settingsHref, true),
    ]);
  }
  renderCrumb(sel);
  const token: TokenFn = () => getToken(sel.tenant, SCOPES);
  const seenKey = `peekaboo.seen.${sel.channelId}`;
  lastSeen = isDemo ? new Date().setHours(11, 0, 0, 0) : Number(localStorage.getItem(seenKey) ?? Date.now());

  $("refresh").addEventListener("click", () => void refresh(sel, token, true));
  olderBtn.addEventListener("click", () => void loadOlder(token));
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible") void refresh(sel, token, false);
  });
  setInterval(updateLiveText, 5000);
  window.addEventListener("pagehide", () => localStorage.setItem(seenKey, String(Date.now())));

  await refresh(sel, token, true);
}

async function refresh(sel: ChannelSelection, token: TokenFn, initial: boolean) {
  if (busy) return;
  busy = true;
  liveEl.classList.add("busy");
  try {
    const page = isDemo ? demoMessages() : await listMessages(token, sel.teamId, sel.channelId);
    const wasNearBottom = feed.scrollHeight - feed.scrollTop - feed.clientHeight < 80;
    const firstLoad = messages.size === 0;
    for (const m of page.value) messages.set(m.id, m);
    if (firstLoad) olderLink = page["@odata.nextLink"];
    hideState();
    renderFeed(token);
    if (firstLoad || wasNearBottom) feed.scrollTop = feed.scrollHeight;
    lastRefresh = Date.now();
    liveEl.classList.remove("error");
    updateLiveText();
    if (initial && !isDemo) void fillChannelInfo(sel, token);
    schedulePoll(sel, token);
  } catch (e) {
    handleError(e, sel, token);
  } finally {
    busy = false;
    liveEl.classList.remove("busy");
  }
}

function schedulePoll(sel: ChannelSelection, token: TokenFn) {
  clearTimeout(pollTimer);
  pollTimer = window.setTimeout(() => {
    if (document.visibilityState === "visible") void refresh(sel, token, false);
    else schedulePoll(sel, token);
  }, POLL_MS);
}

async function loadOlder(token: TokenFn) {
  if (!olderLink) return;
  olderBtn.disabled = true;
  const sel = loadSelection()!;
  try {
    const before = feed.scrollHeight - feed.scrollTop;
    const page = await listMessages(token, sel.teamId, sel.channelId, olderLink);
    for (const m of page.value) messages.set(m.id, m);
    olderLink = page["@odata.nextLink"];
    renderFeed(token);
    feed.scrollTop = feed.scrollHeight - before; // keep the current view anchored
  } catch (e) {
    handleError(e, sel, token);
  } finally {
    olderBtn.disabled = false;
  }
}

function renderFeed(token: TokenFn) {
  const roots = [...messages.values()]
    .filter((m) => m.messageType === "message")
    .sort((a, b) => lastActivity(a).localeCompare(lastActivity(b)));

  const frag = document.createDocumentFragment();
  if (olderLink) frag.append(h("div", { class: "daysep subtle" }, "Older messages available – use “Load older messages”"));
  let day = "";
  for (const m of roots) {
    const label = dayLabel(lastActivity(m));
    if (label !== day) {
      frag.append(h("div", { class: "daysep" }, label));
      day = label;
    }
    frag.append(
      renderPost(m, {
        expanded: expanded.has(m.id),
        unread: Date.parse(lastActivity(m)) > lastSeen,
        onToggle: () => {
          if (expanded.has(m.id)) expanded.delete(m.id);
          else expanded.add(m.id);
          renderFeed(token);
        },
      }),
    );
  }
  if (!roots.length) frag.append(h("div", { class: "empty" }, "No messages in this channel yet."));
  const scroll = feed.scrollTop;
  feed.replaceChildren(frag);
  feed.scrollTop = scroll;
  olderBtn.hidden = !olderLink;
  void loadImages(token);
}

async function loadImages(token: TokenFn) {
  for (const img of feed.querySelectorAll<HTMLImageElement>("img[data-graph-src]")) {
    const src = img.dataset.graphSrc!;
    let p = blobCache.get(src);
    if (!p) {
      p = fetchBlobUrl(token, src);
      blobCache.set(src, p);
      p.catch(() => blobCache.delete(src));
    }
    p.then(
      (url) => {
        img.src = url;
        img.classList.remove("img-loading");
      },
      () => img.classList.replace("img-loading", "img-failed"),
    );
  }
}

async function fillChannelInfo(sel: ChannelSelection, token: TokenFn) {
  try {
    const ch = await getChannel(token, sel.teamId, sel.channelId);
    if (ch.webUrl) {
      openLink.href = ch.webUrl;
      openLink.hidden = false;
    }
    if (!sel.channelName) {
      sel.channelName = ch.displayName;
      renderCrumb(sel);
    }
  } catch {
    /* non-essential */
  }
}

function renderCrumb(sel: ChannelSelection) {
  const tenantLabel = sel.tenantName ?? sel.tenant;
  $("tenant").textContent = `🔒 ${tenantLabel} · external tenant`;
  $("team").textContent = sel.teamName ?? "Team";
  $("channel").textContent = `# ${sel.channelName ?? "Channel"}`;
  $("open-label").textContent = `Open in Teams (${tenantLabel}) ↗`;
  document.title = `${sel.channelName ?? "Channel"} – Peek-a-boo`;
}

function updateLiveText() {
  const sel = isDemo ? demoSelection : loadSelection();
  const who = isDemo ? "demo data" : sel ? signedInUser(sel.tenant) : null;
  if (!lastRefresh) {
    liveText.textContent = "Connecting…";
    return;
  }
  const secs = Math.round((Date.now() - lastRefresh) / 1000);
  const ago = secs < 5 ? "just now" : secs < 60 ? `${secs} s ago` : `${Math.round(secs / 60)} min ago`;
  liveText.textContent = `Live · refreshed ${ago}${who ? ` · signed in as ${who}` : ""}`;
}

function handleError(e: unknown, sel: ChannelSelection, token: TokenFn) {
  liveEl.classList.add("error");
  liveText.textContent = "Not connected";
  const retry = () => void refresh(sel, token, true);
  const doSignIn = async () => {
    try {
      await signIn(sel.tenant, SCOPES);
      retry();
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      if (/AADSTS(65001|90094|90008)|consent|approval/i.test(msg)) showConsent(sel, retry, doSignIn);
      else showState("⚠", "Sign-in failed", msg, [button("Try again", doSignIn, true)]);
    }
  };

  if (e instanceof SignInRequiredError) {
    if (messages.size) return; // keep showing what we have; the user can sign in via refresh
    return showState("🔒", `Sign in to ${sel.tenantName ?? sel.tenant}`, "Use your guest account in the other tenant. A sign-in window will open.", [
      button("Sign in", doSignIn, true),
    ]);
  }
  if (e instanceof GraphError && e.isAuthz) return showConsent(sel, retry, doSignIn, e.message);
  if (messages.size) return; // transient failure during polling: keep the feed, status shows "Not connected"
  showState("⚠", "Could not load the channel", e instanceof Error ? e.message : String(e), [button("Try again", retry, true)]);
}

function showConsent(sel: ChannelSelection, retry: () => void, doSignIn: () => void, detail?: string) {
  const url = adminConsentUrl(sel.tenant);
  const copy = button("Copy consent link", async () => {
    await navigator.clipboard.writeText(url).catch(() => prompt("Copy this link:", url));
    copy.textContent = "Copied ✓";
  });
  showState(
    "🛡",
    "Admin consent needed",
    `Reading channel messages needs the ChannelMessage.Read.All permission, which an admin in ${sel.tenantName ?? sel.tenant} must approve once. Send them the consent link, then try again.${detail ? `\n\nDetails: ${detail}` : ""}`,
    [copy, button("Try again", retry, true), button("Sign in with another account", doSignIn)],
  );
}

function showState(icon: string, title: string, text: string, actions: HTMLElement[] = []) {
  stateBox.replaceChildren(
    h("div", { class: "state-card" }, h("div", { class: "state-icon" }, icon), h("h2", {}, title), h("p", {}, text), h("div", { class: "actions" }, ...actions)),
  );
  stateBox.hidden = false;
  feed.hidden = true;
}

function hideState() {
  stateBox.hidden = true;
  feed.hidden = false;
}

function button(label: string, onClick: () => void, primary = false) {
  const b = h("button", { type: "button", class: primary ? "btn primary" : "btn" }, label);
  b.addEventListener("click", onClick);
  return b;
}

function linkButton(label: string, href: string, primary = false) {
  return h("a", { class: primary ? "btn primary" : "btn", href }, label);
}

void main();
