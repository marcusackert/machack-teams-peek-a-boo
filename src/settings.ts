/** Which external channel to show. Stored per viewer; tab URLs can override it. */
export interface ChannelSelection {
  tenant: string; // domain or tenant ID of the external tenant
  tenantName?: string;
  teamId: string;
  teamName?: string;
  channelId: string;
  channelName?: string;
}

const KEY = "peekaboo.selection";

export const CLIENT_ID: string = import.meta.env.VITE_CLIENT_ID ?? "";
export const DEFAULT_TENANT: string = import.meta.env.VITE_DEFAULT_TENANT ?? "";

export function loadSelection(): ChannelSelection | null {
  const fromUrl = selectionFromParams(new URLSearchParams(location.search));
  if (fromUrl) return fromUrl;
  try {
    const raw = localStorage.getItem(KEY);
    return raw ? (JSON.parse(raw) as ChannelSelection) : null;
  } catch {
    return null;
  }
}

export function saveSelection(sel: ChannelSelection): void {
  try {
    localStorage.setItem(KEY, JSON.stringify(sel));
  } catch {
    /* storage blocked: selection lives only in the tab URL */
  }
}

export function selectionToParams(sel: ChannelSelection): URLSearchParams {
  const p = new URLSearchParams({ tenant: sel.tenant, team: sel.teamId, channel: sel.channelId });
  if (sel.tenantName) p.set("tenantName", sel.tenantName);
  if (sel.teamName) p.set("teamName", sel.teamName);
  if (sel.channelName) p.set("channelName", sel.channelName);
  return p;
}

function selectionFromParams(p: URLSearchParams): ChannelSelection | null {
  const tenant = p.get("tenant");
  const teamId = p.get("team");
  const channelId = p.get("channel");
  if (!tenant || !teamId || !channelId) return null;
  return {
    tenant,
    teamId,
    channelId,
    tenantName: p.get("tenantName") ?? undefined,
    teamName: p.get("teamName") ?? undefined,
    channelName: p.get("channelName") ?? undefined,
  };
}

/** Absolute URL of a page that sits next to the current one (works under any base path). */
export function pageUrl(file: string, params?: URLSearchParams): string {
  const u = new URL(file, location.href);
  u.search = params ? params.toString() : "";
  u.hash = "";
  return u.toString();
}
