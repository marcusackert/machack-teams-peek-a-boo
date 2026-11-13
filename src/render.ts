import DOMPurify from "dompurify";
import type { ChatMessage, Identity } from "./graph";

export function h<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  attrs: Record<string, string> = {},
  ...children: (Node | string | null | undefined | false)[]
): HTMLElementTagNameMap[K] {
  const el = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (k === "class") el.className = v;
    else el.setAttribute(k, v);
  }
  for (const c of children) if (c) el.append(c);
  return el;
}

const PALETTE = ["#e3a0c9", "#8fd3a8", "#9fb6ff", "#f4b183", "#c5a3ff", "#f6d36b", "#86d1e0", "#ff9e9e"];

export function avatar(name: string, small = false): HTMLElement {
  const initials =
    name
      .split(/[\s.@_-]+/)
      .filter(Boolean)
      .slice(0, 2)
      .map((p) => p[0]!.toUpperCase())
      .join("") || "?";
  let hash = 0;
  for (const ch of name) hash = (hash * 31 + ch.charCodeAt(0)) | 0;
  const el = h("div", { class: small ? "av av-sm" : "av", "aria-hidden": "true" }, initials);
  el.style.background = PALETTE[Math.abs(hash) % PALETTE.length]!;
  return el;
}

export function authorName(m: ChatMessage): string {
  return m.from?.user?.displayName ?? m.from?.application?.displayName ?? "Unknown";
}

function isExternal(user: Identity | null | undefined, homeTenantId?: string): boolean {
  return !!(user?.tenantId && homeTenantId && user.tenantId !== homeTenantId);
}

export function formatTime(iso: string): string {
  const d = new Date(iso);
  const now = new Date();
  const time = d.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
  if (d.toDateString() === now.toDateString()) return time;
  const sameYear = d.getFullYear() === now.getFullYear();
  return `${d.toLocaleDateString([], { month: "short", day: "numeric", year: sameYear ? undefined : "numeric" })} ${time}`;
}

export function dayLabel(iso: string): string {
  const d = new Date(iso);
  const today = new Date();
  const yesterday = new Date(Date.now() - 86_400_000);
  if (d.toDateString() === today.toDateString()) return "Today";
  if (d.toDateString() === yesterday.toDateString()) return "Yesterday";
  return d.toLocaleDateString([], { weekday: "long", month: "long", day: "numeric", year: "numeric" });
}

const REACTIONS: Record<string, string> = {
  like: "👍",
  heart: "❤️",
  laugh: "😆",
  surprised: "😮",
  sad: "😢",
  angry: "😡",
};

DOMPurify.addHook("afterSanitizeAttributes", (node) => {
  if (node.tagName === "A") {
    node.setAttribute("target", "_blank");
    node.setAttribute("rel", "noopener noreferrer");
  }
  if (node.tagName === "IMG") {
    const src = node.getAttribute("src") ?? "";
    // Graph-hosted images need a bearer token; main.ts swaps them for blob URLs.
    if (src.startsWith("https://graph.microsoft.com/")) {
      node.setAttribute("data-graph-src", src);
      node.removeAttribute("src");
      node.classList.add("img-loading");
    }
    node.setAttribute("loading", "lazy");
  }
});

function bodyNode(m: ChatMessage): HTMLElement {
  const div = h("div", { class: "body" });
  if (m.body.contentType === "text") {
    div.textContent = m.body.content;
    return div;
  }
  const html = m.body.content
    .replace(/<at\b[^>]*>/gi, '<span class="mention">')
    .replace(/<\/at>/gi, "</span>")
    .replace(/<attachment\b[^>]*><\/attachment>/gi, "");
  div.append(
    DOMPurify.sanitize(html, {
      RETURN_DOM_FRAGMENT: true,
      FORBID_TAGS: ["style", "form", "input", "button"],
      FORBID_ATTR: ["style"],
    }),
  );
  return div;
}

function attachmentNodes(m: ChatMessage): HTMLElement[] {
  const out: HTMLElement[] = [];
  for (const a of m.attachments ?? []) {
    if (a.contentType === "reference" && a.contentUrl) {
      const name = a.name ?? a.contentUrl.split("/").pop() ?? "file";
      const ext = (name.split(".").pop() ?? "").toLowerCase();
      const kind = ["doc", "docx"].includes(ext) ? "w" : ["xls", "xlsx", "csv"].includes(ext) ? "x" : ["ppt", "pptx"].includes(ext) ? "p" : ext === "pdf" ? "pdf" : "f";
      const label = { w: "W", x: "X", p: "P", pdf: "PDF", f: "📄" }[kind];
      out.push(
        h(
          "a",
          { class: "file", href: a.contentUrl, target: "_blank", rel: "noopener noreferrer" },
          h("span", { class: `ficon ficon-${kind}` }, label),
          h("span", {}, name, h("small", {}, "Open in the other tenant")),
        ),
      );
    } else if (a.contentType === "messageReference" && a.content) {
      try {
        const ref = JSON.parse(a.content) as { messagePreview?: string; messageSender?: { user?: Identity } };
        out.push(
          h("blockquote", { class: "quote" }, h("b", {}, ref.messageSender?.user?.displayName ?? ""), " ", ref.messagePreview ?? ""),
        );
      } catch {
        /* ignore malformed reference */
      }
    } else if (a.contentType.startsWith("application/vnd.microsoft.card")) {
      out.push(h("div", { class: "card-ph" }, "🃏 Card – open in Teams to view"));
    }
  }
  return out;
}

function reactionsNode(m: ChatMessage): HTMLElement | null {
  if (!m.reactions?.length) return null;
  const counts = new Map<string, number>();
  for (const r of m.reactions) counts.set(r.reactionType, (counts.get(r.reactionType) ?? 0) + 1);
  return h("div", { class: "react" }, ...[...counts].map(([t, n]) => h("span", { title: t }, `${REACTIONS[t] ?? t} ${n}`)));
}

function messageCore(m: ChatMessage, homeTenantId: string | undefined, isReply: boolean): HTMLElement[] {
  const name = authorName(m);
  if (m.deletedDateTime) {
    return [avatar(name, isReply), h("div", { class: "content" }, h("div", { class: "deleted" }, "This message has been deleted."))];
  }
  const meta = h(
    "div",
    { class: "meta" },
    h("span", { class: "name" }, name),
    isExternal(m.from?.user, homeTenantId) && h("span", { class: "badge" }, "External"),
    h("time", { class: "time", datetime: m.createdDateTime, title: new Date(m.createdDateTime).toLocaleString() }, formatTime(m.createdDateTime)),
    m.lastEditedDateTime && h("span", { class: "time" }, "Edited"),
    m.importance === "urgent" ? h("span", { class: "badge urgent" }, "URGENT!") : m.importance === "high" ? h("span", { class: "badge important" }, "IMPORTANT!") : null,
  );
  return [
    avatar(name, isReply),
    h(
      "div",
      { class: "content" },
      meta,
      !isReply && m.subject ? h("div", { class: "subject" }, m.subject) : null,
      bodyNode(m),
      ...attachmentNodes(m),
      reactionsNode(m),
    ),
  ];
}

export function lastActivity(m: ChatMessage): string {
  let last = m.createdDateTime;
  for (const r of m.replies ?? []) if (r.createdDateTime > last) last = r.createdDateTime;
  return last;
}

/** Renders a root post with its replies. Threads with >1 reply start collapsed unless `expanded`. */
export function renderPost(
  m: ChatMessage,
  opts: { homeTenantId?: string; expanded: boolean; unread: boolean; onToggle: () => void },
): HTMLElement {
  const post = h("article", { class: `post${opts.unread ? " unread" : ""}`, "data-id": m.id });
  post.append(h("div", { class: "msg" }, ...messageCore(m, opts.homeTenantId, false)));

  const replies = [...(m.replies ?? [])]
    .filter((r) => r.messageType === "message")
    .sort((a, b) => a.createdDateTime.localeCompare(b.createdDateTime));
  if (!replies.length) return post;

  const box = h("div", { class: "replies" });
  const collapsible = replies.length > 1;
  if (collapsible) {
    const names = [...new Set(replies.map(authorName))];
    const who = names.length <= 3 ? names.map((n) => n.split(" ")[0]).join(", ").replace(/, ([^,]*)$/, " and $1") : `${names.length} people`;
    const last = replies[replies.length - 1]!;
    const more = h(
      "button",
      { class: "more", type: "button", "aria-expanded": String(opts.expanded) },
      h("span", { class: "avs" }, ...names.slice(0, 3).map((n) => avatar(n, true))),
      `${replies.length} replies from ${who} · last ${formatTime(last.createdDateTime)} ${opts.expanded ? "▴" : "▾"}`,
    );
    more.addEventListener("click", opts.onToggle);
    box.append(more);
  }
  if (!collapsible || opts.expanded) {
    for (const r of replies) box.append(h("div", { class: "reply" }, ...messageCore(r, opts.homeTenantId, true)));
  }
  post.append(box);
  return post;
}
