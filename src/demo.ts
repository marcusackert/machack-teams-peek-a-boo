import type { ChatMessage, Page } from "./graph";
import type { ChannelSelection } from "./settings";

/** `?demo` renders sample data without signing in, for previews and screenshots. */
export const isDemo = new URLSearchParams(location.search).has("demo");

export const demoSelection: ChannelSelection = {
  tenant: "contoso.com",
  tenantName: "Contoso Ltd",
  teamId: "demo-team",
  teamName: "Project Falcon",
  channelId: "demo-channel",
  channelName: "General",
};

const at = (h: number, m: number) => {
  const d = new Date();
  d.setHours(h, m, 0, 0);
  return d.toISOString();
};
const user = (displayName: string) => ({ user: { displayName, userIdentityType: "aadUser" } });

export function demoMessages(): Page<ChatMessage> {
  return {
    value: [
      {
        id: "1",
        messageType: "message",
        createdDateTime: at(9, 12),
        subject: "v2.3 release notes",
        from: user("Anna Lind"),
        body: { contentType: "html", content: "<p>Release notes for v2.3 are up — please review before Friday so we can lock the scope.</p>" },
        attachments: [{ id: "a1", contentType: "reference", contentUrl: "https://contoso.sharepoint.com/release-notes-v2.3.docx", name: "release-notes-v2.3.docx" }],
        reactions: [
          { reactionType: "like", user: {} },
          { reactionType: "like", user: {} },
          { reactionType: "like", user: {} },
          { reactionType: "like", user: {} },
          { reactionType: "👀", user: {} },
          { reactionType: "👀", user: {} },
        ],
        replies: [
          { id: "1a", messageType: "message", createdDateTime: at(9, 20), from: user("Erik Nordin"), body: { contentType: "html", content: "Looks good to me." } },
          { id: "1b", messageType: "message", createdDateTime: at(9, 31), from: user("Johan Karlsson"), body: { contentType: "html", content: "Added a note on the migration step." } },
          { id: "1c", messageType: "message", createdDateTime: at(9, 40), from: user("Priya Shah"), body: { contentType: "html", content: "👍 Approved from QA." } },
        ],
      },
      {
        id: "2",
        messageType: "message",
        createdDateTime: at(10, 5),
        from: user("Erik Nordin"),
        body: { contentType: "html", content: "<p>Build pipeline is green again 🎉 Flaky integration test was a timezone issue.</p>" },
        replies: [{ id: "2a", messageType: "message", createdDateTime: at(10, 7), from: user("Alex Morgan"), body: { contentType: "html", content: "Nice, thanks Erik!" } }],
      },
      {
        id: "3",
        messageType: "message",
        createdDateTime: at(11, 30),
        importance: "high",
        from: user("Sara Ylvén"),
        body: {
          contentType: "html",
          content:
            '<p><at id="0">Alex Morgan</at> can you check the API contract for the orders endpoint? Snippet below.</p><pre>GET /api/v2/orders/{id}\n{\n  "id": "ord_4821",\n  "status": "shipped",\n  "customerRef": "C-1093"\n}</pre>',
        },
      },
    ],
  };
}
