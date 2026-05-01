# Draft GitHub Issue for sbraind/excel-mcp-server

Copy/paste the body below into a new issue at:
**https://github.com/sbraind/excel-mcp-server/issues/new**

Suggested title:

> Add a LICENSE file to clarify MIT terms

---

Suggested body:

```
Hi sbraind 👋

Thank you for building the Excel MCP Server — your foundation (ExcelJS file
mode + the live-editing dispatcher + the PowerShell COM bridge + many of the
chart/pivot/VBA tools) made it possible for me to build a v3.x cross-platform
fork I'm now sharing with my team at Soracom for internal Claude Desktop /
Claude Code use:

→ https://github.com/emilio-S03/Excel-MCP-Server

I credited you prominently in three places (README top-of-file callout,
README links section, package.json contributors[]) and added a joint-
copyright LICENSE file to my fork.

While preparing the fork, I noticed your repo declares the license as MIT
in two places — the badge in README.md and `"license": "MIT"` in
package.json — but doesn't currently include an actual LICENSE file at the
repo root.

Would you mind adding a standard MIT LICENSE file with your copyright? It
would make the license grant unambiguous for anyone forking, packaging
(.mcpb bundles), or reviewing the project (Soracom InfoSec, in my case,
flagged that the README badge alone isn't a legally binding grant).

Happy to open a PR with a standard MIT template pre-filled with your
copyright — just let me know if you'd prefer that, or you can do it
yourself in 30 seconds with GitHub's "Add file → Create new file → name
it LICENSE → use the template" flow.

Either way, thank you again for the great upstream work — much appreciated.

— Emilio
```

---

## What to do once they respond

| Their response | Action |
|---|---|
| They add a LICENSE file (or accept your PR) | You're 100% covered. Update the "Good-faith reliance" note in your LICENSE to reference the upstream LICENSE going forward. |
| They confirm MIT in writing in the issue (no file added) | Screenshot the response, save it. The written confirmation in a public GitHub issue is durable evidence of the grant. |
| They want a different license (e.g., AGPL, custom, attribution-only) | Pause distribution of your fork. Re-evaluate compatibility. Most likely path: just rename + republish with their preferred terms, or contact them to negotiate. |
| They don't respond within 30 days | Document the attempt in the LICENSE good-faith note. Continue using based on the public package.json declaration. The unrebutted public declaration of MIT in their package.json + README badge is what most courts would weigh as intent. |
| They ask you to take it down | Take it down. The cost of compliance is low; the value of being a good open-source citizen is high. |

## Track this

After you post the issue, log it somewhere persistent so you remember to follow up.
