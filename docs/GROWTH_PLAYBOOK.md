# Growth playbook

> Internal doc for the maintainer. Snapshot taken 2026-11-03: **328 stars, 43 forks, 3 open issues** (shields.io).

## 1. Diagnosis

A fork-to-star ratio of 13% is healthy, so people who find the repo use it. The gap is **discovery** and **first-screen conversion**, not product quality. Comparable PPT skills reported by third-party roundups sit at 2k to 48k stars (frontend-slides, guizang-ppt-skill, ppt-master, MiniMax pptx-generator); the winners share three traits: a visual hook above the fold, a one-line install, and presence in skill directories and awesome lists.

What the audit of this repo found:

| # | Finding | Effect on stars | Fixed in v2.4.0 |
|---|---|---|---|
| 1 | README opened with architecture jargon (Harness, GPU to CPU) and ran 21 KB | Visitor decides in 10 seconds; no "what do I get" | Yes: rewritten, 9 KB |
| 2 | No way to try it without an agent and a hard-coded `~/.workbuddy/...` path | Visitors cannot get a result in 60 seconds | Yes: `pip install mck-ppt`, `python -m mck_ppt demo` |
| 3 | Install only documented for WorkBuddy | Excludes the Claude Code and Cursor audience, the largest | Yes: 4 install routes (verify two, see section 3) |
| 4 | Layout count inconsistent: 67 / 70 / 55 across README, SKILL.md, repo description | Looks careless; reviewers distrust numbers | Yes: 62 active layouts everywhere (repo description: run `scripts/gh_setup.sh`) |
| 5 | README's own A/B table showed the Harness deck scoring lower (92 vs 95) | Undercuts the headline claim | Moved to `docs/HARNESS.md` |
| 6 | `docs/screenshots/` empty; no gallery; layouts only described in prose | No proof for 62 layouts | Partly: `docs/GALLERY.md` and `render_preview.py` added; images still to render (section 3) |
| 7 | No `.github/` (CI, issue templates), no Discussions, version stuck at 2.3.0 in code | Reads as a personal dump, not a maintained project | Yes |
| 8 | `run_qa_tests.py` reported 41 errors; one cause was a wrong `waterfall` fixture | A visitor running tests sees failure | Waterfall fixed; the rest are real layout or fixture issues to triage |
| 9 | Chinese audience (the likeliest early adopters) had one collapsed `<details>` block | Chinese readers are the natural base | Yes: `README.zh-CN.md` |

## 2. Positioning

One sentence: **the PPT skill for people whose file has to survive being opened and edited in PowerPoint**, with consulting-style structure and a code-enforced quality gate.

Do not compete on "most layouts" or "prettiest HTML". Compete on three provable things: editable native shapes, gate scripts the agent cannot talk its way past, Chinese typography that works. Every claim in the README should be checkable by running the demo.

## 3. Do this week (ordered by effect per hour)

1. **Render the gallery (2 h).** Install LibreOffice and poppler, run `python -m mck_ppt demo` then `python scripts/render_preview.py mck_demo.pptx docs/gallery`. Commit the PNGs. Then render all 62 layouts from `run_qa_tests.py` the same way and embed 6 to 8 of the best in the README top table (replace the hosted images if these look better).
2. **Record a 20 to 30 second GIF (1 h).** Left: one Chinese prompt typed into the agent. Right: the finished 15-page deck scrolling in PowerPoint. Put it directly under the title. This single asset typically moves README conversion more than any text.
3. **Run `scripts/gh_setup.sh` (5 min)** for description, topics and Discussions; upload `docs/social-preview.png` in Settings. The social preview is what appears when the link is pasted into WeChat, X, Slack.
4. **Verify the two unverified install routes** before announcing: `/plugin marketplace add` with `.claude-plugin/marketplace.json` (the `skills: ["./"]` field depends on the Claude Code version) and `npx skills add likaku/Mck-ppt-design-skill`. If either fails, delete that row from both READMEs.
5. **Publish to PyPI (30 min):** `python -m build && twine upload dist/*`. The name `mck-ppt` was free on 2026-11-03.
6. **Merge, tag `v2.4.0`, publish a GitHub Release** with the demo deck attached (`release.yml` does this on tag). Pin the release.
7. **Triage the QA failures** from `python run_qa_tests.py`. A green CI badge is a trust signal.

## 4. Launch (days 1 to 14)

| Day | Channel | Asset | Ask |
|---|---|---|---|
| 1 | WeChat group (200+) and moments | GIF + one line: "pip install, 60 seconds" | Try the demo, star if useful |
| 1 | Discord | Same | Same |
| 2 | 掘金 / 知乎 article | "让 AI 做 PPT，为什么我用代码而不是提示词做质量门禁" (the gate story, with the A/B numbers honestly stated) | Comment with your worst AI-PPT failure |
| 3 | 即刻 / 小红书 | 6 images: before (agent free-form) vs after (this skill), same prompt | Save and share |
| 4 | V2EX "分享创造" | Short post, link to demo command | Feedback on layouts |
| 5 | Show HN: "Show HN: Editable consulting-style PPTX for AI agents, with code-enforced QA gates" | GIF, repo | Honest technical comments |
| 6 | Reddit r/ClaudeAI, r/ChatGPTCoding | Same GIF, focus on Claude Code install | Feedback |
| 7 | X / LinkedIn | Thread: 5 slides produced by the skill, one prompt each | RT |
| 8 to 14 | Awesome lists | PR one line each: ComposioHQ/awesome-claude-skills (confirmed to exist), plus other awesome-claude-skills / awesome-agent-skills lists (check each list's contribution rules first) | Merge |

Rules for posts: lead with the output image, not the architecture; include the exact command; state limits plainly (needs LibreOffice for preview, python-pptx constraints, no SmartArt). Reply to every comment within a day for the first two weeks.

## 5. Keep it growing (days 15 to 90)

- **Weekly rhythm:** one "deck of the week" from the showcase issues, one new layout or fix, one release note. Small releases beat big ones for the Trending algorithm and for "last updated" trust.
- **Layout of the week post** (image + prompt + code, 5 lines). 62 layouts is about 14 months of material; start with the 10 most striking.
- **Template packs as separate content:** "board update", "market entry", "annual plan", each a storyline JSON in `mck_ppt/storylines/`. Each is a searchable use case and a reason to share.
- **Make contributing cheap:** label 5 issues `good first issue` (new layout in the community-request template, theme color variant, Keynote rendering check).
- **Theme switch** (navy / graphite / warm) is the most visible v2.5 feature: same content, three looks, one screenshot that sells itself.
- **Collaborations:** ask 3 to 5 Chinese AI-tooling creators with decks in the consulting niche to run one real task with it and post the result. Offer help, not payment.

## 6. Measure

| Metric | Where | Review |
|---|---|---|
| Visitors to stars | Insights > Traffic (14-day window, export weekly) | Weekly |
| Clones and unique cloners | Insights > Traffic | Weekly |
| Referrers (which channel converts) | Insights > Traffic > Referring sites | After each launch post |
| PyPI downloads | pypistats.org/packages/mck-ppt | Weekly |
| Showcase issues, external PRs | Issues | Monthly |

A reasonable first goal is 1,000 stars within 60 days of the launch week. That is a target, not a forecast: it depends mostly on whether the GIF and one or two posts land.

## 7. Do not

- Do not buy stars, join star-exchange groups, or ask for reciprocal stars. GitHub removes them and the damage to trust is permanent.
- Do not quote competitor star counts in the README; they go stale and invite comparison on the wrong axis.
- Do not claim "best" or "first". Claim what the demo proves.
