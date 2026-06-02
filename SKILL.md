---
name: weaver-fintech-slides
description: >
  Build on-brand Weaver Fintech PowerPoint presentations as editable .pptx files using pptxgenjs.
  Use whenever anyone asks to create, build, or generate a Weaver Fintech presentation or deck —
  regardless of topic. This skill is fully self-contained: logos are base64-embedded, photos load
  from GitHub raw URLs. Works on any Claude account with zero local file dependencies.
  Always output an editable .pptx file — never a PDF, image, or HTML.
  The brand is strictly enforced — never introduce colours, fonts, shapes, or layouts not defined here.
---

# Weaver Fintech Presentation Skill

## ⚠️ One-Time Setup (per Claude account)

Before generating any deck, run this setup block to write the logo files to disk:

```javascript
// setup.js — run once per session: node setup.js
const fs = require('fs');

const LOGO_FULL_B64  = 'PHN2ZyB2aWV3Qm94PSIwIDAgNzA2LjMyIDMxNC40MSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIiB4bWxuczp4bGluaz0iaHR0cDovL3d3dy53My5vcmcvMTk5OS94bGluayIgaWQ9IkxheWVyXzEiIG92ZXJmbG93PSJoaWRkZW4iPjxnPjxwYXRoIGQ9Ik0xNzAuNDcgNjAuMzkgMTU2LjY3IDk3LjY3IDE1Ni42NyA5Ny42N0MxNTYuNjcgOTcuNjcgMTQzLjE5IDEzNC4wOSAxNDMuMTkgMTM0LjA5TDExNi44NSAxMzAuNjMgMTE4LjMgMTI2LjczIDE0Mi4wNCA2Mi41NkMxNDIuOTYgNjAuMDcgMTQ1LjIyIDU4LjMxIDE0Ny44NyA1OC4wOCAxNTAuNzYgNTcuODMgMTUzLjY5IDU3LjcgMTU2LjY2IDU3LjcgMTYwLjQ1IDU3LjcgMTY0LjE4IDU3LjkxIDE2Ny44NiA1OC4zMiAxNjkuMDUgNTguNDUgMTcwLjA2IDU5LjI2IDE3MC40NyA2MC4zOUwxNzAuNDcgNjAuMzlaIiBmaWxsPSIjQzRGRjAwIi8+PHBhdGggZD0iTTE5NS4wMyAxMjYuNzQgMTg3LjEzIDEzOS44OSAxNzEuNTMgMTM3LjgzIDE1Ni42NyA5Ny42OCAxNTYuNjcgOTcuNjcgMTcwLjQ3IDYwLjM5IDE5NS4wMyAxMjYuNzRaIi8+PHBhdGggZD0iTTI1Ni41OSAxNTcuNjJDMjU2LjU5IDE2MS42MyAyNTYuMzUgMTY1LjU4IDI1NS44OSAxNjkuNDYgMjU1LjUzIDE3Mi41MSAyNTIuNjkgMTc0LjY1IDI0OS42NSAxNzQuMjVMMjEwLjcxIDE2OS4xMyAyMDAuNiAxNjcuNzkgMjA2LjU0IDE1Ny44OCAyMTUuMTEgMTQzLjU5IDI1MS42NyAxNDguNEMyNTQuMzUgMTQ4Ljc1IDI1Ni40MSAxNTAuOTggMjU2LjUxIDE1My42OCAyNTYuNTYgMTU0Ljk5IDI1Ni41OSAxNTYuMzEgMjU2LjU5IDE1Ny42M1oiIGZpbGw9IiM4NzVFRkEiLz48cGF0aCBkPSJNMjI4LjY5IDIyNi44OEMyMjQuODQgMjMwLjg5IDIyMC42NSAyMzQuNTggMjE2LjE3IDIzNy45IDIxMy4xOCAyNDAuMTEgMjA4LjkxIDIzOC44NSAyMDcuNjIgMjM1LjM2TDE4OS40OCAxODYuMzIgMjAwLjYgMTY3Ljc3IDIxMC43MSAxNjkuMTEgMjI5LjkzIDIyMS4wNEMyMzAuNjggMjIzLjA2IDIzMC4xOCAyMjUuMzIgMjI4LjY5IDIyNi44N1oiLz48cGF0aCBkPSJNMTE4LjMgMTI2LjczIDExNi44NSAxMzAuNjMgODguMTYgMTI2Ljg1IDc1LjI4IDEwNS4zOUM3NC4xMyAxMDMuNDcgNzQuMjUgMTAxLjA1IDc1LjU2IDk5LjIzIDc4Ljc3IDk0Ljc3IDgyLjM0IDkwLjU5IDg2LjIzIDg2LjczIDg4LjggODQuMTggOTMuMDggODQuNyA5NC45NSA4Ny44TDExOC4zIDEyNi43MloiIGZpbGw9IiMyQTYyRjUiLz48cGF0aCBkPSJNMTMzLjk3IDE1OSAxMjMuODUgMTg2LjM1IDEwNS43MSAyMzUuMzdDMTA0LjQyIDIzOC44NiAxMDAuMTQgMjQwLjEyIDk3LjE2IDIzNy45MSA5Mi42OCAyMzQuNTkgODguNDkgMjMwLjg5IDg0LjYzIDIyNi44OCA4My4xNCAyMjUuMzMgODIuNjQgMjIzLjA3IDgzLjM5IDIyMS4wNUwxMDYuNzcgMTU3Ljg4IDEwNy42MyAxNTUuNTQgMTMzLjk3IDE1OVoiIGZpbGw9IiNDNEZGMDAiLz48cGF0aCBkPSJNMTg2LjA1IDI1My4xNkMxODAuODYgMjU0Ljc2IDE3NS40OSAyNTUuOTQgMTY5Ljk4IDI1Ni42NyAxNjcuNzYgMjU2Ljk2IDE2NS41OCAyNTUuOTMgMTY0LjQzIDI1NC4wMkwxNTYuNjYgMjQxLjA3IDE1Ni42NiAyNDEuMDdDMTU2LjY2IDI0MS4wNyAxNDEuNTYgMjE1Ljg5IDE0MS41NiAyMTUuODlMMTIzLjg1IDE4Ni4zNiAxMzMuOTcgMTU5LjAxIDEzNy45NiAxNTkuNTQgMTU2LjY3IDE5MC43MSAxNzEuNzYgMjE1Ljg5IDE4OS4yMSAyNDQuOThDMTkxLjA3IDI0OC4wOSAxODkuNTIgMjUyLjExIDE4Ni4wNiAyNTMuMThaIiBmaWxsPSIjMkE2MkY1Ii8+PHBhdGggZD0iTTE1Ni42NiAyNDEuMDUgMTU2LjY2IDI0MS4wNUMxNTYuNjYgMjQxLjA1IDE0OC45IDI1NC4wMSAxNDguOSAyNTQuMDEgMTQ3Ljc1IDI1NS45MyAxNDUuNTYgMjU2Ljk2IDE0My4zNSAyNTYuNjcgMTM3Ljg0IDI1NS45NCAxMzIuNDcgMjU0Ljc2IDEyNy4yOCAyNTMuMTYgMTIzLjgyIDI1Mi4xIDEyMi4yNiAyNDguMDcgMTI0LjEzIDI0NC45N0wxNDEuNTcgMjE1Ljg4IDE1Ni42NyAyNDEuMDVaIiBmaWxsPSIjMURDRDg2Ii8+PHBhdGggZD0iTTIzNy43NSA5OS4yMUMyMzkuMDYgMTAxLjAzIDIzOS4xOCAxMDMuNDQgMjM4LjAzIDEwNS4zNkwyMTUuMTIgMTQzLjU3IDIwNi41NSAxNTcuODYgMjAwLjYxIDE2Ny43NyAxODkuNDkgMTg2LjMyIDE3MS43NiAyMTUuODcgMTU2LjY3IDE5MC42OSAxNzIuNjIgMTY0LjA5IDE3Ny45NSAxNTUuMiAxODcuMTMgMTM5Ljg5IDE5NS4wMiAxMjYuNzQgMjE4LjM2IDg3LjgxQzIyMC4yMiA4NC43IDIyNC41IDg0LjE4IDIyNy4wOCA4Ni43MyAyMzAuOTYgOTAuNTggMjM0LjUzIDk0Ljc2IDIzNy43NCA5OS4yMVoiIGZpbGw9IiMxRENEODYiLz48cGF0aCBkPSJNMTg3LjEzIDEzOS44OSAxNzcuOTUgMTU1LjIgMTcyLjYyIDE2NC4wOSAxMzcuOTggMTU5LjUzIDEzNy45NyAxNTkuNTMgMTMzLjk4IDE1OSAxMDcuNjQgMTU1LjU0IDYyLjU1IDE0OS41OUM1OS41MSAxNDkuMTkgNTcuMzIgMTQ2LjM5IDU3Ljc2IDE0My4zNSA1OC41MSAxMzguMDcgNTkuNjggMTMyLjkxIDYxLjIzIDEyNy45MyA2Mi4wMyAxMjUuMzUgNjQuNTkgMTIzLjc0IDY3LjI3IDEyNC4wOUw4OC4xNyAxMjYuODUgMTE2Ljg2IDEzMC42MyAxNDMuMiAxMzQuMDkgMTcxLjU0IDEzNy44MyAxODcuMTQgMTM5Ljg5WiIgZmlsbD0iIzg3NUVGQSIvPjwvZz48Zz48Zz48cGF0aCBkPSJNNjQ4LjM3IDg1LjY0QzY0Mi44NyA4NS42NCA2MzcuOTYgODYuOTMgNjMzLjYyIDg5LjUyIDYzMy4yMiA4OS43NSA2MzIuODQgOTAuMDEgNjMyLjQ2IDkwLjI2TDYzMi40NiA4Ni44NyA2MTkuMDkgODYuODcgNjE5LjA5IDE0My4xMiA2MzIuNTYgMTQzLjEyIDYzMi41MSAxMTUuODRDNjMyLjU4IDExNS4zMSA2MzIuNjIgMTE0LjgyIDYzMi42MiAxMTQuNCA2MzIuNjIgMTEwLjk4IDYzMy4yNSAxMDcuOTggNjM0LjQ5IDEwNS40IDYzNS44MyAxMDIuNzQgNjM3LjY2IDEwMC42NSA2NDAgOTkuMTUgNjQyLjQxIDk3LjY1IDY0NS4xNiA5Ni45MSA2NDguMjUgOTYuOTFMNjQ4LjYzIDk2LjkxIDY0OC42MyA4NS42NSA2NDguMzggODUuNjVaIi8+PHBhdGggZD0iTTM2My4yOSA4Ni44NiAzNTAuNjQgMTI0Ljk4IDMzOC41NSA5MS4yOEMzMzcuNiA4OC42MyAzMzUuMDkgODYuODYgMzMyLjI3IDg2Ljg2IDMyOS40NSA4Ni44NiAzMjYuOTEgODguNjUgMzI1Ljk3IDkxLjMzTDMxNC4xNyAxMjQuOTggMzAxLjIxIDg2Ljg2IDI4OC4yNyA4Ni44NiAzMDUuNTMgMTM4Ljk3QzMwNi42NyAxNDIuMzEgMzA5LjgxIDE0NC41NSAzMTMuMzMgMTQ0LjU1TDMxMy4zMyAxNDQuNTVDMzE2LjgxIDE0NC41NSAzMTkuOTIgMTQyLjM2IDMyMS4wOSAxMzkuMDhMMzMyLjI5IDEwNi4yMyAzNDMuNjggMTM5LjEyQzM0NC44NiAxNDIuMzggMzQ3Ljk2IDE0NC41NSAzNTEuNDMgMTQ0LjU1IDM1NC45IDE0NC41NSAzNTguMTIgMTQyLjI5IDM1OS4yNSAxMzguOTNMMzc2LjIyIDg2Ljg2IDM2My4yOCA4Ni44NloiLz48cGF0aCBkPSJNNDY1Ljk3IDE0NC40NUM0NzMuMzMgMTQ0LjQ1IDQ3OS40NCAxNDEuNjUgNDg0LjMgMTM2LjE2TDQ4NC4zIDE0My4xIDQ5Ny4xNCAxNDMuMSA0OTcuMTQgODYuODUgNDg0LjMgODYuODUgNDg0LjMgOTRDNDc5LjIzIDg4LjMgNDczLjExIDg1LjUxIDQ2NS45NyA4NS41MSA0NTcuOTkgODUuNTEgNDUxLjI2IDg4LjMxIDQ0NS45OCA5My44IDQ0MC43IDk5LjI5IDQzOC4xMSAxMDYuMzMgNDM4LjExIDExNC43MiA0MzguMTEgMTIzLjExIDQ0MC43IDEzMC4yNiA0NDUuOTggMTM1Ljk1IDQ1MS4yNiAxNDEuNjUgNDU3Ljg5IDE0NC40NCA0NjUuOTcgMTQ0LjQ0Wk00NTYuMTMgMTAyLjgxQzQ1OS4yNCA5OS40OSA0NjMuMTcgOTcuODQgNDY4LjA0IDk3Ljg0IDQ3Mi45MSA5Ny44NCA0NzYuODUgOTkuNSA0NzkuOTUgMTAyLjgxIDQ4My4wNiAxMDYuMTMgNDg0LjYxIDExMC4xNiA0ODQuNjEgMTE0LjkzIDQ4NC42MSAxMTkuNyA0ODMuMDYgMTIzLjg0IDQ3OS45NSAxMjcuMTUgNDc2Ljg0IDEzMC40NyA0NzIuOTEgMTMyLjEyIDQ2OC4wNCAxMzIuMTIgNDYzLjE3IDEzMi4xMiA0NTkuMjQgMTMwLjQ2IDQ1Ni4xMyAxMjcuMTUgNDUzLjEzIDEyMy44NCA0NTEuNTcgMTE5LjggNDUxLjU3IDExNC45MyA0NTEuNTcgMTEwLjA2IDQ1My4xMiAxMDYuMTMgNDU2LjEzIDEwMi44MVoiLz48cGF0aCBkPSJNNTI5LjE3IDE0NC4yNyA1MjkuMTcgMTQ0LjMyQzUzMi4yNSAxNDQuMzIgNTM1LjA0IDE0Mi41IDUzNi4yNyAxMzkuNjdMNTU4Ljc5IDg2Ljg2IDU0NS4xMyA4Ni44NiA1MjkuMzcgMTI0LjU2IDUxMy40MSA4Ni44NiA0OTkuNjQgODYuODYgNTIyLjA2IDEzOS42NkM1MjMuMjkgMTQyLjQ5IDUyNi4wOCAxNDQuMjcgNTI5LjE3IDE0NC4yN1oiLz48cGF0aCBkPSJNNDA3LjU4IDE0My4xNEM0MDEuNjMgMTQzLjE0IDM5Ni4zNCAxNDEuOTIgMzkxLjcgMTM5LjQ3IDM4Ny4wNiAxMzcuMDMgMzgzLjQ0IDEzMy42NSAzODAuODMgMTI5LjMzIDM3OC4zMSAxMjQuOTQgMzc3LjA0IDEyMC4wMSAzNzcuMDQgMTE0LjU1IDM3Ny4wNCAxMDkuMDkgMzc4LjE4IDEwNC4xNyAzODAuNDcgOTkuNzcgMzgyLjgzIDk1LjM3IDM4Ni4wOSA5MS45NSAzOTAuMjQgODkuNSAzOTQuNDggODYuOTggMzk5LjI4IDg1LjcxIDQwNC42NiA4NS43MUw0MDQuOTEgODUuNzFDNDEwLjI4IDg1LjcxIDQxNS4wNSA4Ni45NyA0MTkuMiA4OS41IDQyMy40MyA5MS45NSA0MjYuNzMgOTUuMzYgNDI5LjA5IDk5Ljc3IDQzMS40NiAxMDQuMDggNDMyLjYzIDEwOC45NyA0MzIuNjMgMTE0LjQyTDQzMi42MyAxMTkuNzkgNDIwLjQyIDExOS43OSA0MjAuNDIgMTE1LjY0QzQyMC40MiAxMTEuNjUgNDE5Ljc3IDEwOC4yMyA0MTguNDcgMTA1LjM4IDQxNy4xNiAxMDIuNTMgNDE1LjMzIDEwMC4zNyA0MTIuOTcgOTguOSA0MTAuNjEgOTcuNDQgNDA3LjkzIDk2LjcxIDQwNC45MSA5Ni43MUw0MDQuNTQgOTYuNzFDNDAxLjUzIDk2LjcxIDM5OC44NCA5Ny40NCAzOTYuNDggOTguOSAzOTQuMiAxMDAuMzcgMzkyLjQxIDEwMi40MSAzOTEuMSAxMDUuMDEgMzg5Ljg4IDEwNy41MyAzODkuMjcgMTEwLjQ2IDM4OS4yNyAxMTMuODEgMzg5LjI3IDExOC45NCAzOTAuNzQgMTIzLjIxIDM5My42NyAxMjYuNjMgMzk2LjYgMTI5Ljk3IDQwMS4yIDEzMS42NCA0MDcuNDcgMTMxLjY0TDQyNy42MyAxMzEuNjQgNDI3LjYzIDE0My4xMiA0MDcuNiAxNDMuMTJaTTM4Mi4wNSAxMjAuMDUgMzgyLjA1IDEwOC44MSA0MzEuMDMgMTA4LjgxIDQzMi42MiAxMTYuOTkgNDMyLjYyIDEyMC4wNSAzODIuMDYgMTIwLjA1WiIvPjxwYXRoIGQ9Ik01ODcuNzIgMTQzLjE0QzU4MS43NyAxNDMuMTQgNTc2LjQ4IDE0MS45MiA1NzEuODQgMTM5LjQ3IDU2Ny4yIDEzNy4wMyA1NjMuNTggMTMzLjY1IDU2MC45NyAxMjkuMzMgNTU4LjQ1IDEyNC45NCA1NTcuMTggMTIwLjAxIDU1Ny4xOCAxMTQuNTUgNTU3LjE4IDEwOS4wOSA1NTguMzIgMTA0LjE3IDU2MC42MSA5OS43NyA1NjIuOTcgOTUuMzcgNTY2LjIzIDkxLjk1IDU3MC4zOCA4OS41IDU3NC42MiA4Ni45OCA1NzkuNDIgODUuNzEgNTg0LjggODUuNzFMNTg1LjA1IDg1LjcxQzU5MC40MiA4NS43MSA1OTUuMTkgODYuOTcgNTk5LjM0IDg5LjUgNjAzLjU3IDkxLjk1IDYwNi44NyA5NS4zNiA2MDkuMjMgOTkuNzcgNjExLjYgMTA0LjA4IDYxMi43NyAxMDguOTcgNjEyLjc3IDExNC40Mkw2MTIuNzcgMTE5Ljc5IDYwMC41NiAxMTkuNzkgNjAwLjU2IDExNS42NEM2MDAuNTYgMTExLjY1IDU5OS45MSAxMDguMjMgNTk4LjYxIDEwNS4zOCA1OTcuMyAxMDIuNTMgNTk1LjQ3IDEwMC4zNyA1OTMuMTEgOTguOSA1OTAuNzUgOTcuNDQgNTg4LjA3IDk2LjcxIDU4NS4wNSA5Ni43MUw1ODQuNjggOTYuNzFDNTgxLjY3IDk2LjcxIDU3OC45OCA5Ny40NCA1NzYuNjIgOTguOSA1NzQuMzQgMTAwLjM3IDU3Mi41NSAxMDIuNDEgNTcxLjI0IDEwNS4wMSA1NzAuMDIgMTA3LjUzIDU2OS40MSAxMTAuNDYgNTY5LjQxIDExMy44MSA1NjkuNDEgMTE4Ljk0IDU3MC44OCAxMjMuMjEgNTczLjgxIDEyNi42MyA1NzYuNzQgMTI5Ljk3IDU4MS4zNCAxMzEuNjQgNTg3LjYxIDEzMS42NEw2MDcuNzcgMTMxLjY0IDYwNy43NyAxNDMuMTIgNTg3Ljc0IDE0My4xMlpNNTYyLjE5IDEyMC4wNSA1NjIuMTkgMTA4LjgxIDYxMS4xNyAxMDguODEgNjEyLjc2IDExNi45OSA2MTIuNzYgMTIwLjA1IDU2Mi4yIDEyMC4wNVoiLz48L2c+PGc+PHBhdGggZD0iTTM4OS4yIDE3NC4yNEMzODcuMDMgMTc0LjI0IDM4NS4wOSAxNzQuNTUgMzgzLjM0IDE3NS4wNyAzODAuNDcgMTc1LjY1IDM3Ny44IDE3Ni42NyAzNzUuMzMgMTc4LjEzIDM3NC44NyAxNzguNCAzNzQuNDYgMTc4LjczIDM3NC4wMiAxNzkuMDNMMzc0LjAyIDE3NS41NyAzNjEuODcgMTc1LjU3IDM2MS44NyAyMjcuMTIgMzc0LjIxIDIyNy4xMiAzNzQuMjEgMjAzLjI1QzM3NC4zMiAyMDIuNTggMzc0LjM5IDIwMS45NiAzNzQuMzkgMjAxLjQ1IDM3NC4zOSAxOTguMjQgMzc0Ljk4IDE5NS40MyAzNzYuMTUgMTkzLjAyIDM3Ny40IDE5MC41MiAzNzkuMTIgMTg4LjU3IDM4MS4zMSAxODcuMTYgMzgyLjMgMTg2LjU1IDM4My4zNSAxODYuMSAzODQuNDYgMTg1Ljc1IDM4NS4yIDE4NS42MiAzODUuOTYgMTg1LjU0IDM4Ni43NSAxODUuNTQgMzkzLjU5IDE4NS41NCAzOTguMDUgMTkwLjQ4IDM5OC4wNSAxOTkuMTJMMzk4LjA1IDIyNy4xMyA0MTAuNDkgMjI3LjEzIDQxMC40OSAxOTUuNzFDNDEwLjQ5IDE4OC43OCA0MDguNSAxODMuNTYgNDA0LjYxIDE3OS44NiA0MDAuNzIgMTc2LjE2IDM5NC44MiAxNzQuMjYgMzg5LjIyIDE3NC4yNloiLz48cGF0aCBkPSJNNTM1LjYyIDE4NS43M0M1NDEuMzIgMTg1LjczIDU0Ni4yNSAxODguODYgNTQ5LjI5IDE5My44OUw1NTkuMzUgMTg4LjQ4QzU1NS4xNyAxNzkuNDYgNTQ1Ljk2IDE3NC4zNCA1MzUuNDMgMTc0LjM0IDUyNy4yNyAxNzQuMzQgNTIwLjYyIDE3Ni45IDUxNS40OSAxODEuOTMgNTEwLjM2IDE4Ni45NiA1MDcuOCAxOTMuNDIgNTA3LjggMjAxLjMgNTA3LjggMjA5LjE4IDUxMC4zNiAyMTUuNjQgNTE1LjM5IDIyMC42NyA1MjAuNTIgMjI1LjcgNTI3LjA3IDIyOC4yNiA1MzUuMDQgMjI4LjI2IDU0NS41OCAyMjguMjYgNTU0LjY5IDIyMi45NCA1NTkuMTUgMjE0LjY4TDU0OS4wOSAyMDguNzlDNTQ2LjUzIDIxMy42MyA1NDAuOTMgMjE2Ljg2IDUzNS4yMyAyMTYuODYgNTMwLjg2IDIxNi44NiA1MjcuMjYgMjE1LjM0IDUyNC40MSAyMTIuNCA1MjEuNTYgMjA5LjQ2IDUyMC4xNCAyMDUuNzUgNTIwLjE0IDIwMS4yOSA1MjAuMTQgMTk2LjgzIDUyMS41NiAxOTMuMTMgNTI0LjUxIDE5MC4xOCA1MjcuNDUgMTg3LjI0IDUzMS4xNiAxODUuNzIgNTM1LjYyIDE4NS43MloiLz48cGF0aCBkPSJNNTkxLjAzIDE3NC4yNEM1ODkuMDIgMTc0LjI0IDU4Ny4yIDE3NC41MSA1ODUuNTQgMTc0Ljk3IDU4Mi40NiAxNzUuNTIgNTc5LjYxIDE3Ni41OCA1NzYuOTkgMTc4LjE0IDU3Ni41NiAxNzguNCA1NzYuMTcgMTc4LjcxIDU3NS43NiAxNzguOThMNTc1Ljc2IDE1Ny4zNiA1NjMuNyAxNTcuMzYgNTYzLjcgMjI3LjEyIDU3Ni4wNCAyMjcuMTIgNTc2LjA0IDIwMS41OEM1NzYuMDQgMjAxLjU4IDU3Ni4wNCAyMDEuNDkgNTc2LjA0IDIwMS40NSA1NzYuMDQgMTk4LjI0IDU3Ni42MyAxOTUuNDMgNTc3LjggMTkzLjAyIDU3OS4wNSAxOTAuNTIgNTgwLjc3IDE4OC41NyA1ODIuOTYgMTg3LjE2IDU4My45MyAxODYuNTYgNTg0Ljk4IDE4Ni4xMSA1ODYuMDcgMTg1Ljc3IDU4Ni44NCAxODUuNjMgNTg3LjY0IDE4NS41NSA1ODguNDggMTg1LjU1IDU5NS4zMiAxODUuNTUgNTk5Ljc4IDE5MC4zOSA1OTkuNzggMTk5LjAzTDU5OS43OCAyMjcuMTMgNjEyLjEyIDIyNy4xMyA2MTIuMTIgMTk1LjcxQzYxMi4xMiAxODguNzggNjEwLjEzIDE4My41NiA2MDYuMjMgMTc5Ljg2IDYwMi4zNCAxNzYuMTYgNTk2LjYzIDE3NC4yNiA1OTEuMDMgMTc0LjI2WiIvPjxwYXRoIGQ9Ik0zNDcuMzUgMTU2LjE4QzM0NS41NyAxNTYuMTggMzQ0LjA1IDE1Ni44MSAzNDIuOCAxNTguMDcgMzQxLjU0IDE1OS4yNyAzNDAuOTEgMTYwLjc5IDM0MC45MSAxNjIuNjIgMzQwLjkxIDE2NC40NSAzNDEuNTQgMTY1Ljk1IDM0Mi44IDE2Ny4yNiAzNDQuMDYgMTY4LjUyIDM0NS41NyAxNjkuMTUgMzQ3LjM1IDE2OS4xNSAzNDkuMTMgMTY5LjE1IDM1MC42NSAxNjguNTIgMzUxLjkxIDE2Ny4yNiAzNTMuMjIgMTY1Ljk1IDM1My44NyAxNjQuNCAzNTMuODcgMTYyLjYyIDM1My44NyAxNjAuODQgMzUzLjI1IDE1OS4zMiAzNTEuOTkgMTU4LjA3IDM1MC43MyAxNTYuODEgMzQ5LjE4IDE1Ni4xOCAzNDcuMzUgMTU2LjE4WiIvPjxwYXRoIGQ9Ik0zMTQuMDUgMTU2LjAyQzMxMS4xNyAxNTcuNTcgMzA4LjgxIDE1OS41OCAzMDcuMTggMTYyLjMxIDMwNS41NiAxNjQuOTYgMzA0LjUxIDE2OC4zNSAzMDQuNTEgMTcxLjc1TDMwNC41MSAxNzUuNTQgMjg5LjMgMTc1LjU0IDI4OS4zIDE4NS45NCAzMDQuNTEgMTg1Ljk0IDMwNC41MSAyMjcuMDcgMzE2LjYyIDIyNy4wNyAzMTYuNjIgMTg1Ljk0IDMzMi43OSAxODUuOTQgMzMyLjc5IDE3NS41NCAzMTYuNzEgMTc1LjU0IDMxNi43MSAxNzIuNDJDMzE2LjcxIDE3MC4xNCAzMTcuMzMgMTY4LjQ1IDMxOC44MSAxNjYuOSAzMjAuMzYgMTY1LjI4IDMyMS45MSAxNjQuOSAzMjQuMjcgMTY0LjlMMzMyLjc5IDE2NC45IDMzMi43OSAxNTMuNDYgMzIzLjkzIDE1My41NkMzMjAuMzkgMTUzLjU2IDMxNi45MSAxNTQuNDkgMzE0LjA0IDE1Ni4wNFoiLz48cGF0aCBkPSJNMzQxLjE5IDE4Ni4wNyAzNDEuMTkgMjE2LjU3IDM0MS4xOSAyMjcuMTEgMzUzLjQ4IDIyNy4xMSAzNTMuNDggMjE2LjU3IDM1My40OCAxNzUuNTQgMzQxLjE5IDE3NS41NCAzNDEuMTkgMTg2LjA3WiIvPjxwYXRoIGQ9Ik00MjkuMzMgMjA2LjU2IDQyOS4zMyAxODUuNzYgNDQ2LjY1IDE4NS43NiA0NDYuNjUgMTc1LjM4IDQyOS4zNiAxNzUuMzggNDI5LjM2IDE1Ny4zNSA0MTcuMjIgMTU3LjM1IDQxNy4yMiAyMDcuNThDNDE3LjIyIDIxOS44MiA0MjYuNzEgMjI4LjY2IDQzOS42NCAyMjguNjZMNDQ4Ljc3IDIyOC42NiA0NDguNzcgMjE3LjcxIDQ0MC44MSAyMTcuOEM0MzMuMjkgMjE3LjggNDI5LjQ4IDIxMy44MyA0MjkuMzQgMjA2LjU1WiIvPjxwYXRoIGQ9Ik00ODAuMzYgMjI4LjQ4QzQ3NC43NSAyMjguNDggNDY5Ljc2IDIyNy4zMyA0NjUuMzggMjI1LjAyIDQ2MSAyMjIuNzIgNDU3LjU4IDIxOS41MyA0NTUuMTIgMjE1LjQ2IDQ1Mi43NCAyMTEuMzEgNDUxLjU1IDIwNi42NiA0NTEuNTUgMjAxLjUyIDQ1MS41NSAxOTYuMzggNDUyLjYzIDE5MS43MyA0NTQuNzggMTg3LjU4IDQ1NyAxODMuNDMgNDYwLjA4IDE4MC4yIDQ2My45OSAxNzcuODkgNDY3Ljk5IDE3NS41MSA0NzIuNTIgMTc0LjMyIDQ3Ny41OSAxNzQuMzJMNDc3LjgyIDE3NC4zMkM0ODIuODkgMTc0LjMyIDQ4Ny4zOCAxNzUuNTEgNDkxLjMgMTc3Ljg5IDQ5NS4yOSAxODAuMiA0OTguNDEgMTgzLjQyIDUwMC42MyAxODcuNTggNTAyLjg2IDE5MS42NSA1MDMuOTcgMTk2LjI2IDUwMy45NyAyMDEuNEw1MDMuOTcgMjA2LjQ3IDQ5Mi40NSAyMDYuNDcgNDkyLjQ1IDIwMi41NkM0OTIuNDUgMTk4Ljc5IDQ5MS44NCAxOTUuNTcgNDkwLjYxIDE5Mi44OCA0ODkuMzggMTkwLjE5IDQ4Ny42NSAxODguMTUgNDg1LjQyIDE4Ni43NyA0ODMuMiAxODUuMzkgNDgwLjY2IDE4NC43IDQ3Ny44MiAxODQuN0w0NzcuNDcgMTg0LjdDNDc0LjYzIDE4NC43IDQ3Mi4wOSAxODUuMzkgNDY5Ljg3IDE4Ni43NyA0NjcuNzIgMTg4LjE2IDQ2Ni4wMyAxOTAuMDggNDY0Ljc5IDE5Mi41NCA0NjMuNjQgMTk0LjkyIDQ2My4wNiAxOTcuNjkgNDYzLjA2IDIwMC44NCA0NjMuMDYgMjA1LjY4IDQ2NC40NSAyMDkuNzEgNDY3LjIxIDIxMi45NCA0NjkuOTggMjE2LjA5IDQ3NC4zMiAyMTcuNjcgNDgwLjIzIDIxNy42N0w0OTkuMjUgMjE3LjY3IDQ5OS4yNSAyMjguNSA0ODAuMzUgMjI4LjVaTTQ1Ni4yOCAyMDYuNyA0NTYuMjggMTk2LjEgNTAyLjQ5IDE5Ni4xIDUwMy45OSAyMDMuODIgNTAzLjk5IDIwNi43IDQ1Ni4yOSAyMDYuN1oiLz48L2c+PC9nPjwvc3ZnPg==';
const LOGO_MARK_B64  = 'PHN2ZyB2aWV3Qm94PSIwIDAgMzE0LjUxIDMxNC43IiB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHhtbG5zOnhsaW5rPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5L3hsaW5rIiBpZD0iTGF5ZXJfMSIgb3ZlcmZsb3c9ImhpZGRlbiI+PHBhdGggZD0iTTE3MC41OCA2MC41MyAxNTYuNzggOTcuODEgMTU2Ljc4IDk3LjgxQzE1Ni43OCA5Ny44MSAxNDMuMyAxMzQuMjMgMTQzLjMgMTM0LjIzTDExNi45NiAxMzAuNzcgMTE4LjQxIDEyNi44NyAxNDIuMTUgNjIuN0MxNDMuMDcgNjAuMjEgMTQ1LjMzIDU4LjQ1IDE0Ny45OCA1OC4yMiAxNTAuODcgNTcuOTcgMTUzLjggNTcuODQgMTU2Ljc3IDU3Ljg0IDE2MC41NiA1Ny44NCAxNjQuMjkgNTguMDUgMTY3Ljk3IDU4LjQ2IDE2OS4xNiA1OC41OSAxNzAuMTcgNTkuNCAxNzAuNTggNjAuNTNMMTcwLjU4IDYwLjUzWiIgZmlsbD0iI0M0RkYwMCIvPjxwYXRoIGQ9Ik0xOTUuMTQgMTI2Ljg5IDE4Ny4yNCAxNDAuMDQgMTcxLjY0IDEzNy45OCAxNTYuNzggOTcuODMgMTU2Ljc4IDk3LjgyIDE3MC41OCA2MC41MyAxOTUuMTQgMTI2Ljg5WiIvPjxwYXRoIGQ9Ik0yNTYuNyAxNTcuNzZDMjU2LjcgMTYxLjc3IDI1Ni40NiAxNjUuNzIgMjU2IDE2OS42IDI1NS42NCAxNzIuNjUgMjUyLjggMTc0Ljc5IDI0OS43NiAxNzQuMzlMMjEwLjgyIDE2OS4yNyAyMDAuNzEgMTY3LjkzIDIwNi42NSAxNTguMDIgMjE1LjIyIDE0My43MyAyNTEuNzggMTQ4LjU0QzI1NC40NiAxNDguODkgMjU2LjUyIDE1MS4xMiAyNTYuNjIgMTUzLjgyIDI1Ni42NyAxNTUuMTMgMjU2LjcgMTU2LjQ1IDI1Ni43IDE1Ny43N1oiIGZpbGw9IiM4NzVFRkEiLz48cGF0aCBkPSJNMjI4LjggMjI3LjAzQzIyNC45NSAyMzEuMDQgMjIwLjc2IDIzNC43MyAyMTYuMjggMjM4LjA1IDIxMy4yOSAyNDAuMjYgMjA5LjAyIDIzOSAyMDcuNzMgMjM1LjUxTDE4OS41OSAxODYuNDcgMjAwLjcxIDE2Ny45MiAyMTAuODIgMTY5LjI2IDIzMC4wNCAyMjEuMTlDMjMwLjc5IDIyMy4yMSAyMzAuMjkgMjI1LjQ3IDIyOC44IDIyNy4wMloiLz48cGF0aCBkPSJNMTE4LjQxIDEyNi44NyAxMTYuOTYgMTMwLjc3IDg4LjI3IDEyNi45OSA3NS4zOSAxMDUuNTNDNzQuMjQgMTAzLjYxIDc0LjM2IDEwMS4xOSA3NS42NyA5OS4zNyA3OC44OCA5NC45MSA4Mi40NSA5MC43MyA4Ni4zNCA4Ni44NyA4OC45MSA4NC4zMiA5My4xOSA4NC44NCA5NS4wNiA4Ny45NEwxMTguNDEgMTI2Ljg2WiIgZmlsbD0iIzJBNjJGNSIvPjxwYXRoIGQ9Ik0xMzQuMDggMTU5LjE0IDEyMy45NiAxODYuNDkgMTA1LjgyIDIzNS41MUMxMDQuNTMgMjM5IDEwMC4yNSAyNDAuMjYgOTcuMjcgMjM4LjA1IDkyLjc5IDIzNC43MyA4OC42IDIzMS4wMyA4NC43NCAyMjcuMDIgODMuMjUgMjI1LjQ3IDgyLjc1IDIyMy4yMSA4My41IDIyMS4xOUwxMDYuODggMTU4LjAyIDEwNy43NCAxNTUuNjggMTM0LjA4IDE1OS4xNFoiIGZpbGw9IiNDNEZGMDAiLz48cGF0aCBkPSJNMTg2LjE3IDI1My4zMUMxODAuOTggMjU0LjkxIDE3NS42MSAyNTYuMDkgMTcwLjEgMjU2LjgyIDE2Ny44OCAyNTcuMTEgMTY1LjcgMjU2LjA4IDE2NC41NSAyNTQuMTdMMTU2Ljc4IDI0MS4yMiAxNTYuNzggMjQxLjIyQzE1Ni43OCAyNDEuMjIgMTQxLjY4IDIxNi4wNCAxNDEuNjggMjE2LjA0TDEyMy45NyAxODYuNTEgMTM0LjA5IDE1OS4xNiAxMzguMDggMTU5LjY5IDE1Ni43OSAxOTAuODYgMTcxLjg4IDIxNi4wNCAxODkuMzMgMjQ1LjEzQzE5MS4xOSAyNDguMjQgMTg5LjY0IDI1Mi4yNiAxODYuMTggMjUzLjMzWiIgZmlsbD0iIzJBNjJGNSIvPjxwYXRoIGQ9Ik0xNTYuNzcgMjQxLjE5IDE1Ni43NyAyNDEuMTlDMTU2Ljc3IDI0MS4xOSAxNDkuMDEgMjU0LjE1IDE0OS4wMSAyNTQuMTUgMTQ3Ljg2IDI1Ni4wNyAxNDUuNjcgMjU3LjEgMTQzLjQ2IDI1Ni44MSAxMzcuOTUgMjU2LjA4IDEzMi41OCAyNTQuOSAxMjcuMzkgMjUzLjMgMTIzLjkzIDI1Mi4yNCAxMjIuMzcgMjQ4LjIxIDEyNC4yNCAyNDUuMTFMMTQxLjY4IDIxNi4wMiAxNTYuNzggMjQxLjE5WiIgZmlsbD0iIzFEQ0Q4NiIvPjxwYXRoIGQ9Ik0yMzcuODYgOTkuMzZDMjM5LjE3IDEwMS4xOCAyMzkuMjkgMTAzLjU5IDIzOC4xNCAxMDUuNTFMMjE1LjIzIDE0My43MiAyMDYuNjYgMTU4LjAxIDIwMC43MiAxNjcuOTIgMTg5LjYgMTg2LjQ3IDE3MS44NyAyMTYuMDIgMTU2Ljc4IDE5MC44NCAxNzIuNzMgMTY0LjI0IDE3OC4wNiAxNTUuMzUgMTg3LjI0IDE0MC4wNCAxOTUuMTMgMTI2Ljg5IDIxOC40NyA4Ny45NkMyMjAuMzMgODQuODUgMjI0LjYxIDg0LjMzIDIyNy4xOSA4Ni44OCAyMzEuMDcgOTAuNzMgMjM0LjY0IDk0LjkxIDIzNy44NSA5OS4zNloiIGZpbGw9IiMxRENEODYiLz48cGF0aCBkPSJNMTg3LjI0IDE0MC4wNCAxNzguMDYgMTU1LjM1IDE3Mi43MyAxNjQuMjQgMTM4LjA5IDE1OS42OCAxMzguMDggMTU5LjY4IDEzNC4wOSAxNTkuMTUgMTA3Ljc1IDE1NS42OSA2Mi42NiAxNDkuNzRDNTkuNjIgMTQ5LjM0IDU3LjQzIDE0Ni41NCA1Ny44NyAxNDMuNSA1OC42MiAxMzguMjIgNTkuNzkgMTMzLjA2IDYxLjM0IDEyOC4wOCA2Mi4xNCAxMjUuNSA2NC43IDEyMy44OSA2Ny4zOCAxMjQuMjRMODguMjggMTI3IDExNi45NyAxMzAuNzggMTQzLjMxIDEzNC4yNCAxNzEuNjUgMTM3Ljk4IDE4Ny4yNSAxNDAuMDRaIiBmaWxsPSIjODc1RUZBIi8+PC9zdmc+';

fs.mkdirSync('/home/claude/assets', { recursive: true });
fs.writeFileSync('/home/claude/assets/image1.svg', Buffer.from(LOGO_FULL_B64,  'base64'));
fs.writeFileSync('/home/claude/assets/image2.svg', Buffer.from(LOGO_MARK_B64,  'base64'));
console.log('✅ Logos written to /home/claude/assets/');
```

Photos are loaded from GitHub raw URLs — no local setup needed for photos.

---

## How to Use This Skill

1. Read this file completely before writing any code.
2. Run the **Setup block** above to write logo files to disk (once per session).
3. Copy the **Boilerplate** section verbatim as your script starting point.
4. Add slides using the **Slide Recipes** — each is a drop-in code block.
5. Replace placeholder text/data with the user's actual content.
6. Save to `/mnt/user-data/outputs/<filename>.pptx` and run with `node`.
7. Convert to images and visually QA before presenting the file.

**Always produce an editable `.pptx` file. Never flatten to PDF or image.**

---

## Asset Paths

```
LOGO_FULL  = /home/claude/assets/image1.svg   ← full wordmark (title & thank-you slides)
LOGO_MARK  = /home/claude/assets/image2.svg   ← X mark icon (top-right, all other slides)
```

Photos load from GitHub raw URLs (no download needed):

```
GITHUB_RAW = https://raw.githubusercontent.com/TheLeroyMeyer/weaver-fintech-skill/main/assets/
```

⚠️ Replace `YOUR_USERNAME` with your actual GitHub username before running any script.

---

## Image Sizing Rules — NEVER STRETCH OR WARP PHOTOS

Always derive render width from the source pixel aspect ratio at h = 7.5":
  `render_w = (pixel_w / pixel_h) × 7.5`
Then right-anchor: `x = 13.33 - render_w`

Pre-calculated render widths at h = 7.5":

| File        | Pixel dimensions | render_w |
|-------------|-----------------|----------|
| image5.png  | 1485 × 1124     | 9.91"    |
| image6.png  | 1033 × 1125     | 6.89"    |
| image7.png  | 1033 × 1125     | 6.89"    |
| image8.png  | 1165 × 1093     | 7.99"    |
| image9.png  |  925 × 1009     | 6.88"    |
| image10.png | 1106 × 1124     | 7.38"    |

**Never use `sizing: { type: 'contain' }` or any auto-fit — always use explicit w/h from the table above.**

---

## Brand Constants

```javascript
const C = {
  purple:    '875EFA',  // titles, left accent bars, action cards, numbered circles
  lime:      'C4FF00',  // bottom stripe on thank-you, accent tags
  green:     '1DCD86',  // positive deltas, on-track indicators
  blue:      '2A62F5',  // secondary chart series
  black:     '1A1A1A',  // body text, slide numbers, chart axes
  white:     'FFFFFF',
  offwhite:  'F4F4F4',  // chart panel backgrounds, evidence panel fill
  grey:      'DDDDDD',  // chart panel borders, divider lines
  insightbg: 'EFEFEF',  // insight bar background (light grey)
};
const TF = 'Aptos Display';  // title font
const BF = 'Aptos';          // body font
```

---

## Boilerplate (start every script here)

```javascript
const PptxGenJS = require('pptxgenjs');
const fs        = require('fs');
const pptx      = new PptxGenJS();
pptx.layout     = 'LAYOUT_WIDE'; // 13.33" × 7.5"

const C = {
  purple:'875EFA', lime:'C4FF00', green:'1DCD86', blue:'2A62F5',
  black:'1A1A1A', white:'FFFFFF', offwhite:'F4F4F4', grey:'DDDDDD', insightbg:'EFEFEF',
};
const TF = 'Aptos Display';
const BF = 'Aptos';

// ── ASSETS ───────────────────────────────────────────────────────────────────
const LOGO      = '/home/claude/assets/image1.svg';  // written by setup.js
const LOGO_MARK = '/home/claude/assets/image2.svg';  // written by setup.js

// ⚠️ Replace YOUR_USERNAME with your GitHub username
const GH = 'https://raw.githubusercontent.com/YOUR_USERNAME/weaver-fintech-skill/main/assets/';

const PHOTOS = {
  title: { path: GH+'image5.png',  w: 9.91 },
  div1:  { path: GH+'image6.png',  w: 6.89 },
  div2:  { path: GH+'image7.png',  w: 6.89 },
  div3:  { path: GH+'image8.png',  w: 7.99 },
  div4:  { path: GH+'image9.png',  w: 6.88 },
  div5:  { path: GH+'image10.png', w: 7.38 },
};

// ── HELPERS ──────────────────────────────────────────────────────────────────

function logoMark(s) {
  // X mark icon — top-right, every slide EXCEPT first and last
  s.addImage({ path: LOGO_MARK, x: 12.55, y: 0.08, w: 0.62, h: 0.62 });
}

function slideNum(s, n) {
  s.addText(String(n), {
    x: 12.8, y: 7.25, w: 0.4, h: 0.2,
    fontSize: 9, bold: true, color: C.black, align: 'right', fontFace: BF
  });
}

function slideTitle(s, title, subtitle) {
  s.addText(title, {
    x: 0.35, y: 0.08, w: 11.7, h: subtitle ? 0.44 : 0.52,
    fontSize: 20, bold: true, color: C.purple, fontFace: TF, wrap: true
  });
  if (subtitle) {
    s.addText(subtitle, {
      x: 0.35, y: 0.54, w: 11.7, h: 0.26,
      fontSize: 11, italic: true, color: C.black, fontFace: BF
    });
  }
}

function chartBox(s, x, y, w, h, title) {
  s.addShape(pptx.shapes.RECTANGLE, {
    x, y, w, h, fill: { color: C.offwhite }, line: { color: C.grey, pt: 1 }
  });
  s.addText(title, {
    x: x+0.14, y: y+0.1, w: w-0.28, h: 0.27,
    fontSize: 9, bold: true, color: C.black, fontFace: BF
  });
  s.addShape(pptx.shapes.RECTANGLE, {
    x: x+0.14, y: y+0.38, w: w-0.28, h: 0.012,
    fill: { color: 'CCCCCC' }, line: { color: 'CCCCCC' }
  });
}

function insightBar(s, headline, impl, watch, action) {
  const Y = 5.72, H = 1.78;
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0, y: Y, w: 13.33, h: H, fill: { color: C.insightbg }, line: { color: C.insightbg }
  });
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0, y: Y, w: 0.07, h: H, fill: { color: C.purple }, line: { color: C.purple }
  });
  s.addText('INSIGHT', {
    x: 0.2, y: Y+0.1, w: 2, h: 0.22,
    fontSize: 8, bold: true, color: C.purple, fontFace: BF, charSpacing: 2
  });
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.2, y: Y+0.33, w: 13.0, h: 0.012, fill: { color: 'CCCCCC' }, line: { color: 'CCCCCC' }
  });
  s.addText(headline, {
    x: 0.2, y: Y+0.38, w: 13.0, h: 0.42,
    fontSize: 10.5, bold: true, color: C.black, fontFace: BF, wrap: true
  });
  [
    { label: 'Strategic implication', body: impl },
    { label: 'Risk or watch point',   body: watch },
    { label: 'Recommended action',    body: action },
  ].forEach(({ label, body }, i) => {
    const cx = 0.2 + i * 4.35;
    s.addText(label, {
      x: cx, y: Y+0.88, w: 4.1, h: 0.2,
      fontSize: 8.5, bold: true, color: C.black, fontFace: BF
    });
    s.addText(body, {
      x: cx, y: Y+1.1, w: 4.1, h: 0.58,
      fontSize: 7.5, color: C.black, fontFace: BF, wrap: true
    });
  });
}

// ── SLIDES GO HERE ────────────────────────────────────────────────────────────
// ... use the Slide Recipes below ...

// ── SAVE ──────────────────────────────────────────────────────────────────────
pptx.writeFile({ fileName: '/mnt/user-data/outputs/FILENAME.pptx' })
  .then(() => console.log('✅ Done'))
  .catch(e => console.error(e));
```

---

## Slide Recipes

### TITLE SLIDE
**Rules:** Full wordmark top-left (`x:0.35, y:0.28, w:2.5, h:1.1`). NO logo mark top-right. Photo right-anchored at natural aspect ratio. No slide number.

```javascript
{
  const s = pptx.addSlide();
  s.addImage({ path: LOGO, x: 0.35, y: 0.28, w: 2.5, h: 1.1 });
  s.addText('Presentation Title\nLine Two\nLine Three', {
    x: 0.35, y: 2.0, w: 7.2, h: 2.9,
    fontSize: 44, bold: true, color: C.purple, fontFace: TF, wrap: true
  });
  s.addText('Subheading or one-line descriptor.', {
    x: 0.35, y: 5.1, w: 7.0, h: 0.38,
    fontSize: 15, bold: true, color: C.black, fontFace: BF
  });
  s.addText('Month Year', { x: 0.35, y: 5.58, w: 4, h: 0.3, fontSize: 13, color: C.black, fontFace: BF });
  s.addText('weaverfintech.com', { x: 0.35, y: 5.95, w: 4, h: 0.28, fontSize: 12, color: C.black, fontFace: BF });
  // Photo — natural aspect ratio, right-anchored — NEVER STRETCH
  const p = PHOTOS.title;
  s.addImage({ path: p.path, x: 13.33 - p.w, y: 0, w: p.w, h: 7.5 });
}
```

---

### DIVIDER SLIDE
**Rules:** Logo mark top-right. Large purple bold title, left side. Photo right-anchored at natural aspect ratio. Slide number bottom-right. No subtitle, no insight bar.
Use `fontSize: 46` for long titles (11+ chars per line), `fontSize: 52` for short titles.
Cycle photos: div1 → div2 → div3 → div4 → div5.

```javascript
{
  const s = pptx.addSlide();
  logoMark(s);
  s.addText('Section\nTitle.', {
    x: 0.35, y: 2.0, w: 5.8, h: 3.2,
    fontSize: 52, bold: true, color: C.purple, fontFace: TF, wrap: true
  });
  const p = PHOTOS.div1; // swap to div2/div3/div4/div5 as needed
  s.addImage({ path: p.path, x: 13.33 - p.w, y: 0, w: p.w, h: 7.5 });
  slideNum(s, N);
}
```

---

### TWO-CHART SLIDE (most common analytical layout)
**Rules:** Logo mark top-right. Purple title + italic subtitle. Two equal chart panels side-by-side. Insight bar at bottom. Slide number bottom-right.

```javascript
{
  const s = pptx.addSlide();
  logoMark(s);
  slideTitle(s,
    'Title states the finding — not just the topic.',
    'Section Label  |  Time Period or Context'
  );

  chartBox(s, 0.2, 0.88, 6.4, 4.72, 'Chart A Title');
  s.addChart(pptx.charts.BAR, [
    { name: 'Series 1', labels: ['Cat1','Cat2','Cat3','Cat4','Cat5'], values: [10,20,30,40,50] },
    { name: 'Series 2', labels: ['Cat1','Cat2','Cat3','Cat4','Cat5'], values: [5,15,25,35,45] },
  ], {
    x: 0.25, y: 1.32, w: 6.3, h: 4.22,
    barGrouping: 'clustered',
    chartColors: ['1A1A1A', C.purple],
    showLegend: true, legendPos: 'b', legendFontSize: 8,
    showValue: true, dataLabelFontSize: 8,
    catAxisLabelFontSize: 9, showTitle: false,
    plotAreaFill: { color: C.offwhite },
  });

  chartBox(s, 6.73, 0.88, 6.4, 4.72, 'Chart B Title');
  s.addChart(pptx.charts.LINE, [
    { name: 'Series A',  labels: ['Cat1','Cat2','Cat3','Cat4','Cat5'], values: [10,20,30,40,50] },
    { name: 'Benchmark', labels: ['Cat1','Cat2','Cat3','Cat4','Cat5'], values: [15,15,20,22,25] },
  ], {
    x: 6.78, y: 1.32, w: 6.3, h: 4.22,
    chartColors: [C.purple, '1A1A1A'],
    showLegend: true, legendPos: 'b', legendFontSize: 8,
    showValue: true, dataLabelFontSize: 8,
    catAxisLabelFontSize: 9, showTitle: false,
    plotAreaFill: { color: C.offwhite },
    lineDataSymbol: 'circle',
  });

  insightBar(s,
    'Key insight — one sentence that tells the executive what the data means and demands.',
    'Strategic implication — describe the consequence or opportunity.',
    'Risk or watch point — flag the dependency or risk.',
    'Recommended action — state the specific next step or decision.'
  );
  slideNum(s, N);
}
```

---

### SINGLE FULL-WIDTH CHART SLIDE
**Rules:** One panel spans full width (`x:0.2, w:12.9`). Use when one chart tells the whole story.

```javascript
{
  const s = pptx.addSlide();
  logoMark(s);
  slideTitle(s, 'Title states the finding.', 'Section  |  Period');

  chartBox(s, 0.2, 0.88, 12.9, 4.72, 'Chart Title');
  s.addChart(pptx.charts.BAR, [
    { name: 'Series', labels: ['A','B','C','D','E','F'], values: [10,20,30,40,50,60] },
  ], {
    x: 0.25, y: 1.32, w: 12.8, h: 4.22,
    barGrouping: 'stacked', chartColors: [C.purple],
    showValue: true, dataLabelFontSize: 8, dataLabelColor: C.white,
    catAxisLabelFontSize: 9, showTitle: false,
    plotAreaFill: { color: C.offwhite }, valAxisHidden: true,
  });

  insightBar(s, 'Headline insight.', 'Implication.', 'Watch point.', 'Action.');
  slideNum(s, N);
}
```

---

### FOUR-CHART 2×2 GRID SLIDE

```javascript
{
  const s = pptx.addSlide();
  logoMark(s);
  slideTitle(s, 'All four metrics tell a coherent story.', 'Dashboard  |  Period');

  const yrs = ['FY22','FY23','FY24','FY25','FY26'];
  [
    { x: 0.2,  y: 0.88, w: 6.4, h: 2.28, title: 'Metric A', vals: [10,20,30,40,50] },
    { x: 6.73, y: 0.88, w: 6.4, h: 2.28, title: 'Metric B', vals: [5,8,12,18,25]  },
    { x: 0.2,  y: 3.28, w: 6.4, h: 2.28, title: 'Metric C', vals: [80,75,65,55,42] },
    { x: 6.73, y: 3.28, w: 6.4, h: 2.28, title: 'Metric D', vals: [3.1,3.8,5.2,7.4,9.1] },
  ].forEach(({ x, y, w, h, title, vals }) => {
    chartBox(s, x, y, w, h, title);
    s.addChart(pptx.charts.BAR, [{ name: title, labels: yrs, values: vals }], {
      x: x+0.05, y: y+0.44, w: w-0.1, h: h-0.5,
      barGrouping: 'stacked', chartColors: [C.purple],
      showValue: true, dataLabelFontSize: 7.5, dataLabelColor: C.white,
      catAxisLabelFontSize: 7.5, showTitle: false,
      plotAreaFill: { color: C.offwhite }, valAxisHidden: true,
    });
  });

  insightBar(s,
    'Headline that unifies all four charts into one conclusion.',
    'Top row takeaway — what Chart A & B reveal together.',
    'Bottom row takeaway — what Chart C & D reveal together.',
    'Recommended action based on the combined picture.'
  );
  slideNum(s, N);
}
```

---

### STAT CARDS ROW (5 headline numbers)
**Rules:** Alternating purple / black cards. Large bold number, small label below.

```javascript
[
  { num: '9.1×', lbl: 'Stat label line one\nline two context', color: C.purple },
  { num: 'R18',  lbl: 'Stat label line one\nline two context', color: '1A1A1A' },
  { num: '95%',  lbl: 'Stat label line one\nline two context', color: C.purple },
  { num: '84%',  lbl: 'Stat label line one\nline two context', color: '1A1A1A' },
  { num: '24%',  lbl: 'Stat label line one\nline two context', color: C.purple },
].forEach((st, i) => {
  const xx = 0.2 + i * 2.6;
  s.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: xx, y: 0.95, w: 2.48, h: 2.0,
    fill: { color: st.color }, line: { color: st.color }, rectRadius: 0.1
  });
  s.addText(st.num, {
    x: xx+0.12, y: 1.05, w: 2.24, h: 0.9,
    fontSize: 30, bold: true,
    color: st.color === C.purple ? C.white : C.purple,
    fontFace: TF, align: 'center'
  });
  s.addText(st.lbl, {
    x: xx+0.1, y: 1.98, w: 2.28, h: 0.8,
    fontSize: 8.5, color: C.white,
    fontFace: BF, align: 'center', wrap: true
  });
});
```

---

### RECOMMENDATION SLIDE (evidence panel + 4 action cards)
**Rules:** Smaller title (16pt). Left panel = evidence list in offwhite box. Right = 2×2 grid of action cards alternating purple/black. Insight bar at bottom.

```javascript
{
  const s = pptx.addSlide();
  logoMark(s);

  s.addText('Recommendation title — keep under 120 characters.', {
    x: 0.35, y: 0.05, w: 11.7, h: 0.55,
    fontSize: 16, bold: true, color: C.purple, fontFace: TF, wrap: true
  });
  s.addText('Section Label  |  Context', {
    x: 0.35, y: 0.62, w: 11.7, h: 0.24,
    fontSize: 11, italic: true, color: C.black, fontFace: BF
  });

  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.2, y: 0.92, w: 6.0, h: 4.78,
    fill: { color: C.offwhite }, line: { color: C.grey, pt: 1 }
  });
  s.addText('The Evidence Case', {
    x: 0.35, y: 1.0, w: 5.7, h: 0.28,
    fontSize: 10, bold: true, color: C.black, fontFace: BF
  });
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.35, y: 1.3, w: 5.7, h: 0.012,
    fill: { color: 'CCCCCC' }, line: { color: 'CCCCCC' }
  });

  ['Evidence point 1.','Evidence point 2.','Evidence point 3.','Evidence point 4.','Evidence point 5.']
    .forEach((txt, i) => {
      const yy = 1.42 + i * 0.84;
      s.addShape(pptx.shapes.OVAL, {
        x: 0.35, y: yy+0.04, w: 0.34, h: 0.34,
        fill: { color: '1A1A1A' }, line: { color: '1A1A1A' }
      });
      s.addText(String(i+1).padStart(2,'0'), {
        x: 0.35, y: yy+0.04, w: 0.34, h: 0.34,
        fontSize: 7.5, bold: true, color: C.white, align: 'center', fontFace: BF
      });
      s.addText(txt, {
        x: 0.82, y: yy, w: 5.2, h: 0.72,
        fontSize: 8.5, color: C.black, fontFace: BF, wrap: true
      });
    });

  [
    { n:'1', title:'Action One',   body:'Detail text.', color: C.purple },
    { n:'2', title:'Action Two',   body:'Detail text.', color: '1A1A1A' },
    { n:'3', title:'Action Three', body:'Detail text.', color: C.purple },
    { n:'4', title:'Action Four',  body:'Detail text.', color: '1A1A1A' },
  ].forEach((a, i) => {
    const xx = 6.4 + (i%2)*3.45;
    const yy = 0.92 + Math.floor(i/2)*2.42;
    s.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: xx, y: yy, w: 3.3, h: 2.2,
      fill: { color: a.color }, line: { color: a.color }, rectRadius: 0.1
    });
    s.addShape(pptx.shapes.OVAL, {
      x: xx+0.15, y: yy+0.12, w: 0.46, h: 0.46,
      fill: { color: a.color===C.purple ? C.white : '875EFA' },
      line: { color: a.color===C.purple ? C.white : '875EFA' }
    });
    s.addText(a.n, {
      x: xx+0.15, y: yy+0.12, w: 0.46, h: 0.46,
      fontSize: 13, bold: true,
      color: a.color===C.purple ? C.purple : C.white,
      align: 'center', fontFace: TF
    });
    s.addText(a.title, {
      x: xx+0.15, y: yy+0.68, w: 2.98, h: 0.44,
      fontSize: 9.5, bold: true, color: C.white, fontFace: TF, wrap: true
    });
    s.addText(a.body, {
      x: xx+0.15, y: yy+1.16, w: 2.98, h: 0.92,
      fontSize: 8, color: a.color===C.purple ? C.white : 'CCCCCC',
      fontFace: BF, wrap: true
    });
  });

  insightBar(s,
    'Headline recommendation — one sentence, maximum 25 words.',
    'Strategic implication — what this means for the business.',
    'Risk or watch point — what could go wrong or delay.',
    'Recommended action — the specific decision or next step.'
  );
  slideNum(s, N);
}
```

---

### THANK YOU SLIDE
**Rules:** Full wordmark top-left. NO logo mark top-right. No slide number. Lime stripe at very bottom.

```javascript
{
  const s = pptx.addSlide();
  s.addImage({ path: LOGO, x: 0.35, y: 0.28, w: 2.5, h: 1.1 });
  s.addText('Thank you.', {
    x: 0, y: 2.7, w: 13.33, h: 1.8,
    fontSize: 64, bold: true, color: C.purple, align: 'center', fontFace: TF
  });
  s.addText('weaverfintech.com', {
    x: 0, y: 6.7, w: 13.33, h: 0.38,
    fontSize: 12, color: C.black, align: 'center', fontFace: BF
  });
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0, y: 7.28, w: 13.33, h: 0.22,
    fill: { color: C.lime }, line: { color: C.lime }
  });
}
```

---

## Slide Structure Rules

| Rule | Detail |
|------|--------|
| Logo mark | Top-right (`x:12.55, y:0.08, w:0.62, h:0.62`) on every slide **except** first and last |
| Full wordmark | Top-left (`x:0.35, y:0.28, w:2.5, h:1.1`) on first and last slides only |
| Slide number | Bottom-right on every slide except first and last |
| Title font | 20pt bold purple for content slides; 44pt for title slide; 52pt for dividers |
| Subtitle | 11pt italic black, directly below title |
| Photos | Always natural aspect ratio — never stretch. Use pre-calculated widths. Right-anchor: `x = 13.33 - render_w` |
| Insight bar | Every analytical content slide. `y:5.72, h:1.78`. Light grey bg + purple left accent |
| Chart panels | `chartBox()` helper always — offwhite fill, grey border, bold title, grey rule |
| Chart colours | Primary series: `875EFA` purple. Secondary: `1A1A1A` black. Third: `C4FF00` lime |
| Slide order | Title → Divider → Content × N → Divider → Content × N → ... → Thank You |

---

## Insight Bar Copy Rules

The insight bar must answer **"So what?"** — never just describe the data.

- **Headline** (bold, 10.5pt): One sentence, max 25 words. States what the data means + what it demands.
- **Strategic implication**: The consequence or opportunity this creates for the business.
- **Risk or watch point**: The dependency, constraint, or risk that could undermine the opportunity.
- **Recommended action**: The specific decision, next step, or escalation required.

---

## QA Checklist (run before delivering every deck)

```bash
npm install pptxgenjs
node setup.js          # write logos to disk (once per session)
node your_script.js    # generate the deck
```

Inspect every slide for:
- [ ] Logo mark top-right on slides 2 through N-1 (not first, not last)
- [ ] Full wordmark top-left on slide 1 and last slide only
- [ ] No photo stretching or warping
- [ ] Slide titles fit — no overflow, no truncation
- [ ] Insight bar on every analytical/chart slide
- [ ] Insight bar answers "So what?" not data narration
- [ ] Slide numbers bottom-right on all content slides
- [ ] All colours match brand hex exactly
- [ ] No placeholder text remaining
- [ ] Charts visible and labelled
- [ ] File saves as `.pptx`

---

## Common Errors to Avoid

| Error | Fix |
|-------|-----|
| Stretched/warped photos | Use pre-calculated `render_w` values. Never use `sizing: { type: 'contain' }` |
| Slide title overflows | Use `fontSize: 16` on recommendation slides; keep titles under 120 chars |
| Logo mark on title/thank-you | Only call `logoMark(s)` on slides 2 through N-1 |
| Logo mark missing on content slides | Call `logoMark(s)` at the top of every content/divider block |
| Dark insight bar | Use `C.insightbg` (`EFEFEF`) — light grey, not black |
| Divider title hyphenates | Shorten or use `fontSize: 46` |
| Inventing new colours | Only use the 9 colours in `C` |
| Missing slide numbers | Call `slideNum(s, N)` on every slide except first and last |
| Photos not found | Check GitHub raw URL — ensure repo is public and path is correct |
| Logos not found | Run `node setup.js` at the start of every session |
