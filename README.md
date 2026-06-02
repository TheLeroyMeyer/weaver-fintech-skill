# Weaver Fintech — Claude Presentation Skill

A self-contained Claude skill for generating on-brand Weaver Fintech `.pptx` presentations using `pptxgenjs`. Works on any Claude account once set up correctly.

---

## What's in this repo

```
weaver-fintech-skill/
├── SKILL.md               ← The Claude skill file (upload this to your account)
├── generate-pitch.js      ← Ready-to-run generic pitch deck script
├── assets/
│   ├── ASSET_SETUP.md     ← Instructions for adding the brand assets
│   ├── image1.svg         ← PLACEHOLDER: Full wordmark (add from template PPTX)
│   ├── image2.svg         ← PLACEHOLDER: X mark icon (add from template PPTX)
│   ├── image5.png         ← PLACEHOLDER: Title photo  1485×1124px
│   ├── image6.png         ← PLACEHOLDER: Divider photo 1033×1125px
│   ├── image7.png         ← PLACEHOLDER: Divider photo 1033×1125px
│   ├── image8.png         ← PLACEHOLDER: Divider photo 1165×1093px
│   ├── image9.png         ← PLACEHOLDER: Divider photo  925×1009px
│   └── image10.png        ← PLACEHOLDER: Divider photo 1106×1124px
└── README.md              ← This file
```

---

## Setup (one time per Claude account)

### Step 1 — Extract assets from the master template PPTX

1. Rename `Weaver_Master_Template.pptx` → `Weaver_Master_Template.zip`
2. Unzip it
3. Open the unzipped folder → `ppt/media/`
4. Copy these files into the `assets/` folder of this repo:
   - `image1.svg` (full wordmark)
   - `image2.svg` (X mark icon)
   - `image5.png` through `image10.png` (photos)

### Step 2 — Push to GitHub (or keep locally)

```bash
git add assets/
git commit -m "add brand assets"
git push
```

If the repo is **private**, use raw GitHub URLs with a personal access token.  
If the repo is **public**, raw GitHub URLs work directly (see SKILL.md for URL format).

### Step 3 — Add the skill to Claude

1. In Claude.ai → Settings → Skills → Add Skill
2. Upload or paste the contents of `SKILL.md`
3. Done — Claude will now use this skill whenever you ask for a Weaver Fintech presentation

---

## Generating a deck

Once the skill is installed, just ask Claude:

> *"Create a Weaver Fintech pitch deck for [topic/audience]"*

Claude will read the skill, run `generate-pitch.js` (or write a custom version), and output an editable `.pptx` file.

---

## Asset sizing rules (critical — never stretch photos)

All photos must render at their **natural aspect ratio**. Pre-calculated render widths at `h = 7.5"`:

| File       | Pixel dimensions | render_w |
|------------|-----------------|----------|
| image5.png | 1485 × 1124     | 9.91"    |
| image6.png | 1033 × 1125     | 6.89"    |
| image7.png | 1033 × 1125     | 6.89"    |
| image8.png | 1165 × 1093     | 7.99"    |
| image9.png |  925 × 1009     | 6.88"    |
| image10.png| 1106 × 1124     | 7.38"    |

Formula: `render_w = (pixel_w / pixel_h) × 7.5`  
Placement: always right-anchored → `x = 13.33 - render_w`

---

## Brand colours

| Token      | Hex       | Usage                                      |
|------------|-----------|--------------------------------------------|
| purple     | `#875EFA` | Titles, accent bars, action cards          |
| lime       | `#C4FF00` | Thank-you slide bottom stripe              |
| green      | `#1DCD86` | Positive deltas, on-track indicators       |
| blue       | `#2A62F5` | Secondary chart series                     |
| black      | `#1A1A1A` | Body text, slide numbers                   |
| offwhite   | `#F4F4F4` | Chart panel backgrounds                    |
| insightbg  | `#EFEFEF` | Insight bar background                     |

Fonts: **Aptos Display** (titles) · **Aptos** (body)

---

## Slide order convention

```
1. Title slide
2. Divider
3–N. Content slides (charts, stats, recommendations)
N+1. Divider
N+2. Thank You
```

---

## Troubleshooting

| Problem | Fix |
|---------|-----|
| Photos look stretched | Check render_w values — never use `sizing: { type: 'contain' }` |
| Logo missing on a slide | `logoMark(s)` on every slide except first and last |
| Logo appearing on title/thank-you | Remove `logoMark(s)` call — first and last slides use full wordmark only |
| Slide title overflows | Use `fontSize: 16` on recommendation slides; keep titles under 120 chars |
| Assets not found | Verify GitHub raw URLs are correct and repo is public (or token is valid) |
