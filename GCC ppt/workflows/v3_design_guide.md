# V3 Design Guide: McKinsey-Grade Presentation Improvements

## Purpose
This document details every design change needed to upgrade the India GCC presentation from V2 to McKinsey-grade quality. Use this as a specification when building the presentation in any tool (Canva, PowerPoint, Google Slides, python-pptx).

---

## 1. Design System — Color Palette

### Primary Colors
| Token | Hex | Usage |
|-------|-----|-------|
| Black | `#1A1A1A` | Action titles only |
| Charcoal | `#333333` | Sub-headers |
| Body Gray | `#4A4A4A` | All body text — softer than pure black, prevents competition with titles |
| Mid Gray | `#666666` | Captions, data callout labels, secondary text |
| Rule Gray | `#DDDDDD` | Thin horizontal divider lines between slide sections |
| Border Gray | `#CCCCCC` | Box borders |
| Light Gray | `#F5F5F5` | Box backgrounds, callout containers (slightly warmer than pure #F2F2F2) |
| White | `#FFFFFF` | Slide background, table alternate rows |

### Accent Colors
| Token | Hex | Usage |
|-------|-----|-------|
| Accent Orange | `#E84D0E` | Section labels, data callout numbers, box titles, chevron flows |
| Banner Gold | `#F5C518` | "So What" bottom banner |
| Green | `#27AE60` | Positive metrics, GCC advantage boxes, ROI numbers |
| Light Green BG | `#E8F8EE` | Background fill for green-themed bordered boxes |
| Red | `#C0392B` | Warning metrics, risk callouts, challenge boxes |
| Light Red BG | `#FDEDED` | Background fill for red-themed bordered boxes |
| Yellow BG | `#FEF9E7` | Heatmap mid-range cells (scores 6-7) |
| Blue Accent | `#2C3E50` | Methodology/appendix box titles, case study callouts |
| Divider Navy | `#1A1A2E` | Dark background for section divider slides |

---

## 2. Typography — 5-Level Hierarchy

This is the single most impactful change. V2 used a flat 2-level system (title + body). V3 introduces 5 distinct levels that create scannable visual hierarchy.

### Level 0: Section Label
- **Size:** 9pt
- **Weight:** Bold
- **Color:** Accent Orange (`#E84D0E`)
- **Case:** UPPERCASE
- **Position:** Top of slide, y=0.15", above the action title
- **Content:** "ACT I: THE OPPORTUNITY", "LOCATION DEEP DIVE", "RISK ANALYSIS", etc.
- **Purpose:** Orients the reader within the narrative structure

### Level 1: Action Title
- **Size:** 26pt (V2 was 22pt — too small)
- **Weight:** Bold
- **Color:** Black (`#1A1A1A`)
- **Position:** y=0.35", left margin
- **Width:** Full content width (11.9")
- **Purpose:** THE slide message. If someone only reads titles, they get the full story.
- **Rule:** Every title must be a complete sentence/assertion, not a topic label
- **Examples:**
  - GOOD: "Bangalore: The Established Leader Facing Saturation Risk"
  - BAD: "Bangalore Overview"

### Level 2: Sub-header
- **Size:** 14pt
- **Weight:** Bold
- **Color:** Charcoal (`#333333`)
- **Purpose:** Section labels within a slide when header bars feel too heavy
- **Usage:** Optional — use when slide has 2-3 distinct sections that need labeling without the visual weight of a dark header bar

### Level 3: Body Text
- **Size:** 11pt (10pt in tight layouts)
- **Weight:** Regular
- **Color:** Body Gray (`#4A4A4A`) — NOT black
- **Line spacing:** 1.3x-1.45x font size
- **Purpose:** Primary readable content
- **Critical change from V2:** Body text must be gray, not black. Black body text competes with the 26pt title and flattens the hierarchy. Gray body text lets the eye naturally flow: title → metrics → body.

### Level 4: Captions & Small Text
- **Size:** 8-9pt
- **Weight:** Regular
- **Color:** Mid Gray (`#666666`)
- **Purpose:** Data callout labels, table cells, box body text, chevron descriptions

### Level 5: Source/Footer
- **Size:** 7pt
- **Weight:** Regular (slide number: bold)
- **Color:** Mid Gray (`#666666`)
- **Position:** Bottom of slide, y=7.15"
- **Content:** "Sources: Zinnov, NASSCOM, ANSR..." on left, slide number on right

---

## 3. Layout & Spacing — White Space Rules

### The 30% Rule
McKinsey's #1 design principle: **30% of every slide should be white space.** This means:

### Margins
| Edge | V2 Value | V3 Value | Change |
|------|----------|----------|--------|
| Left margin | 0.6" | 0.7" | +0.1" |
| Right margin | 0.6" | 0.7" | +0.1" |
| Content width | 12.1" | 11.9" | -0.2" |
| Content start (Y) | 1.2"-1.3" | 1.5" | +0.2"-0.3" more breathing room below title |

### Spacing Between Elements
| Element | V2 | V3 |
|---------|----|----|
| Section label to title | 15px | 20px (0.2") |
| Title to first content | 0.9"-1.0" gap | 1.15" gap |
| Bullet item spacing (space_after) | 3pt | 6pt |
| Bullet line spacing | 1.4x | 1.45x |
| Between content sections | No separator | Horizontal rule + 0.15" padding above/below |
| Between grid boxes | 0.1" | 0.15" |

### Horizontal Rules
- **New in V3:** Thin 1pt gray lines (`#DDDDDD`) placed between major sections on a slide
- **Purpose:** Creates visual compartments without heavy borders or header bars
- **Placement:** Full content width, positioned between metrics row and body content, between body and bottom callouts
- **Example positions on a city slide:**
  - After metrics row: y=2.85"
  - Before bottom callouts: y=5.6" (if applicable)

---

## 4. Component Design Changes

### 4.1 Data Callouts — Contained in Boxes (P0/P1)

**V2 Problem:** Big orange numbers floating above gray labels with no visual boundary. Looks unfinished.

**V3 Solution:** Every data callout sits inside a light bordered container.

**Spec:**
- Background: Light Gray (`#F5F5F5`)
- Border: 0.75pt, Rule Gray (`#DDDDDD`)
- Corner radius: 0 (sharp corners — consulting style)
- Padding: 0.08" top, 0.1" sides
- Internal layout:
  - Number: centered, bold, 26-32pt, Accent Orange (or Green/Red for context)
  - Label: centered, 8-9pt, Mid Gray, 2-3 words max
- Box dimensions: ~2.4-2.8" wide x 1.15" tall
- Spacing between boxes: 0.12-0.15" gap

**Example — metrics row:**
```
┌──────────┐  ┌──────────┐  ┌──────────┐  ┌──────────┐
│  1,800+  │  │   $65B   │  │   230+   │  │  $40.4B  │
│ GCCs in  │  │  Annual  │  │  BFSI    │  │  BFSI    │
│  India   │  │ Revenue  │  │  GCCs    │  │  Market  │
└──────────┘  └──────────┘  └──────────┘  └──────────┘
```

### 4.2 Heatmap Scorecard Table (P1)

**V2 Problem:** The scorecard table on Slide 3 is a plain table with alternating gray/white rows. All scores look the same — the reader has to manually compare numbers.

**V3 Solution:** Conditional formatting on score cells creates an instant visual heatmap.

**Color rules for score cells (columns 3-6 of scorecard):**

| Score | Background | Text Color | Text Weight |
|-------|-----------|------------|-------------|
| 8-10 | Light Green (`#E8F8EE`) | Dark Green (`#1B7A43`) | Bold |
| 6-7 | Light Yellow (`#FEF9E7`) | Dark Yellow (`#7D6C08`) | Regular |
| 1-5 | Light Red (`#FDEDED`) | Red (`#C0392B`) | Regular |

**Total row:** Gray background (`#E0E0E0`), bold black text.

**Result:** The reader instantly sees Bangalore dominating in green, Pune's cost advantage in green, Delhi NCR's connectivity in green — without reading a single number.

### 4.3 City Slides — 2x2 Quad Grid (P1)

**V2 Problem:** City slides (4-7) had 7-8 bullet points in a long list with a side box. Too much text, no visual structure. Violates McKinsey's "one quick scan" principle.

**V3 Solution:** Replace bullet lists with a **2x2 grid of bordered boxes**, each containing 4 short items.

**Grid spec:**
- Total area: full content width x ~3.35" height
- 2 columns, 2 rows
- Gap between boxes: 0.15"
- Each box: ~5.9" wide x ~1.6" tall

**Standard 4-box layout for every city slide:**

| Position | Box Title | Title Color | Content |
|----------|-----------|-------------|---------|
| Top-left | STRENGTHS | Orange (`#E84D0E`) | 4 key advantages |
| Top-right | CHALLENGES | Red (`#C0392B`) | 4 key risks/limitations |
| Bottom-left | KEY BFSI GCCs | Green (`#27AE60`) | 4 major companies |
| Bottom-right | POLICY INCENTIVES | Blue (`#2C3E50`) | 4 government incentives |

**Box title format:** Unicode symbol + space + title (e.g., "■  STRENGTHS", "▲  CHALLENGES")

**Each box contains:** 4 bullet items, 9pt, Body Gray, with bullet character "•"

**Why this works:** The reader can scan one box in 3 seconds. Total slide scan time drops from 30s (reading bullets) to 12s (scanning 4 boxes).

### 4.4 "So What" Bottom Banner — Slimmer (P2)

**V2 Problem:** Banner is 0.55" tall with 11pt bold black text. It dominates the bottom of the slide and visually shouts.

**V3 Solution:**

| Property | V2 | V3 |
|----------|----|----|
| Height | 0.55" | 0.40" |
| Y position | 6.55" | 6.65" |
| Font size | 11pt | 10pt |
| Font color | Black | Dark Gray (`#2D2D2D`) |
| Text left margin | 0.6" | 0.7" |
| Background | Gold (`#F5C518`) | Gold (`#F5C518`) — unchanged |

**The banner should feel like an annotation, not a call-to-action.** It's the "so what" whisper at the bottom that rewards the reader who made it to the end of the slide.

### 4.5 Header Bars (Minor Refinement)

**V2:** 0.38" tall, 12pt bold white text on dark gray.

**V3:**
- Height: 0.35" (slightly thinner)
- Font size: 11pt (slightly smaller)
- All other properties unchanged
- **Usage rule:** Maximum ONE header bar per slide (two header bars create visual confusion). If a slide needs two sections, use one header bar + one horizontal rule + sub-header text.

### 4.6 Bordered Boxes (Minor Refinement)

**V2:** 1pt border, `#CCCCCC`.

**V3:**
- Border: 0.75pt (thinner — more refined)
- Padding: 0.12" left/right (increased from 0.1")
- Body text size: 9pt default (explicitly set, not inherited)
- Body text color: Body Gray (`#4A4A4A`)
- Line spacing in body: 1.35x

### 4.7 Chevron Process Flows (Minor Refinement)

**V2:** Chevrons touching each other, 0.65" tall.

**V3:**
- Height: 0.60" (slightly shorter)
- Gap between chevrons: 0.08" (adds breathing room)
- Description text below: 8pt, Mid Gray, line spacing 1.25x
- Gradient still dark-to-light orange from left to right

---

## 5. New Component: Section Divider Slides

**V2 Problem:** The deck jumps from Act I directly into Act II content with no visual pause. For a 20+ slide deck, this makes the narrative structure invisible.

**V3 Solution:** Add 3 dark-background divider slides between Acts.

### Divider Slide Spec

- **Background:** Full-bleed Divider Navy (`#1A1A2E`)
- **Orange accent bar:** 1.2" wide x 0.06" tall at y=2.8", left-aligned at x=0.7"
- **Act number:** 14pt bold, Accent Orange, y=3.05" (e.g., "ACT II")
- **Act title:** 34pt bold, White, y=3.5" (e.g., "Location Analysis")
- **Subtitle:** 14pt regular, Light Gray (`#999999`), y=4.4" (e.g., "Four cities scored across 8 dimensions with BFSI-specific weights")

### Placement

| Divider | Position | Act Number | Title | Subtitle |
|---------|----------|-----------|-------|----------|
| 1 | After Slide 2, before Slide 3 | ACT II | Location Analysis | Four cities scored across 8 dimensions with BFSI-specific weights |
| 2 | After Slide 9, before Slide 10 | ACT III | Operating Model & Talent | GCC advantage, risk landscape, talent strategy, and EVP framework |
| 3 | After Slide 17, before Slide 18 | ACT IV | From Analysis to Action | Implementation roadmap, specific recommendations, and source methodology |

**Total slide count:** 23 (20 content + 3 dividers)

---

## 6. Slide-by-Slide Specifications

### Slide 1: Title
- Dark navy background (`#1A1A2E`)
- Top orange accent bar: full width, 0.06" tall
- Small left accent bar: 1.2" wide x 0.05" at y=2.0"
- "India GCC Landscape" — 42pt bold white, y=2.3"
- "Strategic Analysis for Financial Services" — 26pt regular, Accent Orange, y=3.3"
- "Location Strategy | Operating Model | Talent Playbook" — 13pt, light gray (`#AAAAAA`), y=4.4"
- "Board-Ready Strategy Document | 2025-2026 Decision Window" — 11pt, darker gray (`#777777`), y=5.6"
- Bottom orange accent bar: full width, 0.06" at y=7.44"

### Slide 2: Market Inflection Point
- Section label: "ACT I: THE OPPORTUNITY"
- Title: "The India GCC Market Has Reached an Inflection Point"
- 6 contained metric callouts in a row (1,800+ | $65B | 230+ | $40.4B | 2,400+ | $100B)
- Horizontal rule at y=2.85"
- Header bar: "BFSI: THE LARGEST GCC VERTICAL"
- 5 bullet points (11pt, Body Gray)
- Warning box (red bg): "DECISION WINDOW: 2025-2026" with ⚠ icon
- Gold banner: "BFSI is the largest GCC vertical growing 3x the market..."

### Slides 3: Scorecard
- Heatmap table (see Section 4.2 above)
- Horizontal rule after table
- 4 contained score summary callouts (8.1 Bangalore | 7.2 Hyderabad | 6.8 Delhi NCR | 6.5 Pune)

### Slides 4-7: City Deep Dives
Each follows identical structure:
1. Section label: "LOCATION DEEP DIVE"
2. Action title (unique per city — assertion, not topic)
3. 4 contained metric callouts
4. Horizontal rule
5. 2x2 quad grid (Strengths | Challenges | Key GCCs | Policy/Model)
6. Gold "So What" banner
7. Source footer

### Slide 8: Hub-and-Spoke
- Header bar: "FUNCTION → CITY MAPPING MATRIX"
- 6-row table (Function | City | Rationale)
- Horizontal rule
- 2 contained callouts (70%+ multi-city | 30-40% Tier-2 savings)
- Bordered recommendation box with → icon

### Slide 9: TCO Model
- 9-row comparison table (Outsourcing | GCC | BOT Hybrid)
- Horizontal rule
- 3 contained callouts (M24-30 | 15-20% | 4-6 mo)
- Red warning box: "COST OF INACTION"

### Slide 10: GCC vs Third-Party
- 9-row comparison table with Winner column
- Horizontal rule
- 2 contained callouts (6 of 8 in green | 2 of 8 in gray)
- Gray box: "WHEN OUTSOURCING WINS"

### Slide 11: Innovation Engine
- 4 contained callouts (3.2x | 55% | 63% | 80%)
- Horizontal rule
- Header bar: "INNOVATION IN ACTION"
- 3 case study boxes across (Goldman | JP Morgan | Deutsche)
- Green financial impact bar

### Slide 12: Third-Party Risks
- 4 RED contained callouts (30% | $6.08M | $4.6B | $0.5-1B)
- Horizontal rule
- Header bar: regulatory landscape
- Regulation table (5 rows)
- Red warning box: real-world breaches

### Slide 13: Risk Register
- 11-row table (# | Risk | Severity | Mitigation)
- Full slide height table — no additional elements needed
- This is the most text-dense slide — acceptable for a risk register

### Slide 14: Hybrid Model
- Chevron flow (4 phases)
- Horizontal rule
- Header bar: "WHAT STAYS CAPTIVE vs WHAT TO OUTSOURCE"
- 2 side-by-side boxes: Green "GCC-OWNED CORE" | Orange "SELECTIVE OUTSOURCING"
- Horizontal rule
- 2 contained callouts (<10%→40% BOT surge | 2x mega GCC outsourcing)

### Slide 15: Talent Crisis
- 4 RED contained callouts (60% | 49% | 75% | 10.4%)
- Horizontal rule + header bar
- 5 bullet points (left side, 7.3" wide)
- Red box: "ANNUAL COST OF ATTRITION" (right side)
- Orange box: "HARDEST TO HIRE" (right side, below red)

### Slide 16: EVP Framework
- 5 pillar boxes with numbered orange circles above each
- Horizontal rule
- Header bar: "GARTNER HUMAN DEAL FRAMEWORK"
- 6-row table (Component | Description | GCC Application)

### Slide 17: Branding Playbook
- 3-phase chevron flow
- Horizontal rule + header bar: "ROI OF EMPLOYER BRANDING"
- 5 GREEN contained callouts (43% | 28% | 50% | 40% | 3x)
- Horizontal rule + header bar: "KEY BRANDING TACTICS"
- 4 bullet points (9pt)

### Slide 18: Implementation Roadmap
- 4 phase header bars (orange gradient) with content boxes below
- Horizontal rule
- Header bar: "KEY MILESTONES"
- 6-row milestones table

### Slide 19: Recommendations
- 5 numbered recommendations with orange icon circles
- Each recommendation: 13pt bold title + 10pt gray detail line
- Horizontal rule
- Red bordered "COST OF INACTION" warning strip

### Slide 20: Sources
- Header bar: "44 SOURCES ACROSS 7 CATEGORIES"
- 7 bordered boxes in a grid (4 columns, 2 rows)
- Blue-tinted methodology box at bottom

---

## 7. Quick Reference — V2 vs V3 Changes Summary

| Element | V2 | V3 | Impact |
|---------|----|----|--------|
| Action title size | 22pt | 26pt | High — commands attention |
| Body text color | Black (#1A1A1A) | Body Gray (#4A4A4A) | High — creates hierarchy |
| Content start Y | 1.2"-1.3" | 1.5" | High — breathing room |
| Bullet spacing | 3pt after | 6pt after | Medium — readability |
| Data callouts | Floating numbers | Contained in bordered boxes | High — polished look |
| Scorecard table | Plain alternating rows | Heatmap (green/yellow/red) | High — instant insight |
| City slides | Long bullet lists | 2x2 quad grid | High — scannable |
| Section dividers | None | 3 dark navy slides between Acts | Medium — narrative pacing |
| Bottom banner height | 0.55" | 0.40" | Medium — less visual weight |
| Bottom banner text | 11pt bold black | 10pt bold dark gray | Medium — quieter |
| Horizontal rules | None | Thin gray lines between sections | Medium — structure |
| Header bar height | 0.38" | 0.35" | Low — subtle refinement |
| Box borders | 1pt | 0.75pt | Low — more refined |
| Total slides | 20 | 23 (20 + 3 dividers) | Medium — better flow |

---

## 8. Design Principles to Follow

1. **One message per slide.** The action title IS the message. Everything else is evidence.
2. **30% white space.** If a slide feels crowded, remove content — don't shrink fonts.
3. **Body text is never black.** Only the action title gets pure black. Everything else is gray.
4. **Maximum 4-5 bullets per section.** If you need more, restructure into boxes or a table.
5. **Every visual element earns its pixel.** If a shape doesn't communicate data, remove it.
6. **Contain your data.** Numbers in boxes > numbers floating in space.
7. **Horizontal rules > heavy borders.** Subtle separation > visual walls.
8. **The "So What" banner whispers.** It's the reward for reading the slide, not the headline.
9. **Consistent structure breeds trust.** City slides should feel identical in layout. The reader learns the pattern once and then scans faster.
10. **Section dividers create narrative pacing.** A 2-second pause between acts lets the audience reset.
