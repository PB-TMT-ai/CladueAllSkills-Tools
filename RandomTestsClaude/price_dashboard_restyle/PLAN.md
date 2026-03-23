# Restyle JSW_TMT_Price_Dashboard.pptx to Match Haryana Design

## Context
The JSW_TMT_Price_Dashboard.pptx has 11 slides with a dark theme (dark navy backgrounds, Arial/Consolas fonts, red JSW accents). It needs to be restyled to match the Haryana_NewSlides.pptx design system (white background, Calibri font, blue/navy accents) while keeping all content, layout, and data identical.

## Current vs Target Design

| Element | Current (JSW Dashboard) | Target (Haryana Design) |
|---------|------------------------|------------------------|
| Background | Dark (#0D1B2A, #1A2B3C) | White |
| Font | Arial, Trebuchet MS, Consolas | Calibri (exclusively) |
| Brand accent | Red #E8272A | Blue #18489D |
| Section headers | #223447 bg | #18489D (blue rounded rect) |
| Card backgrounds | #1A2B3C | #F0F4F8 (light gray) |
| Body text | White #FFFFFF (on dark) | Dark #1E293B (on white) |
| Labels/secondary | Muted blue #8BA3C0 | Gray #7F7F7F |
| Amber/benchmark | #F5A623 | #D97706 |
| Green (competitive) | #27AE60 | #05723A |
| Footer bar | N/A (just text) | Dark navy #0D2B5E bar |
| Title style | White bold | Gray #7F7F7F bold |

## Color Mapping

### Shape fill colors:
```
#E8272A (red accent bars, status badges) → #18489D (blue)
#1A2B3C (card/row backgrounds)          → #F0F4F8 (light gray)
#0D1B2A (alternate row backgrounds)     → #FFFFFF (white)
#223447 (section header bg)             → #18489D (blue)
#2A2A18 (TISCON benchmark row)          → #F0F4F8 (light banding)
#2A1518 (JSW brand row highlight)       → #F0F4F8 (light banding)
#1A6FB5 (summary slide accent)          → #18489D (blue)
#F5A623 (amber status fill)             → #D97706 (Haryana amber)
#27AE60 (green status fill)             → #05723A (Haryana green)
```

### Text color mapping (context-aware):
- White text on shapes that become light/white → #1E293B (dark text)
- White text on shapes that stay dark (blue headers, status badges) → stays #FFFFFF
- #E8272A (red brand text) → #18489D (blue)
- #8BA3C0 (muted labels) → #7F7F7F (gray)
- #F5A623 (amber text) → #D97706
- #27AE60 (green text) → #05723A

### Font mapping:
- Arial → Calibri
- Trebuchet MS → Calibri
- Consolas → Calibri

## Implementation

### Single Python script: `restyle_dashboard.py`
1. Open `D:\JSW_TMT_Price_Dashboard.pptx`
2. Iterate ALL shapes on ALL 11 slides
3. For each shape:
   - Map fill color using the fill mapping
   - Determine if the shape's new fill is "dark" (blue/navy/status) or "light" (white/F0F4F8)
   - For each text run: map font name to Calibri, map text color based on context
4. Replace chart images with native PPT charts
5. Save to `D:\JSW_TMT_Price_Dashboard_Restyled.pptx`

### Chart regeneration
- Replace embedded chart IMAGES with native PPT grouped bar charts
- Read market data from Excel (`North TMT pricing_9th Mar.xlsx`, 1st tab)
- Use Haryana color scheme: Distributor bars=#18489D (blue), Dealer bars=#D97706 (amber)
- Add TISCON reference as dashed line
- Reuse chart creation logic from `generate_price_slides.py`

### Slide dimensions
- Keep current size (9,144,000 x 5,143,500 EMU / 10x5.63in)

## Verification
1. Run `python restyle_dashboard.py`
2. Open output and verify:
   - White backgrounds with light gray cards
   - All text in Calibri
   - Blue accents instead of red
   - Correct gap badge colors (green/amber/red mapped to Haryana palette)
   - Readable text (no white-on-white issues)
   - Native charts with correct data and colors
