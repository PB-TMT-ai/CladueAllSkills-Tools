# Price Dashboard Restyle Project

## Purpose
Restyle `D:\JSW_TMT_Price_Dashboard.pptx` to match the Haryana PPT design system.

## Source Files
- **Input PPT**: `D:\JSW_TMT_Price_Dashboard.pptx`
- **Design Reference**: `D:\RandomTestsClaude\Haryana_NewSlides.pptx`
- **Excel Data**: `D:\RandomTestsClaude\North TMT pricing_9th Mar.xlsx` (sheet: `1_Summary at PL`)
- **Chart Logic**: `D:\RandomTestsClaude\generate_price_slides.py` (reusable chart + XML combo helpers)

## Haryana Design System
- **Font**: Calibri (exclusively)
- **Background**: White
- **Colors**:
  - Blue (headers/accents): #18489D
  - Dark Navy (footer bar): #0D2B5E
  - Amber: #D97706
  - Green: #05723A
  - Body text: #1E293B
  - Gray (titles): #7F7F7F
  - Table alt rows: #F0F4F8
  - White text on dark: #FFFFFF

## Color Mapping (Current → Target)
- #E8272A (red accent) → #18489D (blue)
- #1A2B3C (card bg) → #F0F4F8 (light gray)
- #0D1B2A (alt row bg) → #FFFFFF (white)
- #223447 (section header) → #18489D (blue)
- #2A2A18 (TISCON row) → #F0F4F8
- #2A1518 (JSW row) → #F0F4F8
- #F5A623 (amber) → #D97706
- #27AE60 (green) → #05723A

## Font Mapping
- Arial → Calibri
- Trebuchet MS → Calibri
- Consolas → Calibri

## Rules
- Keep slide dimensions (9,144,000 x 5,143,500 EMU)
- Keep all content and layout positions
- Replace embedded chart IMAGES with native PPT charts using Haryana colors
- Do NOT change data values or slide structure

## Market → Excel Column Mapping
| Market | Distr Col | Dealer Col | Gap Col |
|--------|-----------|------------|---------|
| DL | 2 | 3 | 4 |
| HR | 5 | 6 | 7 |
| RJ | 8 | 9 | 10 |
| UP | 11 | 12 | 13 |
| CH | 14 | 15 | 16 |
| PB | 17 | 18 | 19 |
| UK | 20 | 21 | 22 |
| HP | 23 | 24 | 25 |
| JK | 26 | 27 | 28 |
