# Mallacoota Beauty Pivot - Slide Exports

## PNG Export Instructions

Unfortunately, automatic PPTX to PNG conversion is not available in this environment. Here are several methods to export the slides as PNG images:

### Method 1: Using PowerPoint (Windows/Mac)

1. Open `mallacoota-beauty-pivot.pptx` in Microsoft PowerPoint
2. Go to **File** → **Export** → **Export as Pictures**
3. Choose **PNG** format
4. Set resolution to **1920 x 1080** pixels
5. Click **Export All Slides**
6. Save to this folder (`slides-png/`)

### Method 2: Using LibreOffice (Linux/Windows/Mac)

1. Open `mallacoota-beauty-pivot.pptx` in LibreOffice Impress
2. Go to **File** → **Export**
3. Select **PNG - Portable Network Graphic** as file type
4. In the PNG Options dialog:
   - Width: **1920** px
   - Height: **1080** px
   - Resolution: **96** DPI (or higher for print quality)
5. Click **Export**
6. Choose "Export every slide" option

### Method 3: Using Google Slides

1. Upload `mallacoota-beauty-pivot.pptx` to Google Drive
2. Open with Google Slides
3. Go to **File** → **Download** → **PNG image (current slide)**
4. Repeat for each slide, or use an add-on for batch export

### Method 4: Using Command Line (Linux/Mac with LibreOffice installed)

```bash
# Convert PPTX to PDF first
libreoffice --headless --convert-to pdf mallacoota-beauty-pivot.pptx

# Then convert PDF to PNG (requires pdftoppm)
pdftoppm -png -r 150 -scale-to-x 1920 -scale-to-y 1080 mallacoota-beauty-pivot.pdf slide
```

### Method 5: Online Conversion Tools

- **Zamzar**: https://www.zamzar.com/convert/pptx-to-png/
- **CloudConvert**: https://cloudconvert.com/pptx-to-png
- **Online-Convert**: https://www.online-convert.com/

## Slide Reference

All 17 slides are Full HD resolution (1920x1080px) with professional ocean/luxury beauty design:

### Slide Descriptions:

| # | Title | Description |
|---|-------|-------------|
| 1 | Title Slide | "From Crisis to Opportunity: The Beauty Pivot" |
| 2 | The Crisis | 75% revenue collapse, price competition visualization |
| 3 | Why Traditional Won't Work | Chef programs, retail marketing, branding challenges |
| 4 | The Strategic Insight | Category shift, cultural beliefs, lab opportunity |
| 5 | Market Opportunity | $82B market, growth trends, consumer segments |
| 6 | Stage-Gate Overview | 5-gate flowchart with investment amounts |
| 7 | Gate 1: Lab Testing | Detailed testing requirements and success criteria |
| 8 | Gates 2-3: Legal & Testing | Compliance and micro-influencer validation |
| 9 | Gate 4: Limited Launch | "Mallacoota Beauty Reserve" premium line |
| 10 | Gate 5: Scale & Sustain | Multi-market expansion strategy |
| 11 | Messaging Framework | "Pure. Wild. Proven." brand positioning |
| 12 | Financial Projections | Revenue forecasts and ROI calculations |
| 13 | Risk Management | Gate system and controlled exposure |
| 14 | Implementation Timeline | 12-month roadmap with milestones |
| 15 | Why This Works | 6 strategic rationale points |
| 16 | Next Steps | 5-step action plan |
| 17 | Call to Action | Decision framework and closing |

## Expected PNG Files

After conversion, you should have 17 PNG files:

```
slide-01.png  (Title Slide)
slide-02.png  (The Crisis)
slide-03.png  (Why Traditional Won't Work)
slide-04.png  (The Strategic Insight)
slide-05.png  (Market Opportunity)
slide-06.png  (Stage-Gate Overview)
slide-07.png  (Gate 1: Lab Testing)
slide-08.png  (Gates 2-3: Legal & Testing)
slide-09.png  (Gate 4: Limited Launch)
slide-10.png  (Gate 5: Scale & Sustain)
slide-11.png  (Messaging Framework)
slide-12.png  (Financial Projections)
slide-13.png  (Risk Management)
slide-14.png  (Implementation Timeline)
slide-15.png  (Why This Works)
slide-16.png  (Next Steps)
slide-17.png  (Call to Action)
```

## Design Specifications

- **Resolution**: 1920 x 1080 pixels (Full HD)
- **Aspect Ratio**: 16:9
- **Color Palette**:
  - Deep Ocean Blue: #003366
  - Pearl White: #F8F8F8
  - Gold Accent: #D4AF37
  - Seafoam Green: #7FCDCD
- **Typography**: Segoe UI (Bold for headlines, Regular for body)

## Files in This Directory

- `index.html` - Preview all 17 slides in your browser
- `manifest.json` - Slide metadata and information
- `README.md` - This file

## Preview Slides in Browser

Open `index.html` in your web browser to preview all slides before exporting to PNG.

## Need Help?

If you have issues with conversion:

1. Check that your PowerPoint application is up to date
2. Ensure you're exporting at full resolution (1920x1080)
3. Verify PNG files maintain transparency if needed
4. Contact support if automated batch export is required
