# Quick Conversion Guide

## What's in This Folder?

This folder contains everything you need to convert the Mallacoota Beauty Pivot pitch deck slides to PNG format.

### Files Included:

- **17 HTML slide files** (`slide1.html` through `slide17.html`) - Full HD 1920x1080px
- **index.html** - Preview all slides in your browser
- **manifest.json** - Slide metadata
- **README.md** - Comprehensive conversion instructions
- **convert-html-to-png.sh** - Automated conversion script (Mac/Linux)
- **convert-html-to-png.bat** - Automated conversion script (Windows)

## Fastest Methods to Get PNG Files

### Option A: From PowerPoint File (Recommended)

1. Open `../mallacoota-beauty-pivot.pptx` in PowerPoint or LibreOffice
2. File → Export → PNG/Images
3. Select 1920x1080 resolution
4. Export all slides
5. Done! ✓

### Option B: From HTML Files (Automated)

**On Mac/Linux:**
```bash
./convert-html-to-png.sh
```

**On Windows:**
```cmd
convert-html-to-png.bat
```

This will create PNG files in `png-output/` folder using Chrome headless mode.

### Option C: Manual Screenshot (Individual Slides)

1. Open `index.html` in your browser
2. Click on any slide to open full size
3. Take screenshot or use browser screenshot tool
4. Crop to 1920x1080 if needed

## Expected Output

After conversion, you'll have 17 PNG files:

```
slide-01.png  →  Title: "From Crisis to Opportunity"
slide-02.png  →  The Crisis (75% revenue collapse)
slide-03.png  →  Why Traditional Strategies Won't Work
slide-04.png  →  The Strategic Insight
slide-05.png  →  Market Opportunity ($82B)
slide-06.png  →  Stage-Gate Strategy Overview
slide-07.png  →  Gate 1: Lab Testing
slide-08.png  →  Gates 2-3: Legal & Testing
slide-09.png  →  Gate 4: Limited Launch
slide-10.png  →  Gate 5: Scale & Sustain
slide-11.png  →  Messaging: "Pure. Wild. Proven."
slide-12.png  →  Financial Projections & ROI
slide-13.png  →  Risk Management
slide-14.png  →  Implementation Timeline
slide-15.png  →  Why This Strategy Works
slide-16.png  →  Next Steps (5-step action plan)
slide-17.png  →  Call to Action
```

## Quality Settings

- **Resolution**: 1920 x 1080 pixels (Full HD)
- **Format**: PNG (supports transparency)
- **DPI**: 96+ recommended
- **Color**: RGB, sRGB color space

## Troubleshooting

**Chrome script not working?**
- Install Google Chrome if not already installed
- Update the path in the script if Chrome is in a custom location

**PowerPoint export blurry?**
- Check export settings for resolution
- Ensure you're selecting "High Quality (1920x1080)" option

**Need different sizes?**
- Edit the `--window-size=1920,1080` parameter in scripts
- Or resize in image editor after export

## Questions?

See the full `README.md` in this folder for detailed instructions and alternative methods.
