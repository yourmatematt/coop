#!/bin/bash
# Convert HTML slides to PNG using Chrome/Chromium headless mode
# Usage: ./convert-html-to-png.sh

# Check if Chrome/Chromium is available
CHROME=""
if command -v google-chrome &> /dev/null; then
    CHROME="google-chrome"
elif command -v chromium &> /dev/null; then
    CHROME="chromium"
elif command -v chromium-browser &> /dev/null; then
    CHROME="chromium-browser"
elif command -v chrome &> /dev/null; then
    CHROME="chrome"
else
    echo "Error: Chrome/Chromium not found. Please install Google Chrome or Chromium."
    exit 1
fi

echo "Using: $CHROME"
echo "Converting HTML slides to PNG..."
echo ""

# Create output directory if it doesn't exist
mkdir -p png-output

# Convert each slide
for i in {1..17}; do
    HTML_FILE="slide${i}.html"
    OUTPUT_FILE="png-output/slide-$(printf '%02d' $i).png"

    if [ -f "$HTML_FILE" ]; then
        echo "Converting $HTML_FILE -> $OUTPUT_FILE"
        $CHROME --headless --screenshot="$OUTPUT_FILE" --window-size=1920,1080 --default-background-color=0 "$HTML_FILE" 2>/dev/null
    else
        echo "Warning: $HTML_FILE not found"
    fi
done

echo ""
echo "✓ Conversion complete!"
echo "PNG files saved to: png-output/"
ls -lh png-output/
