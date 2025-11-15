#!/usr/bin/env python3
"""
Convert PowerPoint slides to PNG images
"""

import os
import sys
import subprocess

# Try using LibreOffice's uno interface
def convert_with_uno():
    """Convert PPTX to PNG using LibreOffice UNO"""
    try:
        import uno
        from com.sun.star.beans import PropertyValue

        print("Using LibreOffice UNO interface...")
        # This would require a running LibreOffice instance
        return False
    except ImportError:
        return False

# Try using libreoffice command line with proper syntax
def convert_with_libreoffice_cli():
    """Convert PPTX to individual PNG files using LibreOffice"""

    pptx_file = os.path.abspath("mallacoota-beauty-pivot.pptx")
    output_dir = os.path.abspath("slides-png")

    if not os.path.exists(pptx_file):
        print(f"Error: {pptx_file} not found")
        return False

    print(f"Converting {pptx_file}...")
    print(f"Output directory: {output_dir}")

    # First, convert PPTX to PDF
    pdf_file = os.path.join(output_dir, "presentation.pdf")

    cmd = [
        "libreoffice",
        "--headless",
        "--invisible",
        "--convert-to", "pdf",
        "--outdir", output_dir,
        pptx_file
    ]

    print(f"Running: {' '.join(cmd)}")

    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
        print("STDOUT:", result.stdout)
        print("STDERR:", result.stderr)
        print("Return code:", result.returncode)

        if result.returncode == 0:
            print("PDF created successfully")

            # Now convert PDF to PNG using pdftoppm if available
            if os.path.exists(pdf_file):
                return convert_pdf_to_png(pdf_file, output_dir)

        return False
    except Exception as e:
        print(f"Error: {e}")
        return False

def convert_pdf_to_png(pdf_file, output_dir):
    """Convert PDF to PNG images"""
    # Check if pdftoppm is available
    try:
        subprocess.run(["pdftoppm", "-v"], capture_output=True, check=True)

        # Convert PDF to PNG at high resolution (1920x1080)
        cmd = [
            "pdftoppm",
            "-png",
            "-r", "150",  # DPI
            "-scale-to-x", "1920",
            "-scale-to-y", "1080",
            pdf_file,
            os.path.join(output_dir, "slide")
        ]

        print(f"Converting PDF to PNG: {' '.join(cmd)}")
        result = subprocess.run(cmd, capture_output=True, text=True)

        if result.returncode == 0:
            print("Successfully converted PDF to PNG images")
            return True
        else:
            print(f"pdftoppm error: {result.stderr}")
            return False

    except (subprocess.CalledProcessError, FileNotFoundError):
        print("pdftoppm not available")
        return False

# Try using ImageMagick
def convert_with_imagemagick():
    """Convert PPTX using ImageMagick"""
    try:
        pptx_file = os.path.abspath("mallacoota-beauty-pivot.pptx")
        output_pattern = os.path.join(os.path.abspath("slides-png"), "slide-%02d.png")

        cmd = [
            "convert",
            "-density", "150",
            "-resize", "1920x1080",
            pptx_file,
            output_pattern
        ]

        print(f"Trying ImageMagick: {' '.join(cmd)}")
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)

        if result.returncode == 0:
            print("Successfully converted with ImageMagick")
            return True
        else:
            print(f"ImageMagick error: {result.stderr}")
            return False

    except (subprocess.CalledProcessError, FileNotFoundError) as e:
        print(f"ImageMagick not available: {e}")
        return False

if __name__ == "__main__":
    os.chdir("/home/user/coop/presentation")

    print("=" * 60)
    print("PowerPoint to PNG Converter")
    print("=" * 60)

    # Try different methods
    methods = [
        ("LibreOffice CLI", convert_with_libreoffice_cli),
        ("ImageMagick", convert_with_imagemagick),
    ]

    success = False
    for method_name, method_func in methods:
        print(f"\nTrying method: {method_name}")
        print("-" * 40)
        if method_func():
            success = True
            print(f"✓ Success with {method_name}")
            break
        else:
            print(f"✗ {method_name} failed")

    if not success:
        print("\n" + "=" * 60)
        print("Unable to convert slides automatically.")
        print("=" * 60)
        print("\nAlternative options:")
        print("1. Open mallacoota-beauty-pivot.pptx in PowerPoint/LibreOffice")
        print("2. Use File > Export > Export as PNG/Images")
        print("3. Use online conversion tools")
        print("4. Use 'pdftoppm' after converting to PDF manually")
        sys.exit(1)
    else:
        print("\n" + "=" * 60)
        print("✓ Conversion completed successfully!")
        print("=" * 60)
        print(f"\nPNG files saved to: {os.path.abspath('slides-png')}")

        # List created files
        png_files = sorted([f for f in os.listdir('slides-png') if f.endswith('.png')])
        if png_files:
            print(f"\nCreated {len(png_files)} PNG files:")
            for f in png_files:
                print(f"  - {f}")
