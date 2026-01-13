"""
Virtual Printer Configuration Settings
All configurable settings for the TCP virtual printer application.

Output Format Options:
- Set OUTPUT_FORMAT to control the output file type:
  * "PDF" - Standard PDF format (default)
  * "PNG" - PNG image format
  * "JPEG" - JPEG image format  
  * "TIFF" - TIFF image format
  * "PS" - PostScript
  * "RAW" - Raw print data

Image Quality Settings:
- IMAGE_DPI: Resolution for image output (72-1200, typically 300 for print quality)
- IMAGE_COLOR_DEPTH: Color depth for PNG images
  * "24bit" - Full color (default)
  * "8bit" - Grayscale
  * "1bit" - Monochrome/black & white

Note: Image conversion requires GhostScript. It can be installed by running install.py
or by manually downloading it from https://www.ghostscript.com/download/gsdnld.html
"""

# TCP Configuration
TCP_HOST = "127.0.0.1"  # Loopback address for local printer
TCP_PORT = 9100  # Standard RAW/JetDirect port

# Windows Service Configuration
SERVICE_NAME = "virtprint"
SERVICE_DISPLAY_NAME = "virtprint Virtual Printer Service"
SERVICE_DESCRIPTION = "Virtual printer service that converts print jobs to PDF/image files"

# Printer Configuration
WINDOWS_PRINTER_PORT_NAME = f"{SERVICE_NAME}_{TCP_PORT}"
PRINTER_NAME = "Virtual File Printer"
DRIVER_NAME = "Microsoft PS Class Driver"  # PostScript driver for Windows printer integration

# Output Settings
OUTPUT_DIR = "C:\\temp"
OUTPUT_FORMAT = "PDF"  # Options: "PDF", "PNG", "JPEG", "TIFF", "PS", "RAW"
IMAGE_DPI = 300  # DPI for image output (higher = better quality, larger file)
IMAGE_COLOR_DEPTH = "24bit"  # Options: "24bit" (color), "8bit" (grayscale), "1bit" (monochrome)


# Logging Configuration
LOG_FILE = 'virtprint.log'
LOG_LEVEL = 'INFO'
LOG_FORMAT = '%(asctime)s - %(levelname)s - %(message)s'
