"""
Virtual Printer Configuration Settings
All configurable settings for the virtual printer application.

Output Format Options:
- Set OUTPUT_FORMAT to control the output file type:
  * "PDF" - Standard PDF format (default)
  * "PNG" - PNG image format
  * "JPEG" - JPEG image format  
  * "TIFF" - TIFF image format

Image Quality Settings:
- IMAGE_DPI: Resolution for image output (72-1200, typically 300 for print quality)
- IMAGE_COLOR_DEPTH: Color depth for PNG images
  * "24bit" - Full color (default)
  * "8bit" - Grayscale
  * "1bit" - Monochrome/black & white

Note: Image conversion requires GhostScript to be installed and in PATH.
For best results with image output, install GhostScript from: https://ghostscript.com/
"""

# Printer Configuration
PRINTER_NAME = "Virtual File Printer"
DRIVER_NAME = "Microsoft PS Class Driver"  # PostScript driver - widely available
PORT_NAME = "LPT1:"  # Default port for virtual printer

# Output Settings
OUTPUT_DIR = "C:\\temp"
OUTPUT_FORMAT = "PNG"  # Options: "PDF", "PNG", "JPEG", "TIFF"
IMAGE_DPI = 300  # DPI for image output (higher = better quality, larger file)
IMAGE_COLOR_DEPTH = "24bit"  # Options: "24bit" (color), "8bit" (grayscale), "1bit" (monochrome)

# Printer Info Dictionary
PRINTER_INFO = {
    'pServerName': None,  # None means local machine
    'pShareName': '',  # Empty string means not shared
    'pPrintProcessor': 'winprint',
    'pDatatype': 'RAW',
    'pSepFile': '',  # Separator file (empty means none)
    'pLocation': 'Virtual PDF Printer',
    'pComment': 'Virtual printer that converts print jobs to PDF',
    'pParameters': '',  # Printer parameters (empty means default)
    'pSecurityDescriptor': None,  # Security descriptor (None means default)
    'Attributes': None,  # Will be set to win32print.PRINTER_ATTRIBUTE_LOCAL at runtime
    'Priority': 1,  # Printer priority (1 = normal)
    'DefaultPriority': 1,  # Default priority for print jobs (1 = normal)
    'StartTime': 0,  # Start time for printer availability (0 = always available)
    'UntilTime': 0,  # End time for printer availability (0 = always available)
    'Status': 0,  # Printer status (0 = ready/idle)
    'cJobs': 0,  # Number of jobs currently in queue (0 = empty)
    'AveragePPM': 60  # Average pages per minute (60 = fast virtual printer)
}

# Alternative ports to try if the default port is not available
ALTERNATIVE_PORTS = ["LPT1:", "COM1:", "FILE:", "USB001", "USB002"]

# Logging Configuration
LOG_FILE = 'virtprint.log'
LOG_LEVEL = 'INFO'
LOG_FORMAT = '%(asctime)s - %(levelname)s - %(message)s'
