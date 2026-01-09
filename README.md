# VirtPrint - Virtual Printer for Windows

A Python program that creates a virtual printer on Windows that accepts print jobs and converts them to PDF or image files, saving them to `C:\temp`.

## Features

- Creates a virtual printer 
- Automatically captures print jobs sent to the virtual printer
- Converts print jobs to PDF or image formats (PNG, JPEG, TIFF)
- Configurable output format and image quality settings
- Saves output files to `C:\temp` with timestamps
- Provides fallback text file creation if conversion fails
- Comprehensive logging of all operations

## Requirements

- Windows 10 or later (requires Microsoft PostScript driver)
- Python 3.6 or later
- Administrator privileges (for printer installation)
- GhostScript (downloaded during installation)

## Installation

1. Modify `settings.py` to your liking

2. If using default settings, Install the Windows PostScript Class Driver

3. Install GhostScript, service and its virtual env:
   ```cmd
   python install.py
   ```

4. Start the service from Windows services.msc:


## Configuration

Edit `settings.py` to customize the virtual printer behavior:

### Output Format
```python
OUTPUT_FORMAT = "PDF"  # Options: "PDF", "PNG", "JPEG", "TIFF"
```

### Image Quality (for image output)
```python
IMAGE_DPI = 300  # Resolution: 72-1200 (300 recommended for print quality)
IMAGE_COLOR_DEPTH = "24bit"  # Options: "24bit", "8bit", "1bit"
```

### Other Settings
- `PRINTER_NAME` - Name of the virtual printer
- `OUTPUT_DIR` - Directory where files are saved
- `DRIVER_NAME` - Windows printer driver to use
- `PORT_NAME` - Printer port configuration

## Custom Callbacks

virtprint supports custom callbacks that are triggered when a print job is completed. Edit `callbacks.py` to define what happens after each print job.

### Basic Usage

The `on_print_job_complete()` function in `callbacks.py` is called with:
- `output_files`: List of created files (in page order) or `None` if failed
- `job_info`: Dictionary with job details (document name, user, pages, etc.)

### Examples

**Copy files to network share:**
```python
def on_print_job_complete(output_files, job_info):
    if output_files:
        for file_path in output_files:
            shutil.copy2(file_path, r"\\server\share\documents")
```

**Email attachments:**
```python
def on_print_job_complete(output_files, job_info):
    if output_files:
        send_email(
            to="user@example.com",
            subject=f"Print: {job_info['document_name']}",
            attachments=output_files
        )
```

**Organize by date:**
```python
def on_print_job_complete(output_files, job_info):
    if output_files:
        today = datetime.now()
        dest = Path(f"C:\\Printed\\{today.year}\\{today.month:02d}")
        dest.mkdir(parents=True, exist_ok=True)
        for file_path in output_files:
            shutil.move(str(file_path), str(dest / file_path.name))
```

See `callbacks.py` for more examples including OCR, cloud uploads, and database logging.

## Log Files

The program creates detailed logs in `virtprint.log` for troubleshooting and monitoring purposes.

## Security Notes

- This program requires Administrator privileges to install system printers
- Print job data is temporarily stored in the system's temp directory
- Output files contain the original print job data - handle sensitive documents appropriately
