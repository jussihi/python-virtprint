# virtprint - Virtual Printer for Windows

A Python-based Windows service that creates a TCP network printer. Print jobs are sent over TCP/IP to a local listener service which converts them to PDF or image files and saves them to a user-defined location on the file system.

## Features

- Creates a virtual Windows printer with Standard TCP/IP port (RAW protocol)
- Runs as a Windows service
- Listens on configurable TCP port (default: 9100 - standard RAW/JetDirect port)
- Automatically captures and processes print jobs sent to the virtual printer
- Converts print jobs to PDF or image formats (PNG, JPEG, TIFF)
- Supports multiple input formats: PostScript, PDF, XPS
- Configurable output folder, format and image quality settings
- Custom callbacks for post-processing (email, cloud upload, OCR, etc.)
- Comprehensive logging of all operations

## Requirements

- Windows 10 or later (requires Microsoft PostScript Class Driver)
- Python 3.8 or later
- Administrator privileges (for service installation and printer creation)
- GhostScript (automatically downloaded to local `gxps` folder during `install.py`)

## Installation

1. **Configure settings** - Edit `settings.py` to customize:
   - TCP port (default: 9100)
   - Output directory (default: C:\temp)
   - Output format (PDF, PNG, JPEG, TIFF)
   - Service name and display name
   - Printer name

2. **Run installation** (requires Administrator privileges):
   ```cmd
   python install.py
   ```
   
   This will:
   - Download and extract GhostScript to the `gxps` folder
   - Create a Python virtual environment (`.venv`)
   - Install required Python packages
   - Install virtprint as a service if you want to

### Printer management

The printer is automatically installed when the program/service starts. You can also manually manage it:

```cmd
# Install printer only
python virtprint.py --install

# Uninstall printer
python virtprint.py --uninstall
```

### Service management

If you chose to install the service, you can manage it from Windows' `services.msc`. You can also
manually manage the service:

```cmd
# Install the service (run as Administrator)
python service.py install
```

```cmd
# Start the service
python service.py start
```

```cmd
# Stop the service
python service.py stop
```

```cmd
# Remove the service
python service.py remove
```

## How It Works

1. A Windows printer is created with a Standard TCP/IP port pointing to `127.0.0.1:9100` (configurable)
2. The virtprint service listens on the TCP port for incoming print jobs
3. When a document is printed to the virtual printer, Windows sends the print data over TCP
4. The service receives the data, detects the format (PostScript, PDF, or XPS)
5. GhostScript converts the data to the configured output format (PDF or images)
6. Files are saved to the configured output directory
7. Optional callbacks are triggered for post-processing


## Custom Callbacks !

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


## Troubleshooting

### Service won't start
- Check `virtprint.log` for error messages
- Ensure port 9100 (or your configured port) is not in use by another application
- Verify Administrator privileges

### Printer shows offline or error state
- Verify the service is running: `python service.py start`
- Check that the TCP port matches in both `settings.py` and the printer configuration
- Disable bidirectional support in printer properties if enabled

### Print jobs don't create files
- Check `virtprint.log` for conversion errors
- Verify GhostScript is installed in the `gxps` folder
- Ensure the output directory exists and is writable
- Check that the PostScript driver is installed correctly

## Log Files

The program creates detailed logs in `virtprint.log` for troubleshooting and monitoring purposes. Logs include:
- Service start/stop events
- TCP connection information
- Print job details (size, format, processing time)
- Conversion status and errors
- Callback execution results

## Security Notes

- This program requires Administrator privileges to install system printers and services
- The TCP listener binds to localhost (127.0.0.1) by default
- Print job data is temporarily stored in the system's temp directory during processing
- Output files contain the original print job data - handle sensitive documents appropriately
- If being run as a service, runs under the Local System account by default
