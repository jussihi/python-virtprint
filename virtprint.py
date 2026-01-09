#!/usr/bin/env python3
"""
Virtual Printer for Windows
Creates a virtual printer that accepts print jobs and converts them to PDF files.
Saves output to C:\temp directory.
"""

import os
import sys
import time
import subprocess
import threading
from datetime import datetime
import win32print
import win32api
import win32con
import win32gui
import win32file
import pywintypes
import tempfile
import shutil
from pathlib import Path
import logging
import settings

# Import user-defined callback
try:
    from callbacks import on_print_job_complete
    logger_init = logging.getLogger(__name__)
    logger_init.info("Loaded user callback from callbacks.py")
except ImportError:
    # Fallback if callbacks.py doesn't exist
    def on_print_job_complete(output_files, job_info):
        pass
    logger_init = logging.getLogger(__name__)
    logger_init.warning("callbacks.py not found, using default (no-op) callback")

# Configure logging
logging.basicConfig(
    level=getattr(logging, settings.LOG_LEVEL),
    format=settings.LOG_FORMAT,
    handlers=[
        logging.FileHandler(settings.LOG_FILE),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class VirtualPrinter:
    def __init__(self, printer_name=None, output_dir=None, output_format=None):
        self.printer_name = printer_name or settings.PRINTER_NAME
        self.output_dir = Path(output_dir or settings.OUTPUT_DIR)
        self.output_dir.mkdir(exist_ok=True)
        self.output_format = (output_format or settings.OUTPUT_FORMAT).upper()
        self.running = False
        self.monitor_thread = None

        # Ensure we have required directories
        self.temp_dir = Path(tempfile.gettempdir()) / "virtprint"
        self.temp_dir.mkdir(exist_ok=True)
        
        # Find GhostScript executable
        self.gs_executable = self.find_ghostscript_executable()
        if self.gs_executable:
            logger.info(f"Found GhostScript at: {self.gs_executable}")
        else:
            logger.warning("GhostScript executable not found - conversion may fail")

        # Validate output format
        supported_formats = ['PDF', 'PNG', 'JPEG', 'TIFF']
        if self.output_format not in supported_formats:
            logger.warning(f"Unsupported format {self.output_format}, defaulting to PDF")
            self.output_format = 'PDF'

        logger.info(f"Output format set to: {self.output_format}")
    
    def find_ghostscript_executable(self):
        """Find the GhostScript executable in the gxps folder or system PATH."""
        script_dir = Path(__file__).parent.resolve()
        
        # Check in gxps folder first (installed by install.py)
        gxps_dir = script_dir / "gxps"
        possible_executables = [
            gxps_dir / "gxpswin64.exe",
            gxps_dir / "gxpswin32.exe",
            gxps_dir / "gswin64c.exe",
            gxps_dir / "gswin32c.exe",
        ]
        
        for exe in possible_executables:
            if exe.exists():
                return str(exe)
        
        # Check if in PATH
        for exe_name in ["gxpswin64.exe", "gswin64c.exe", "gswin32c.exe", "gs.exe"]:
            try:
                result = subprocess.run(["where", exe_name], capture_output=True, text=True)
                if result.returncode == 0:
                    path = result.stdout.strip().split('\n')[0]
                    if path:
                        return path
            except:
                pass
        
        return None

    def create_local_port(self, port_name):
        """Create a local port for the virtual printer if it doesn't exist."""
        try:
            # Check if port already exists
            ports = win32print.EnumPorts(None, 1)
            existing_ports = [port[1] for port in ports]
            
            if port_name in existing_ports:
                logger.info(f"Port {port_name} already exists")
                return True
            
            # Add a new local port
            # For virtual printers, we can use a local port that redirects to NUL
            # or create a custom port name
            logger.info(f"Creating local port: {port_name}")
            
            # Use XcvData to add a local port
            # This requires the Local Port Monitor
            try:
                # Open the local port monitor
                xcv_handle = win32print.OpenPrinter(",XcvMonitor Local Port,", {"DesiredAccess": win32print.SERVER_ACCESS_ADMINISTER})
                
                # Add the port
                port_name_bytes = (port_name + '\0').encode('utf-16le')
                
                # Use XcvData to add the port
                result = win32print.XcvData(
                    xcv_handle,
                    "AddPort",
                    port_name_bytes,
                    0
                )
                
                win32print.ClosePrinter(xcv_handle)
                logger.info(f"Successfully created local port: {port_name}")
                return True
                
            except Exception as e:
                logger.error(f"Failed to create port via XcvData: {e}")
                # Fall back to using an existing port
                return False
                
        except Exception as e:
            logger.error(f"Error creating local port: {e}")
            return False
    
    def verify_port_available(self, port_name):
        """Verify that the specified port is available in the system."""
        try:
            # Check if the port exists in the system
            ports = win32print.EnumPorts(None, 1)
            existing_ports = [port[1] for port in ports]

            logger.info(f"Available ports in system: {existing_ports}")

            if port_name in existing_ports:
                logger.info(f"Port {port_name} is available")
                return True
            else:
                logger.error(f"Port {port_name} not found in system")

                # Try to find a suitable alternative port
                for alt_port in settings.ALTERNATIVE_PORTS:
                    if alt_port in existing_ports:
                        logger.info(f"Found alternative port: {alt_port}")
                        return True

                # If no suitable port found, we'll try to use the first available port
                if existing_ports:
                    logger.warning(f"Using first available port: {existing_ports[0]}")
                    return True

                return False

        except Exception as e:
            logger.error(f"Error checking port availability: {e}")
            # If we can't enumerate ports, assume LPT1: exists (it usually does)
            logger.warning("Cannot verify ports, assuming LPT1: is available")
            return True

    def create_devmode(self, driver_name):
        """Create a DEVMODE structure for the printer."""
        try:
            # Try to get default DEVMODE from an existing printer using the same driver
            printers = win32print.EnumPrinters(2)
            drivers = win32print.EnumPrinterDrivers(None, None, 2)
            print(printers)
            print(drivers)
            for printer in printers:
                try:
                    if len(printer) >= 3:
                        printer_info = printer[1]
                        printer_name = printer[2]
                        print(printer_name, driver_name)
                        if driver_name in printer_info:
                            hprinter = win32print.OpenPrinter(printer_name)
                            printer_data = win32print.GetPrinter(hprinter, 2)
                            devmode = printer_data.get('pDevMode')
                            win32print.ClosePrinter(hprinter)
                            if devmode:
                                logger.info("Using DEVMODE from existing printer")
                                return devmode
                except Exception as e:
                    logger.warning(f"Could not get DEVMODE from printer {printer}: {e}")
                    continue

            # If we can't get DEVMODE from existing printers, just return None
            # This will let Windows use the driver's default settings
            logger.info("No existing printer with matching driver found, using Windows default DEVMODE")
            return None

        except Exception as e:
            logger.warning(f"Could not create DEVMODE: {e}")
            return None

    def install_printer(self):
        """Install the virtual printer driver and create printer instance."""
        try:
            logger.info("Installing virtual printer...")

            # Check if printer already exists and verify its configuration
            printers = win32print.EnumPrinters(2)
            printer_exists = False
            for printer in printers:
                if len(printer) >= 3 and printer[2] == self.printer_name:
                    printer_exists = True
                    # Check if it's configured correctly
                    try:
                        handle = win32print.OpenPrinter(self.printer_name)
                        printer_info = win32print.GetPrinter(handle, 2)
                        current_port = printer_info.get('pPortName', '')
                        win32print.ClosePrinter(handle)
                        
                        if current_port == settings.PORT_NAME:
                            logger.info(f"Printer '{self.printer_name}' already exists with correct port ({current_port}).")
                            return True
                        else:
                            logger.warning(f"Printer exists but has wrong port ({current_port}), expected {settings.PORT_NAME}. Recreating...")
                            self.uninstall_printer()
                            printer_exists = False
                    except Exception as e:
                        logger.warning(f"Could not verify printer configuration: {e}")
                    break
            
            if printer_exists:
                logger.info(f"Printer '{self.printer_name}' already exists.")
                return True

            # Use PostScript driver for true virtual printing
            # Use PostScript driver with LPT1 port (we'll intercept before it reaches the port)
            driver_name = settings.DRIVER_NAME
            port_name = settings.PORT_NAME
            
            # Fall back to checking if the configured port exists
            if not self.verify_port_available(port_name):
                logger.error(f"Port {port_name} is not available and could not create custom port")
                # Try alternative ports
                port_found = False
                for alt_port in settings.ALTERNATIVE_PORTS:
                    if alt_port != "FILE:" and self.verify_port_available(alt_port):
                        port_name = alt_port
                        port_found = True
                        logger.info(f"Using alternative port: {alt_port}")
                        break
                
                if not port_found:
                    logger.error("No suitable port found")
                    return False

            # Check if the driver exists
            drivers = win32print.EnumPrinterDrivers(None, None, 2)
            driver_exists = any(driver['Name'] == driver_name for driver in drivers)

            if not driver_exists:
                logger.error(f"Driver '{driver_name}' not found. Please install PostScript support or check available drivers.")
                # List available drivers for debugging
                available_drivers = [driver['Name'] for driver in drivers[:5]]  # Show first 5
                logger.info(f"Available drivers include: {available_drivers}")
                return False

            # Verify the port is available
            if not self.verify_port_available(port_name):
                logger.error(f"Port {port_name} is not available")
                return False

            logger.info(f"Using port: {port_name}")

            # Create printer info structure
            # Get DEVMODE from the create_devmode method
            devmode = self.create_devmode(driver_name)

            # Build printer info from settings
            printer_info = settings.PRINTER_INFO.copy()
            printer_info['pPrinterName'] = self.printer_name
            printer_info['pPortName'] = port_name
            printer_info['pDriverName'] = driver_name
            printer_info['pDevMode'] = devmode
            printer_info['Attributes'] = win32print.PRINTER_ATTRIBUTE_LOCAL

            # Add the printer
            win32print.AddPrinter(None, 2, printer_info)
            logger.info(f"Successfully installed printer '{self.printer_name}'")
            return True

        except Exception as e:
            logger.error(f"Failed to install printer: {e}")
            return False

    def uninstall_printer(self):
        """Remove the virtual printer."""
        try:
            # Open and delete the printer
            handle = win32print.OpenPrinter(self.printer_name)
            win32print.DeletePrinter(handle)
            win32print.ClosePrinter(handle)
            logger.info(f"Successfully uninstalled printer '{self.printer_name}'")
            return True
        except Exception as e:
            logger.error(f"Failed to uninstall printer: {e}")
            return False

    def monitor_print_jobs(self):
        """Monitor print spooler and intercept jobs before they reach the port."""
        logger.info("Starting print job interception...")

        processed_jobs = set()

        while self.running:
            try:
                handle = win32print.OpenPrinter(self.printer_name)

                # Get all jobs for our printer
                jobs = win32print.EnumJobs(handle, 0, -1, 1)

                for job in jobs:
                    job_id = job.get('JobId')
                    if job_id and job_id not in processed_jobs:
                        logger.info(f"New print job detected: {job.get('pDocument', 'Unknown')} (ID: {job_id})")

                        # Process the job - don't pause to avoid error state
                        if self.intercept_and_process_job(handle, job):
                            processed_jobs.add(job_id)
                        else:
                            # If processing failed, just delete the job
                            try:
                                win32print.SetJob(handle, job_id, 0, None, win32print.JOB_CONTROL_DELETE)
                                logger.info(f"Deleted failed job {job_id}")
                                processed_jobs.add(job_id)
                            except Exception as e:
                                logger.error(f"Could not delete failed job {job_id}: {e}")

                win32print.ClosePrinter(handle)
                time.sleep(0.5)  # Check frequently to catch jobs quickly

            except Exception as e:
                logger.error(f"Error monitoring print jobs: {e}")
                time.sleep(2)

    def intercept_and_process_job(self, printer_handle, job):
        """Intercept a print job and process it before it reaches the port."""
        output_files = None
        job_info = {
            'job_id': job.get('JobId'),
            'document_name': job.get('pDocument', 'Unknown'),
            'user_name': job.get('pUserName', 'Unknown'),
            'machine_name': job.get('pMachineName', 'Unknown'),
            'pages': job.get('TotalPages', 0),
            'output_format': self.output_format
        }

        try:
            job_id = job['JobId']
            document_name = job.get('pDocument', f"print_job_{job_id}")

            logger.info(f"Intercepting job {job_id}: {document_name}")

            # Don't pause the job - let it spool naturally to avoid "error" state
            # Wait for the job to be fully spooled first
            job_status = job.get('Status', 0)
            logger.info(f"Job {job_id} status: {job_status}")

            # Try to read the job data directly from the spooler
            spool_data = self.read_job_from_spooler(job_id)

            if spool_data:
                # Create output file from the intercepted data
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                safe_doc_name = "".join(c for c in document_name if c.isalnum() or c in (' ', '-', '_')).strip()
                file_ext = self.output_format.lower()
                output_filename = f"{timestamp}_{safe_doc_name}.{file_ext}"
                output_path = self.output_dir / output_filename

                output_files = self.convert_intercepted_data_to_output(spool_data, output_path, document_name)

                if output_files:
                    logger.info(f"Successfully intercepted and converted: {len(output_files)} file(s)")
                else:
                    # Create info file as fallback
                    self.create_job_info_file(job, output_path.with_suffix('.txt'))
                    output_files = None

                # Mark the job as complete and then delete it
                # Don't pause/restart - just delete directly to avoid errors
                try:
                    win32print.SetJob(printer_handle, job_id, 0, None, win32print.JOB_CONTROL_DELETE)
                    logger.info(f"Job {job_id} deleted successfully")
                except Exception as e:
                    logger.warning(f"Could not delete job {job_id}: {e}")

                # Call user-defined callback
                try:
                    on_print_job_complete(output_files, job_info)
                except Exception as e:
                    logger.error(f"Error in user callback: {e}")

                return True
            else:
                logger.warning(f"Could not read job data for job {job_id}")
                # Just delete the job directly
                try:
                    win32print.SetJob(printer_handle, job_id, 0, None, win32print.JOB_CONTROL_DELETE)
                except Exception as e:
                    logger.warning(f"Could not delete job {job_id}: {e}")

                # Call callback with None to indicate failure
                try:
                    on_print_job_complete(None, job_info)
                except Exception as e:
                    logger.error(f"Error in user callback: {e}")

                return False

        except Exception as e:
            logger.error(f"Error intercepting job: {e}")
            # Call callback with None to indicate failure
            try:
                on_print_job_complete(None, job_info)
            except Exception as e:
                logger.error(f"Error in user callback: {e}")
            return False

    def read_job_from_spooler(self, job_id):
        """Read raw print data from the Windows spooler with proper timing."""
        try:
            # Look for spool files in the Windows spooler directory
            spool_dir = Path(os.environ.get('SystemRoot', 'C:\\Windows')) / "System32" / "spool" / "PRINTERS"

            # Wait for the job to be fully spooled before reading
            max_attempts = 30  # Increased for large documents
            for attempt in range(max_attempts):
                try:
                    # Check job status first
                    handle = win32print.OpenPrinter(self.printer_name)
                    job_info = win32print.GetJob(handle, job_id, 1)
                    job_status = job_info.get('Status', 0)
                    job_pages = job_info.get('TotalPages', 0)
                    job_size = job_info.get('Size', 0)
                    win32print.ClosePrinter(handle)

                    logger.debug(f"Attempt {attempt + 1}: Job {job_id} - Status: {job_status}, Pages: {job_pages}, Size: {job_size}")

                    # Look for spool files
                    spool_files = list(spool_dir.glob("*.SPL"))

                    if spool_files:
                        # Get the most recently modified file (likely our job)
                        latest_spool = max(spool_files, key=lambda f: f.stat().st_mtime)

                        # Check if file is still being written to
                        file_size = latest_spool.stat().st_size
                        if file_size == 0:
                            logger.debug(f"Spool file {latest_spool} is empty, waiting...")
                            time.sleep(0.5)
                            continue

                        # Wait for file size to be stable (multiple checks)
                        # This is critical for large multi-page documents
                        stable_size = self.wait_for_spool_file_stable(latest_spool, min_checks=5, max_wait=30)

                        if not stable_size:
                            logger.debug(f"Spool file {latest_spool} is still changing, continuing to wait...")
                            continue

                        # File size is stable, try to read it
                        try:
                            logger.info(f"Reading spool file {latest_spool} ({stable_size} bytes)")
                            with open(latest_spool, 'rb') as f:
                                data = f.read()

                            if len(data) > 0:
                                logger.info(f"Successfully read {len(data)} bytes from spool file: {latest_spool}")
                                return data
                            else:
                                logger.debug(f"Spool file {latest_spool} is empty, retrying...")

                        except (PermissionError, FileNotFoundError) as e:
                            logger.debug(f"Could not read spool file (attempt {attempt + 1}): {e}")

                    else:
                        logger.debug(f"No spool files found yet (attempt {attempt + 1})")

                    # Wait before next attempt
                    time.sleep(0.5)

                except Exception as e:
                    logger.debug(f"Error in attempt {attempt + 1}: {e}")
                    time.sleep(0.5)

            # If we get here, all attempts failed
            logger.warning(f"Failed to read spool data for job {job_id} after {max_attempts} attempts")

            # As a last resort, try to find any spool file with data
            try:
                spool_files = list(spool_dir.glob("*.SPL"))
                for spool_file in sorted(spool_files, key=lambda f: f.stat().st_mtime, reverse=True):
                    try:
                        if spool_file.stat().st_size > 0:
                            with open(spool_file, 'rb') as f:
                                data = f.read()
                            if len(data) > 100:  # At least some reasonable amount of data
                                logger.info(f"Found fallback spool data: {len(data)} bytes from {spool_file}")
                                return data
                    except:
                        continue
            except:
                pass

            return None

        except Exception as e:
            logger.error(f"Error reading from spooler: {e}")
            return None

    def wait_for_spool_file_stable(self, spool_file, min_checks=5, max_wait=30, check_interval=0.5):
        """Wait for spool file to stop growing (fully written).

        Args:
            spool_file: Path to the spool file
            min_checks: Minimum number of consecutive stable size checks required
            max_wait: Maximum time to wait in seconds
            check_interval: Time between size checks in seconds

        Returns:
            File size if stable, None if still changing
        """
        try:
            start_time = time.time()
            last_size = -1
            stable_count = 0

            while time.time() - start_time < max_wait:
                try:
                    current_size = spool_file.stat().st_size

                    if current_size == last_size and current_size > 0:
                        stable_count += 1
                        logger.debug(f"Spool file stable check {stable_count}/{min_checks}: {current_size} bytes")

                        if stable_count >= min_checks:
                            logger.info(f"Spool file is stable at {current_size} bytes after {stable_count} checks")
                            return current_size
                    else:
                        if last_size >= 0:  # Not first check
                            logger.debug(f"Spool file size changed: {last_size} -> {current_size} bytes (resetting stability counter)")
                        stable_count = 0

                    last_size = current_size
                    time.sleep(check_interval)

                except FileNotFoundError:
                    logger.debug(f"Spool file disappeared during stability check")
                    return None

            # Timeout - file might still be stable
            if last_size > 0 and stable_count >= 2:
                logger.warning(f"Timeout waiting for spool file stability, but seems stable at {last_size} bytes")
                return last_size

            logger.warning(f"Spool file not stable after {max_wait} seconds")
            return None

        except Exception as e:
            logger.error(f"Error waiting for spool file stability: {e}")
            return None

    def convert_intercepted_data_to_output(self, data, output_path, document_name):
        """Convert intercepted print data to the configured output format (PDF or image).

        Returns:
            List of Path objects for created files, or None if conversion failed
        """
        try:
            # First convert to PDF (if not already PDF)
            if self.output_format == 'PDF':
                pdf_file = self.convert_intercepted_data_to_pdf(data, output_path, document_name)
                return [pdf_file] if pdf_file else None
            else:
                # For image formats, first convert to PDF, then to image
                temp_pdf = self.temp_dir / f"temp_{int(time.time())}_{os.getpid()}.pdf"

                # Convert to PDF first
                pdf_file = self.convert_intercepted_data_to_pdf(data, temp_pdf, document_name)
                if pdf_file:
                    # Ensure PDF is fully written to disk before conversion
                    if temp_pdf.exists():
                        # Wait for file to be stable (fully written)
                        self.wait_for_file_ready(temp_pdf)

                        logger.info(f"Temp PDF created: {temp_pdf.stat().st_size} bytes")

                        # Then convert PDF to image - returns list of image files
                        image_files = self.convert_pdf_to_image(temp_pdf, output_path)

                        # Keep temp PDF for debugging if conversion fails
                        if image_files:
                            temp_pdf.unlink()  # Clean up temp PDF
                        else:
                            logger.warning(f"Keeping temp PDF for debugging: {temp_pdf}")

                        return image_files
                    else:
                        logger.error(f"Temp PDF was not created: {temp_pdf}")
                        return None
                else:
                    logger.warning(f"Could not convert to PDF, trying direct conversion to {self.output_format}")
                    # Try direct conversion for some formats
                    image_file = self.convert_data_to_image(data, output_path, document_name)
                    return [image_file] if image_file else None

        except Exception as e:
            logger.error(f"Error converting to {self.output_format}: {e}")
            return None

    def wait_for_file_ready(self, file_path, max_wait=5):
        """Wait for a file to be fully written and stable."""
        try:
            start_time = time.time()
            last_size = -1
            stable_count = 0

            while time.time() - start_time < max_wait:
                if not file_path.exists():
                    time.sleep(0.1)
                    continue

                current_size = file_path.stat().st_size

                if current_size == last_size and current_size > 0:
                    stable_count += 1
                    if stable_count >= 3:  # File size stable for 3 checks
                        logger.debug(f"File ready: {file_path} ({current_size} bytes)")
                        return True
                else:
                    stable_count = 0

                last_size = current_size
                time.sleep(0.1)

            # If we get here, we timed out but file might still be usable
            if file_path.exists() and file_path.stat().st_size > 0:
                logger.warning(f"File may not be fully stable, but proceeding: {file_path}")
                return True

            return False

        except Exception as e:
            logger.error(f"Error waiting for file: {e}")
            return False

    def convert_intercepted_data_to_pdf(self, data, output_path, document_name):
        """Convert intercepted print data to PDF.

        Returns:
            Path object of created PDF file, or None if conversion failed
        """
        try:
            logger.info(f"Analyzing {len(data)} bytes of print data")

            # Check the first 500 bytes to determine format
            header = data[:500]
            logger.debug(f"Data header (first 50 bytes): {header[:50]}")

            if data.startswith(b'%PDF'):
                # Data is already PDF
                with open(output_path, 'wb') as f:
                    f.write(data)
                    f.flush()
                    os.fsync(f.fileno())
                logger.info(f"Saved PDF data directly: {output_path}")
                return Path(output_path)

            elif data.startswith(b'%!PS') or b'%!PS-Adobe' in header:
                # Data is PostScript
                logger.info("Detected PostScript data")
                return Path(output_path) if self.convert_ps_data_to_pdf(data, output_path) else None

            elif b'%!PS' in data[:1000] or b'PostScript' in header or b'%%Creator' in header:
                # PostScript data with some prefix
                logger.info("Detected PostScript data with prefix")
                return Path(output_path) if self.convert_ps_data_to_pdf(data, output_path) else None

            elif data.startswith(b'PK') or b'FixedDocument' in data or b'Documents/' in data:
                # XPS (XML Paper Specification) format - it's a ZIP file
                logger.info("Detected XPS data format")
                return Path(output_path) if self.convert_xps_to_pdf(data, output_path, document_name) else None

            elif b'<html>' in header.lower() or b'<!doctype' in header.lower():
                # HTML data
                logger.info("Detected HTML data")
                return Path(output_path) if self.convert_html_to_pdf(data, output_path, document_name) else None

            elif header.startswith(b'\x1b') or b'PCL' in header:
                # PCL (Printer Command Language) data
                logger.info("Detected PCL data")
                return Path(output_path) if self.convert_pcl_to_pdf(data, output_path, document_name) else None

            elif data.startswith(b'\x89PNG') or data.startswith(b'\xff\xd8\xff'):
                # Image data (PNG, JPEG)
                logger.info("Detected image data")
                return Path(output_path) if self.convert_image_to_pdf(data, output_path, document_name) else None

            else:
                # Unknown binary format - create info PDF instead of trying to extract text
                logger.info(f"Unknown binary format, creating info PDF for: {document_name}")
                return Path(output_path) if self.create_simple_pdf(f"Binary Print Job: {document_name}\n\nData format not recognized.\nFile size: {len(data)} bytes", output_path) else None

        except Exception as e:
            logger.error(f"Error converting intercepted data: {e}")
            return None

    def convert_xps_to_pdf(self, data, output_path, document_name):
        """Convert XPS (XML Paper Specification) data to PDF using GhostScript."""
        try:
            logger.info("Converting XPS to PDF using GhostScript")

            # Save the XPS data as a temporary file
            temp_xps = self.temp_dir / f"temp_{int(time.time())}.xps"
            with open(temp_xps, 'wb') as f:
                f.write(data)

            # Try to convert XPS to PDF using GhostScript
            if self.convert_xps_with_ghostscript(temp_xps, output_path):
                logger.info(f"Successfully converted XPS to PDF: {output_path}")
                temp_xps.unlink()  # Clean up
                return True

            # If GhostScript conversion fails, try alternative methods
            logger.warning("GhostScript XPS conversion failed, trying alternative approach")

            # Fallback: Try to extract some information about the XPS file
            try:
                import zipfile
                with zipfile.ZipFile(temp_xps, 'r') as zip_ref:
                    file_list = zip_ref.namelist()
                    logger.info(f"XPS contains {len(file_list)} files")

                    content_info = f"XPS Document: {document_name}\n\n"
                    content_info += f"File size: {len(data)} bytes\n"
                    content_info += f"Contains {len(file_list)} internal files\n\n"
                    content_info += "Structure:\n" + "\n".join(file_list[:15])
                    if len(file_list) > 15:
                        content_info += f"\n... and {len(file_list) - 15} more files"

                    temp_xps.unlink()
                    return self.create_text_pdf(content_info, output_path, f"XPS Info: {document_name}")

            except Exception as e:
                logger.warning(f"Could not read XPS structure: {e}")
                temp_xps.unlink()
                return self.create_simple_pdf(f"XPS Document: {document_name}\n\nSize: {len(data)} bytes\n\nNote: GhostScript conversion failed", output_path)

        except Exception as e:
            logger.error(f"Error processing XPS data: {e}")
            return self.create_simple_pdf(f"XPS Processing Error: {document_name}\nError: {str(e)}", output_path)

    def convert_xps_with_ghostscript(self, xps_file, output_file):
        """Convert XPS file to PDF using GhostScript."""
        try:
            # GhostScript command for XPS to PDF conversion
            gs_commands = []
            
            # If we found a GhostScript executable, use it first
            if self.gs_executable:
                gs_commands.append([self.gs_executable, '-sDEVICE=pdfwrite', f'-sOutputFile={output_file}', '-dNOPAUSE', '-dBATCH', str(xps_file)])
            
            # Fallback commands
            gs_commands.extend([
                # Try with gxps (XPS interpreter)
                ['gxpswin64.exe', '-sDEVICE=pdfwrite', f'-sOutputFile={output_file}', '-dNOPAUSE', '-dBATCH', str(xps_file)],
                ['gxps/gxpswin64.exe', '-sDEVICE=pdfwrite', f'-sOutputFile={output_file}', '-dNOPAUSE', '-dBATCH', str(xps_file)],
                # Alternative: use gswin64c with XPS device
                ['gswin64c.exe', '-sDEVICE=pdfwrite', f'-sOutputFile={output_file}', '-dNOPAUSE', '-dBATCH', str(xps_file)],
                # Another alternative with different options
                ['gswin32c.exe', '-sDEVICE=pdfwrite', f'-sOutputFile={output_file}', '-dNOPAUSE', '-dBATCH', str(xps_file)]
            ])

            for gs_command in gs_commands:
                try:
                    logger.debug(f"Trying GhostScript command: {' '.join(gs_command)}")
                    result = subprocess.run(gs_command, capture_output=True, text=True, timeout=30)

                    if result.returncode == 0 and Path(output_file).exists():
                        logger.info(f"XPS conversion successful with: {gs_command[0]}")
                        return True
                    else:
                        logger.debug(f"Command failed with return code {result.returncode}")
                        logger.debug(f"Error output: {result.stderr}")

                except FileNotFoundError:
                    logger.debug(f"Command not found: {gs_command[0]}")
                    continue
                except subprocess.TimeoutExpired:
                    logger.warning(f"GhostScript conversion timed out")
                    continue
                except Exception as e:
                    logger.debug(f"Command failed: {e}")
                    continue

            logger.warning("All GhostScript XPS conversion attempts failed")
            return False

        except Exception as e:
            logger.error(f"Error in GhostScript XPS conversion: {e}")
            return False

    def convert_image_to_pdf(self, data, output_path, document_name):
        """Convert image data to PDF."""
        try:
            # For now, create a simple PDF noting it was an image
            # You could enhance this to embed the actual image
            if data.startswith(b'\x89PNG'):
                image_type = "PNG"
            elif data.startswith(b'\xff\xd8\xff'):
                image_type = "JPEG"
            else:
                image_type = "Unknown"

            return self.create_simple_pdf(f"Image Print Job: {document_name}\n\nFormat: {image_type}\nSize: {len(data)} bytes", output_path)

        except Exception as e:
            logger.error(f"Error processing image data: {e}")
            return False

    def convert_pdf_to_image(self, pdf_path, output_path):
        """Convert PDF to image format using GhostScript. Multi-page PDFs create multiple image files.

        Returns:
            List of Path objects for created image files (in page order), or None if conversion failed
        """
        try:
            logger.info(f"Converting PDF to {self.output_format}")

            # Determine GhostScript device based on output format
            gs_device_map = {
                'PNG': 'png16m',      # 24-bit color PNG
                'JPEG': 'jpeg',       # JPEG
                'TIFF': 'tiff24nc'    # 24-bit color TIFF
            }

            # Adjust device based on color depth setting
            if self.output_format == 'PNG' and settings.IMAGE_COLOR_DEPTH == '8bit':
                gs_device = 'pnggray'
            elif self.output_format == 'PNG' and settings.IMAGE_COLOR_DEPTH == '1bit':
                gs_device = 'pngmono'
            else:
                gs_device = gs_device_map.get(self.output_format, 'png16m')

            # Prepare output filename with %d placeholder for multi-page support
            # e.g., "output.png" becomes "output_page%d.png" -> "output_page1.png", "output_page2.png"
            output_path_obj = Path(output_path)
            stem = output_path_obj.stem
            suffix = output_path_obj.suffix
            parent = output_path_obj.parent
            multi_page_output = str(parent / f"{stem}_page%d{suffix}")

            # GhostScript commands to try with improved rendering options
            gs_commands = []
            
            # If we found a GhostScript executable, use it first
            if self.gs_executable:
                gs_commands.append([self.gs_executable, '-dSAFER', '-dBATCH', '-dNOPAUSE', '-dQUIET',
                     f'-sDEVICE={gs_device}', f'-r{settings.IMAGE_DPI}',
                     '-dTextAlphaBits=4', '-dGraphicsAlphaBits=4',
                     '-dMaxBitmap=500000000', '-dAlignToPixels=0',
                     '-dGridFitTT=2', '-dPDFFitPage',
                     f'-sOutputFile={multi_page_output}', str(pdf_path)])
            
            # Fallback commands
            gs_commands.extend([
                ['gswin64c.exe', '-dSAFER', '-dBATCH', '-dNOPAUSE', '-dQUIET',
                 f'-sDEVICE={gs_device}', f'-r{settings.IMAGE_DPI}',
                 '-dTextAlphaBits=4', '-dGraphicsAlphaBits=4',
                 '-dMaxBitmap=500000000', '-dAlignToPixels=0',
                 '-dGridFitTT=2', '-dPDFFitPage',
                 f'-sOutputFile={multi_page_output}', str(pdf_path)],
                ['gswin32c.exe', '-dSAFER', '-dBATCH', '-dNOPAUSE', '-dQUIET',
                 f'-sDEVICE={gs_device}', f'-r{settings.IMAGE_DPI}',
                 '-dTextAlphaBits=4', '-dGraphicsAlphaBits=4',
                 '-dMaxBitmap=500000000', '-dAlignToPixels=0',
                 '-dGridFitTT=2', '-dPDFFitPage',
                 f'-sOutputFile={multi_page_output}', str(pdf_path)]
            ])

            for gs_command in gs_commands:
                try:
                    logger.debug(f"Trying GhostScript command: {' '.join(gs_command)}")
                    result = subprocess.run(gs_command, capture_output=True, text=True, timeout=30)

                    if result.returncode == 0:
                        # Check for generated files (could be single or multi-page)
                        page1_file = parent / f"{stem}_page1{suffix}"
                        if page1_file.exists():
                            # Multi-page output - collect all page files
                            output_files = []
                            page_num = 1
                            while True:
                                page_file = parent / f"{stem}_page{page_num}{suffix}"
                                if page_file.exists():
                                    output_files.append(page_file)
                                    page_num += 1
                                else:
                                    break
                            logger.info(f"Successfully converted {len(output_files)} page(s) to {self.output_format}")
                            logger.info(f"Output files: {stem}_page1{suffix} to {stem}_page{len(output_files)}{suffix}")
                            return output_files
                        else:
                            logger.debug(f"No output files found")
                    else:
                        logger.debug(f"Command failed with return code {result.returncode}")
                        if result.stderr:
                            logger.debug(f"GhostScript stderr: {result.stderr}")
                        if result.stdout:
                            logger.debug(f"GhostScript stdout: {result.stdout}")

                except FileNotFoundError:
                    logger.debug(f"Command not found: {gs_command[0]}")
                    continue
                except subprocess.TimeoutExpired:
                    logger.warning(f"GhostScript conversion timed out")
                    continue
                except Exception as e:
                    logger.debug(f"Command failed: {e}")
                    continue

            logger.error(f"Failed to convert PDF to {self.output_format}")
            return None

        except Exception as e:
            logger.error(f"Error converting PDF to image: {e}")
            return None

    def convert_data_to_image(self, data, output_path, document_name):
        """Try to convert raw data directly to image format.

        Returns:
            Path object of created file, or None if conversion failed
        """
        try:
            # Check if data is already an image
            if data.startswith(b'\x89PNG'):
                logger.info("Data is already PNG format")
                if self.output_format == 'PNG':
                    with open(output_path, 'wb') as f:
                        f.write(data)
                    return Path(output_path)
                else:
                    # Convert PNG to requested format
                    temp_png = self.temp_dir / f"temp_{int(time.time())}.png"
                    with open(temp_png, 'wb') as f:
                        f.write(data)
                    success = self.convert_image_format(temp_png, output_path)
                    temp_png.unlink()
                    return Path(output_path) if success else None

            elif data.startswith(b'\xff\xd8\xff'):
                logger.info("Data is already JPEG format")
                if self.output_format == 'JPEG':
                    with open(output_path, 'wb') as f:
                        f.write(data)
                    return Path(output_path)
                else:
                    # Convert JPEG to requested format
                    temp_jpg = self.temp_dir / f"temp_{int(time.time())}.jpg"
                    with open(temp_jpg, 'wb') as f:
                        f.write(data)
                    success = self.convert_image_format(temp_jpg, output_path)
                    temp_jpg.unlink()
                    return Path(output_path) if success else None
            else:
                logger.warning(f"Cannot directly convert unknown format to {self.output_format}")
                return None

        except Exception as e:
            logger.error(f"Error in direct image conversion: {e}")
            return None

    def convert_image_format(self, input_path, output_path):
        """Convert between image formats using ImageMagick or PIL."""
        try:
            # Try using PIL/Pillow first
            try:
                from PIL import Image
                img = Image.open(input_path)
                # Set DPI if supported
                if self.output_format in ['PNG', 'JPEG', 'TIFF']:
                    dpi = (settings.IMAGE_DPI, settings.IMAGE_DPI)
                    img.save(output_path, dpi=dpi)
                else:
                    img.save(output_path)
                logger.info(f"Converted image using PIL: {output_path}")
                return True
            except ImportError:
                logger.debug("PIL/Pillow not available, trying ImageMagick")
            except Exception as e:
                logger.debug(f"PIL conversion failed: {e}")

            # Fallback to ImageMagick
            convert_commands = [
                ['magick', 'convert', str(input_path), '-density', str(settings.IMAGE_DPI), str(output_path)],
                ['convert', str(input_path), '-density', str(settings.IMAGE_DPI), str(output_path)]
            ]

            for cmd in convert_commands:
                try:
                    result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
                    if result.returncode == 0 and Path(output_path).exists():
                        logger.info(f"Converted image using ImageMagick: {output_path}")
                        return True
                except (FileNotFoundError, subprocess.TimeoutExpired):
                    continue

            logger.error("Failed to convert image format - neither PIL nor ImageMagick available")
            return False

        except Exception as e:
            logger.error(f"Error converting image format: {e}")
            return False

    def extract_text_and_create_pdf(self, data, output_path, document_name):
        """Extract readable text from print data and create a PDF."""
        try:
            # Try to decode as text with various encodings
            text_content = ""

            for encoding in ['utf-8', 'latin-1', 'cp1252', 'utf-16']:
                try:
                    decoded = data.decode(encoding, errors='ignore')
                    # Filter out control characters but keep printable text
                    readable_text = ''.join(char for char in decoded if char.isprintable() or char in '\n\r\t')
                    if len(readable_text) > len(text_content):
                        text_content = readable_text
                except:
                    continue

            # If we found some readable text, use it
            if text_content and len(text_content.strip()) > 10:
                logger.info(f"Extracted {len(text_content)} characters of text")
                return self.create_text_pdf(text_content[:5000], output_path, document_name)  # Limit to first 5000 chars
            else:
                # Fallback to simple info PDF
                logger.info("No readable text found, creating info PDF")
                return self.create_simple_pdf(document_name, output_path)

        except Exception as e:
            logger.error(f"Error extracting text: {e}")
            return self.create_simple_pdf(document_name, output_path)

    def convert_html_to_pdf(self, data, output_path, document_name):
        """Convert HTML data to PDF."""
        try:
            # Save HTML to temp file
            temp_html = self.temp_dir / f"temp_{int(time.time())}.html"
            with open(temp_html, 'wb') as f:
                f.write(data)

            # Here you could use a library like wkhtmltopdf or similar
            # For now, extract text content
            html_text = data.decode('utf-8', errors='ignore')
            return self.create_text_pdf(html_text, output_path, document_name)

        except Exception as e:
            logger.error(f"Error converting HTML: {e}")
            return False

    def convert_pcl_to_pdf(self, data, output_path, document_name):
        """Convert PCL data to PDF."""
        try:
            # PCL conversion is complex - for now, create info file
            logger.info("PCL conversion not implemented, creating info PDF")
            return self.create_simple_pdf(f"PCL Document: {document_name}", output_path)

        except Exception as e:
            logger.error(f"Error converting PCL: {e}")
            return False

    def create_text_pdf(self, text_content, output_path, document_name):
        """Create a PDF with actual text content."""
        try:
            # Clean the text to remove problematic Unicode characters
            # Replace common Unicode characters with ASCII equivalents
            cleaned_text = self.clean_text_for_pdf(text_content)
            cleaned_doc_name = self.clean_text_for_pdf(document_name)

            # Calculate content length for PDF structure
            lines = cleaned_text.split('\n')[:50]  # Limit to 50 lines
            content_lines = [f"({self.escape_pdf_text(line[:80])}) Tj\nT*\n" for line in lines]
            content_text = ''.join(content_lines)

            # Create PDF with proper length calculation
            stream_content = f"""BT
/F1 10 Tf
50 750 Td
15 TL
({self.escape_pdf_text(cleaned_doc_name)}) Tj
T*
T*
{content_text}ET"""

            stream_length = len(stream_content.encode('latin-1'))

            pdf_content = f"""%PDF-1.4
1 0 obj
<<
/Type /Catalog
/Pages 2 0 R
>>
endobj

2 0 obj
<<
/Type /Pages
/Kids [3 0 R]
/Count 1
>>
endobj

3 0 obj
<<
/Type /Page
/Parent 2 0 R
/MediaBox [0 0 612 792]
/Contents 4 0 R
/Resources <<
  /Font <<
    /F1 <<
      /Type /Font
      /Subtype /Type1
      /BaseFont /Helvetica
    >>
  >>
>>
>>
endobj

4 0 obj
<<
/Length {stream_length}
>>
stream
{stream_content}
endstream
endobj

xref
0 5
0000000000 65535 f 
0000000009 00000 n 
0000000058 00000 n 
0000000115 00000 n 
0000000273 00000 n 
trailer
<<
/Size 5
/Root 1 0 R
>>
startxref
{400 + stream_length}
%%EOF"""

            # Write as binary to avoid encoding issues
            with open(output_path, 'wb') as f:
                f.write(pdf_content.encode('latin-1'))
                f.flush()
                os.fsync(f.fileno())

            logger.info(f"Created text PDF with {len(lines)} lines: {output_path}")
            return True

        except Exception as e:
            logger.error(f"Error creating text PDF: {e}")
            # Fallback to simple PDF without text content
            return self.create_simple_pdf(document_name, output_path)

    def clean_text_for_pdf(self, text):
        """Clean text to remove problematic Unicode characters for PDF."""
        if not text:
            return ""

        # Dictionary of common Unicode replacements
        unicode_replacements = {
            '\u2013': '-',     # en dash
            '\u2014': '--',    # em dash
            '\u2018': "'",     # left single quotation mark
            '\u2019': "'",     # right single quotation mark
            '\u201c': '"',     # left double quotation mark
            '\u201d': '"',     # right double quotation mark
            '\u2020': '+',     # dagger
            '\u2021': '++',    # double dagger
            '\u2022': '*',     # bullet
            '\u2026': '...',   # horizontal ellipsis
            '\u00a0': ' ',     # non-breaking space
            '\u00b0': 'deg',   # degree symbol
            '\u00a9': '(c)',   # copyright
            '\u00ae': '(R)',   # registered
            '\u2122': '(TM)',  # trademark
        }

        # Replace known Unicode characters
        for unicode_char, replacement in unicode_replacements.items():
            text = text.replace(unicode_char, replacement)

        # Remove any remaining non-ASCII characters
        cleaned = ''.join(char if ord(char) < 256 else '?' for char in text)

        return cleaned

    def escape_pdf_text(self, text):
        """Escape special characters for PDF text streams."""
        if not text:
            return ""

        # Escape PDF special characters
        text = text.replace('\\', '\\\\')  # Backslash must be first
        text = text.replace('(', '\\(')    # Left parenthesis
        text = text.replace(')', '\\)')    # Right parenthesis
        text = text.replace('\r', '\\r')   # Carriage return
        text = text.replace('\t', '    ')  # Tab to spaces

        return text

    def convert_ps_data_to_pdf(self, ps_data, output_path):
        """Convert PostScript data to PDF."""
        try:
            # Write PostScript to temp file
            temp_ps = self.temp_dir / f"temp_{int(time.time())}.ps"
            with open(temp_ps, 'wb') as f:
                f.write(ps_data)

            # Try to convert with GhostScript
            if self.convert_ps_with_ghostscript_file(temp_ps, output_path):
                temp_ps.unlink()  # Clean up
                return True

            # Fallback - just copy as .ps file with PDF extension for now
            shutil.copy2(temp_ps, output_path.with_suffix('.ps'))
            temp_ps.unlink()
            logger.info(f"Saved PostScript data as .ps file: {output_path.with_suffix('.ps')}")
            return True

        except Exception as e:
            logger.error(f"Error converting PostScript data: {e}")
            return False

    def convert_ps_with_ghostscript_file(self, ps_file, output_file):
        """Convert PostScript file to PDF using GhostScript."""
        try:
            gs_exe = self.gs_executable if self.gs_executable else 'gswin64c.exe'
            gs_command = [
                gs_exe,
                '-dNOPAUSE',
                '-dBATCH', 
                '-sDEVICE=pdfwrite',
                f'-sOutputFile={output_file}',
                str(ps_file)
            ]

            result = subprocess.run(gs_command, capture_output=True, text=True)
            return result.returncode == 0

        except Exception:
            return False

    def process_virtual_print_job(self, job):
        """Process a print job from our virtual printer and convert to PDF."""
        try:
            job_id = job['JobId']
            document_name = job.get('pDocument', f"print_job_{job_id}")
            logger.info(f"Processing virtual print job: {document_name}")

            # Create output filename with timestamp
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            safe_doc_name = "".join(c for c in document_name if c.isalnum() or c in (' ', '-', '_')).strip()
            output_filename = f"{timestamp}_{safe_doc_name}.pdf"
            output_path = self.output_dir / output_filename

            # Read the raw print data from the spooler
            raw_data = self.get_job_data(job_id)

            if raw_data:
                # Convert PostScript/raw data to PDF
                success = self.convert_raw_to_pdf(raw_data, output_path, document_name)

                if success:
                    logger.info(f"Successfully created PDF: {output_path}")
                    # Delete the job from queue since we've processed it
                    self.delete_print_job(job_id)
                else:
                    logger.warning("PDF conversion failed, creating fallback file")
                    self.create_job_info_file(job, output_path.with_suffix('.txt'))
            else:
                logger.warning("Could not retrieve job data")
                self.create_job_info_file(job, output_path.with_suffix('.txt'))

        except Exception as e:
            logger.error(f"Error processing virtual print job: {e}")

    def get_job_data(self, job_id):
        """Extract raw print data from a print job."""
        try:
            handle = win32print.OpenPrinter(self.printer_name)

            # Try to read the job data
            # Note: This is simplified - real implementation would need 
            # to handle the Windows spooler API more carefully
            job_info = win32print.GetJob(handle, job_id, 1)

            win32print.ClosePrinter(handle)

            # For now, return a placeholder - in a real implementation
            # you'd extract the actual print data from the spooler
            return b"PostScript data placeholder"

        except Exception as e:
            logger.error(f"Error getting job data: {e}")
            return None

    def delete_print_job(self, job_id):
        """Delete a processed print job from the queue."""
        try:
            handle = win32print.OpenPrinter(self.printer_name)
            win32print.SetJob(handle, job_id, 0, None, win32print.JOB_CONTROL_DELETE)
            win32print.ClosePrinter(handle)
            logger.debug(f"Deleted processed job {job_id}")
        except Exception as e:
            logger.error(f"Error deleting job {job_id}: {e}")

    def convert_raw_to_pdf(self, raw_data, output_path, document_name):
        """Convert raw print data (PostScript) to PDF."""
        try:
            # Method 1: Use GhostScript to convert PostScript to PDF
            if self.convert_ps_with_ghostscript(raw_data, output_path):
                return True

            # Method 2: Create a simple PDF with the document info
            return self.create_simple_pdf(document_name, output_path)

        except Exception as e:
            logger.error(f"Error converting raw data to PDF: {e}")
            return False

    def convert_ps_with_ghostscript(self, ps_data, output_path):
        """Convert PostScript data to PDF using GhostScript."""
        try:
            # Write PostScript data to temporary file
            temp_ps = self.temp_dir / f"temp_{int(time.time())}.ps"
            with open(temp_ps, 'wb') as f:
                f.write(ps_data)

            # Convert using GhostScript
            gs_exe = self.gs_executable if self.gs_executable else 'gswin64c.exe'
            gs_command = [
                gs_exe,
                '-dNOPAUSE',
                '-dBATCH',
                '-sDEVICE=pdfwrite',
                f'-sOutputFile={output_path}',
                str(temp_ps)
            ]

            result = subprocess.run(gs_command, capture_output=True, text=True)

            # Clean up temp file
            temp_ps.unlink()

            return result.returncode == 0

        except Exception as e:
            logger.warning(f"GhostScript conversion failed: {e}")
            return False

    def create_simple_pdf(self, document_name, output_path):
        """Create a simple PDF document as fallback."""
        try:
            # This is a very basic PDF creation - you might want to use
            # a library like reportlab for more sophisticated PDFs
            pdf_content = f"""%PDF-1.4
1 0 obj
<<
/Type /Catalog
/Pages 2 0 R
>>
endobj

2 0 obj
<<
/Type /Pages
/Kids [3 0 R]
/Count 1
>>
endobj

3 0 obj
<<
/Type /Page
/Parent 2 0 R
/MediaBox [0 0 612 792]
/Contents 4 0 R
/Resources <<
  /Font <<
    /F1 <<
      /Type /Font
      /Subtype /Type1
      /BaseFont /Helvetica
    >>
  >>
>>
>>
endobj

4 0 obj
<<
/Length 58
>>
stream
BT
/F1 12 Tf
100 700 Td
(Virtual Print Job: {document_name}) Tj
ET
endstream
endobj

xref
0 5
0000000000 65535 f 
0000000009 00000 n 
0000000058 00000 n 
0000000115 00000 n 
0000000205 00000 n 
trailer
<<
/Size 5
/Root 1 0 R
>>
startxref
314
%%EOF"""

            with open(output_path, 'wb') as f:
                f.write(pdf_content.encode('latin-1'))
                f.flush()
                os.fsync(f.fileno())

            logger.info(f"Created simple PDF: {output_path}")
            return True

        except Exception as e:
            logger.error(f"Error creating simple PDF: {e}")
            return False

    def create_job_info_file(self, job, output_file):
        """Create a text file with job information as fallback."""
        try:
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write("VIRTUAL PRINTER JOB CAPTURE\n")
                f.write("=" * 40 + "\n\n")
                f.write(f"Job ID: {job.get('JobId', 'Unknown')}\n")
                f.write(f"Document: {job.get('pDocument', 'Unknown')}\n")
                f.write(f"User: {job.get('pUserName', 'Unknown')}\n")
                f.write(f"Machine: {job.get('pMachineName', 'Unknown')}\n")
                f.write(f"Status: {job.get('Status', 'Unknown')}\n")
                f.write(f"Pages: {job.get('TotalPages', 'Unknown')}\n")
                f.write(f"Submitted: {job.get('Submitted', 'Unknown')}\n")
                f.write(f"Data Type: {job.get('pDatatype', 'Unknown')}\n")
                f.write(f"\nCaptured: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")

            logger.info(f"Created job info file: {output_file}")

        except Exception as e:
            logger.error(f"Error creating job info file: {e}")

    def start_monitoring(self):
        """Start the print job monitoring in a separate thread."""
        if self.running:
            logger.warning("Monitoring is already running.")
            return

        self.running = True
        self.monitor_thread = threading.Thread(target=self.monitor_print_jobs)
        self.monitor_thread.daemon = True
        self.monitor_thread.start()
        logger.info("Print job monitoring started.")

    def stop_monitoring(self):
        """Stop the print job monitoring."""
        self.running = False
        if self.monitor_thread:
            self.monitor_thread.join(timeout=5)
        logger.info("Print job monitoring stopped.")

    def list_available_printers(self):
        """List all available printers on the system."""
        try:
            printers = win32print.EnumPrinters(2)
            logger.info("Available printers:")
            for printer in printers:
                print(f"  - {printer[2]}")
            return printers
        except Exception as e:
            logger.error(f"Error listing printers: {e}")
            return []


def main():
    """Main function to run the virtual printer."""
    print("VirtPrint - Virtual Printer")
    print("=" * 40)

    # Check if running as administrator
    try:
        is_admin = win32api.GetUserName() != os.getlogin()
    except:
        is_admin = False

    if not is_admin:
        print("\nWARNING: This program may require administrator privileges to install printers.")
        print("If you encounter permission errors, try running as administrator.")

    # Create virtual printer instance
    vp = VirtualPrinter()

    try:
        # Install the printer
        if not vp.install_printer():
            print("Failed to install virtual printer. Exiting.")
            return 1

        print(f"\nVirtual printer '{vp.printer_name}' is ready!")
        print(f"Output format: {vp.output_format}")
        print(f"Print jobs will be saved to: {vp.output_dir}")
        if vp.output_format != 'PDF':
            print(f"Image resolution: {settings.IMAGE_DPI} DPI")
        print("\nStarting print job monitoring...")

        # Start monitoring
        vp.start_monitoring()

        print("\nPress Ctrl+C to stop monitoring and exit.")

        # Keep the program running
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            print("\nShutting down...")

    except Exception as e:
        logger.error(f"Error in main: {e}")
        return 1

    finally:
        # Clean up
        vp.stop_monitoring()
        try:
            vp.uninstall_printer()
        except:
            pass

    return 0


if __name__ == "__main__":
    sys.exit(main())