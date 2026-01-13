#!/usr/bin/env python3
"""
TCP Virtual Printer for Windows
Creates a network printer that listens on a TCP port and converts print jobs to files.
Supports multiple output formats: PDF, PNG, JPEG, TIFF, PS, RAW
"""

import os
import sys
import socket
import threading
import time
from datetime import datetime
from pathlib import Path
import logging
import subprocess
import tempfile
import shutil
import settings

from utils import add_tcp_printer

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
    """TCP-based virtual printer that listens on a network port."""
    
    def __init__(self, printer_name=None, host=None, port=None, output_dir=None, output_format=None):
        self.printer_name = printer_name or settings.PRINTER_NAME
        self.host = host or settings.TCP_HOST
        self.port = port or settings.TCP_PORT
        self.output_dir = Path(output_dir or settings.OUTPUT_DIR)
        self.output_dir.mkdir(exist_ok=True)
        self.output_format = (output_format or settings.OUTPUT_FORMAT).upper()
        
        self.running = False
        self.server_socket = None
        self.server_thread = None
        self.job_counter = 0
        self.job_counter_lock = threading.Lock()
        
        # Ensure temp directory exists
        self.temp_dir = Path(tempfile.gettempdir()) / "virtprint"
        self.temp_dir.mkdir(exist_ok=True)
        
        # Find GhostScript executable for conversions
        self.gs_executable = self.find_ghostscript_executable()
        if self.gs_executable:
            logger.info(f"Found GhostScript at: {self.gs_executable}")
        else:
            logger.warning("GhostScript not found - image conversion may not work")
        
        # Validate output format
        supported_formats = ['PDF', 'PS', 'RAW', 'PNG', 'JPEG', 'TIFF']
        if self.output_format not in supported_formats:
            logger.warning(f"Unsupported format {self.output_format}, defaulting to PDF")
            self.output_format = 'PDF'
        
        logger.info(f"TCP Virtual Printer initialized: {self.host}:{self.port}")
        logger.info(f"Output directory: {self.output_dir}")
        logger.info(f"Output format: {self.output_format}")
        if self.output_format in ['PNG', 'JPEG', 'TIFF']:
            logger.info(f"Image DPI: {settings.IMAGE_DPI}")
    
    def find_ghostscript_executable(self):
        """Find the GhostScript executable."""
        script_dir = Path(__file__).parent.resolve()
        
        # Check in gxps folder first
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
        
        # Check system PATH
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
    
    def install_printer(self):
        add_tcp_printer(printer_name=self.printer_name,
                        driver_name=settings.DRIVER_NAME,
                        port_name=settings.WINDOWS_PRINTER_PORT_NAME,
                        ip_address=self.host,
                        tcp_port=self.port)

    def uninstall_printer(self):
        """Remove the TCP virtual printer."""
        try:
            import win32print
            
            handle = win32print.OpenPrinter(self.printer_name)
            win32print.DeletePrinter(handle)
            win32print.ClosePrinter(handle)
            logger.info(f"Successfully uninstalled printer '{self.printer_name}'")
            return True
        except ImportError:
            logger.warning("win32print not available - cannot uninstall printer")
            return False
        except Exception as e:
            logger.error(f"Failed to uninstall printer: {e}")
            return False

    def start(self):
        """Start the TCP server and begin listening for print jobs."""
        if self.running:
            logger.warning("TCP server is already running")
            return
        
        try:
            # Create server socket
            self.server_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.server_socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            
            # Bind and listen
            self.server_socket.bind((self.host, self.port))
            self.server_socket.listen(5)
            
            # Set non-blocking mode with timeout for responsiveness
            self.server_socket.settimeout(0.5)
            
            self.running = True
            
            # Start server thread
            self.server_thread = threading.Thread(target=self._server_loop, daemon=False)
            self.server_thread.start()
            
            logger.info(f"TCP server started on {self.host}:{self.port}")
            logger.info("Waiting for print jobs...")
            
        except Exception as e:
            logger.error(f"Failed to start TCP server: {e}")
            self.running = False
            raise
    
    def stop(self):
        """Stop the TCP server."""
        if not self.running:
            return
        
        logger.info("Stopping TCP server...")
        self.running = False
        
        # Close server socket
        if self.server_socket:
            try:
                self.server_socket.close()
            except:
                pass
        
        # Wait for server thread to finish
        if self.server_thread and self.server_thread.is_alive():
            self.server_thread.join(timeout=5)
        
        logger.info("TCP server stopped")
    
    def _server_loop(self):
        """Main server loop - accepts connections and handles print jobs."""
        connection_count = 0
        logger.info("Server loop started and ready to accept connections")
        
        while self.running:
            try:
                try:
                    client_socket, client_address = self.server_socket.accept()
                    connection_count += 1
                    logger.info(f"Connection #{connection_count} from {client_address}")
                    
                    # Set socket options for client connection
                    client_socket.setsockopt(socket.SOL_SOCKET, socket.SO_KEEPALIVE, 1)
                    client_socket.setsockopt(socket.IPPROTO_TCP, socket.TCP_NODELAY, 1)
                    
                    # Handle this connection in a separate thread
                    handler_thread = threading.Thread(
                        target=self._handle_client,
                        args=(client_socket, client_address),
                        daemon=True
                    )
                    handler_thread.start()
                    
                except socket.timeout:
                    # No connection yet, continue loop
                    continue
                except OSError as e:
                    # Socket was closed
                    if not self.running:
                        break
                    logger.error(f"Socket error in accept: {e}")
                    time.sleep(0.1)
                    
            except Exception as e:
                if self.running:
                    logger.error(f"Error in server loop: {e}", exc_info=True)
                    time.sleep(1)
        
        logger.info("Server loop ended")
    
    def _handle_client(self, client_socket, client_address):
        """Handle a single client connection (print job)."""
        data = b""
        
        try:
            # Set timeout for receiving data
            client_socket.settimeout(10.0)
            
            logger.debug(f"Reading data from {client_address}")
            
            # Receive data from client
            while True:
                try:
                    chunk = client_socket.recv(8192)  # Increased buffer size
                    if not chunk:
                        break
                    data += chunk
                    logger.debug(f"Received {len(chunk)} bytes (total: {len(data)})")
                except socket.timeout:
                    # No more data after timeout
                    if len(data) > 0:
                        logger.debug("Timeout reached, assuming data is complete")
                        break
                    else:
                        logger.warning(f"Timeout with no data from {client_address}")
                        return
            
            if len(data) > 0:
                logger.info(f"Received {len(data)} bytes from {client_address}")
                self._process_print_job(data, client_address)
            else:
                logger.warning(f"No data received from {client_address}")
            
        except Exception as e:
            logger.error(f"Error handling client {client_address}: {e}", exc_info=True)
        finally:
            try:
                client_socket.close()
                logger.debug(f"Closed connection from {client_address}")
            except Exception as e:
                logger.error(f"Error handling client {client_address}: {e}", exc_info=True)

    
    def _process_print_job(self, data, client_address):
        """Process received print job data."""
        try:
            # Generate unique job ID
            with self.job_counter_lock:
                self.job_counter += 1
                job_id = self.job_counter
            
            # Create output filename
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            client_ip = client_address[0] if client_address else "unknown"
            
            # Job info for callback
            job_info = {
                'job_id': job_id,
                'timestamp': timestamp,
                'client_address': client_address,
                'data_size': len(data),
                'output_format': self.output_format
            }
            
            logger.info(f"Processing job {job_id} from {client_ip}")
            
            # Save based on output format
            output_files = []
            
            if self.output_format == 'RAW':
                # Save raw data
                output_filename = f"{timestamp}_job{job_id}_raw.prn"
                output_path = self.output_dir / output_filename
                
                with open(output_path, 'wb') as f:
                    f.write(data)
                
                output_files.append(output_path)
                logger.info(f"Saved raw data to {output_path}")
                
            elif self.output_format == 'PS':
                # Save as PostScript
                output_filename = f"{timestamp}_job{job_id}.ps"
                output_path = self.output_dir / output_filename
                
                with open(output_path, 'wb') as f:
                    f.write(data)
                
                output_files.append(output_path)
                logger.info(f"Saved PostScript to {output_path}")
                
            elif self.output_format in ['PDF', 'PNG', 'JPEG', 'TIFF']:
                # Convert using GhostScript
                output_files = self._convert_to_format(data, job_id, timestamp)
                
            else:
                logger.error(f"Unknown output format: {self.output_format}")
                output_files = None
            
            # Call user callback
            try:
                on_print_job_complete(output_files, job_info)
            except Exception as e:
                logger.error(f"Error in user callback: {e}")
            
            if output_files:
                logger.info(f"Job {job_id} completed successfully: {len(output_files)} file(s)")
            else:
                logger.warning(f"Job {job_id} completed with no output")
                
        except Exception as e:
            logger.error(f"Error processing print job: {e}")
            # Call callback with None to indicate failure
            try:
                on_print_job_complete(None, job_info)
            except:
                pass
    
    def _convert_to_format(self, data, job_id, timestamp):
        """Convert print data to the configured output format using GhostScript."""
        if not self.gs_executable:
            logger.error("GhostScript not available for conversion")
            # Fallback to raw
            output_filename = f"{timestamp}_job{job_id}_raw.prn"
            output_path = self.output_dir / output_filename
            with open(output_path, 'wb') as f:
                f.write(data)
            return [output_path]
        
        try:
            # Detect input format from data
            data_format = self._detect_format(data)
            logger.info(f"Detected input format: {data_format}")
            
            # Handle plain text - save as text file
            if data_format == 'TEXT':
                output_filename = f"{timestamp}_job{job_id}.txt"
                output_path = self.output_dir / output_filename
                with open(output_path, 'wb') as f:
                    f.write(data)
                logger.info(f"Saved plain text to {output_path}")
                return [output_path]
            
            # Save data to temp file with appropriate extension
            if data_format == 'XPS':
                temp_input = self.temp_dir / f"job{job_id}_input.xps"
            elif data_format == 'PDF':
                temp_input = self.temp_dir / f"job{job_id}_input.pdf"
            else:  # PostScript or unknown
                temp_input = self.temp_dir / f"job{job_id}_input.ps"
            
            with open(temp_input, 'wb') as f:
                f.write(data)
            
            # If already in desired format, just copy it
            if data_format == 'PDF' and self.output_format == 'PDF':
                output_filename = f"{timestamp}_job{job_id}.pdf"
                output_path = self.output_dir / output_filename
                shutil.copy2(temp_input, output_path)
                temp_input.unlink()
                return [output_path]
            
            # Determine GhostScript device and output based on format and color depth
            if self.output_format == 'PDF':
                device = 'pdfwrite'
                output_filename = f"{timestamp}_job{job_id}.pdf"
            elif self.output_format == 'PNG':
                # Adjust PNG device based on color depth setting
                if settings.IMAGE_COLOR_DEPTH == '8bit':
                    device = 'pnggray'
                elif settings.IMAGE_COLOR_DEPTH == '1bit':
                    device = 'pngmono'
                else:
                    device = 'png16m'
                output_filename = f"{timestamp}_job{job_id}_%03d.png"
            elif self.output_format == 'JPEG':
                device = 'jpeg'
                output_filename = f"{timestamp}_job{job_id}_%03d.jpg"
            elif self.output_format == 'TIFF':
                device = 'tiff24nc'
                output_filename = f"{timestamp}_job{job_id}_%03d.tiff"
            else:
                device = 'pdfwrite'
                output_filename = f"{timestamp}_job{job_id}.pdf"
            
            output_path = self.output_dir / output_filename
            
            # Choose correct GhostScript executable based on input format
            gs_exe = self._get_gs_executable_for_format(data_format)
            if not gs_exe:
                logger.error(f"No suitable GhostScript executable found for {data_format}")
                # Fallback to raw
                fallback_path = self.output_dir / f"{timestamp}_job{job_id}_raw.prn"
                shutil.copy2(temp_input, fallback_path)
                temp_input.unlink()
                return [fallback_path]
            
            # Build GhostScript command with enhanced options
            gs_args = [
                gs_exe,
                '-dNOPAUSE',
                '-dBATCH',
                '-dSAFER',
                '-dQUIET',
                f'-sDEVICE={device}',
                f'-r{settings.IMAGE_DPI}',
            ]
            
            # Add anti-aliasing only for raster devices (not for vector devices like PDF/PS)
            if self.output_format in ['PNG', 'JPEG', 'TIFF']:
                gs_args.extend([
                    '-dTextAlphaBits=4',
                    '-dGraphicsAlphaBits=4',
                ])
            
            gs_args.extend([
                f'-sOutputFile={output_path}',
                str(temp_input)
            ])
            
            logger.info(f"Converting with GhostScript: {' '.join(gs_args)}")
            
            # Run GhostScript
            result = subprocess.run(
                gs_args,
                capture_output=True,
                text=True,
                timeout=60
            )
            
            # Log GhostScript output for debugging
            if result.stdout:
                logger.debug(f"GhostScript stdout: {result.stdout}")
            if result.stderr:
                logger.warning(f"GhostScript stderr: {result.stderr}")
            
            if result.returncode != 0:
                logger.error(f"GhostScript failed with return code {result.returncode}")
                logger.error(f"GhostScript stderr: {result.stderr}")
                # Save raw data as fallback
                fallback_path = self.output_dir / f"{timestamp}_job{job_id}_raw.prn"
                with open(fallback_path, 'wb') as f:
                    f.write(data)
                logger.info(f"Saved raw data as fallback: {fallback_path}")
                return [fallback_path]
            
            # Find all generated files (for multi-page images)
            if '%' in output_filename:
                # Multi-page output
                base_pattern = output_filename.replace('%03d', '*')
                output_files = list(self.output_dir.glob(base_pattern))
                if output_files:
                    logger.info(f"Generated {len(output_files)} file(s)")
                else:
                    logger.error(f"No output files found matching pattern: {base_pattern}")
                    logger.info(f"Expected pattern: {output_path}")
                    # List files in output directory for debugging
                    all_files = list(self.output_dir.glob(f"{timestamp}_job{job_id}*"))
                    logger.debug(f"Files in output dir matching job: {all_files}")
            else:
                if output_path.exists():
                    output_files = [output_path]
                    logger.info(f"Generated output file: {output_path}")
                else:
                    logger.error(f"Expected output file not found: {output_path}")
                    output_files = []
            
            # Clean up temp file
            try:
                temp_input.unlink()
            except:
                pass
            
            if output_files:
                return output_files
            else:
                logger.error("No output files were created by GhostScript")
                # Save raw data as fallback
                fallback_path = self.output_dir / f"{timestamp}_job{job_id}_raw.prn"
                with open(fallback_path, 'wb') as f:
                    f.write(data)
                logger.info(f"Saved raw data as fallback: {fallback_path}")
                return [fallback_path]
                
        except Exception as e:
            logger.error(f"Error converting to {self.output_format}: {e}")
            return None
    
    def _detect_format(self, data):
        """Detect the format of print data."""
        if not data or len(data) < 10:
            return 'UNKNOWN'
        
        # Check first bytes for format signatures
        if data.startswith(b'%PDF'):
            return 'PDF'
        elif data.startswith(b'%!PS') or b'%!PS-Adobe' in data[:100]:
            return 'PS'
        elif data.startswith(b'PK') or b'FixedDocument' in data[:1000]:
            return 'XPS'
        elif data.startswith(b'\x1b'):
            return 'PCL'
        else:
            # Check if it's plain text (all printable ASCII/UTF-8)
            try:
                # Try to decode as text
                text = data[:500].decode('utf-8', errors='ignore')
                if len([c for c in text if c.isprintable() or c in '\n\r\t']) > len(text) * 0.9:
                    return 'TEXT'
            except:
                pass
            return 'UNKNOWN'
    
    def _get_gs_executable_for_format(self, data_format):
        """Get the appropriate GhostScript executable for the input format."""
        script_dir = Path(__file__).parent.resolve()
        gxps_dir = script_dir / "gxps"
        
        if data_format == 'XPS':
            # Use gxps for XPS files
            for exe in [gxps_dir / "gxpswin64.exe", gxps_dir / "gxpswin32.exe"]:
                if exe.exists():
                    return str(exe)
        else:
            # Use regular GhostScript for PS/PDF
            for exe in [gxps_dir / "gswin64c.exe", gxps_dir / "gswin32c.exe"]:
                if exe.exists():
                    return str(exe)
            
            # Try system PATH
            for exe_name in ["gswin64c.exe", "gswin32c.exe", "gs.exe"]:
                try:
                    result = subprocess.run(["where", exe_name], capture_output=True, text=True)
                    if result.returncode == 0:
                        path = result.stdout.strip().split('\n')[0]
                        if path:
                            return path
                except:
                    pass
        
        # Fallback to the initially found executable
        return self.gs_executable
    
    def start_monitoring(self):
        """Alias for start() - for compatibility with service.py."""
        self.start()
    
    def stop_monitoring(self):
        """Alias for stop() - for compatibility with service.py."""
        self.stop()
    
    def run(self):
        """Run the TCP printer server (blocking)."""
        try:
            self.start()
            
            # Keep running until interrupted
            while self.running:
                time.sleep(1)
                
        except KeyboardInterrupt:
            logger.info("Received keyboard interrupt")
        finally:
            self.stop()


def main():
    """Main entry point."""
    import argparse
    
    parser = argparse.ArgumentParser(description='TCP Virtual Printer')
    parser.add_argument('--host', default=settings.TCP_HOST, help='TCP host to bind to')
    parser.add_argument('--port', type=int, default=settings.TCP_PORT, help='TCP port to listen on')
    parser.add_argument('--output-dir', default=settings.OUTPUT_DIR, help='Output directory for print jobs')
    parser.add_argument('--output-format', default=settings.OUTPUT_FORMAT, 
                       choices=['PDF', 'PS', 'RAW', 'PNG', 'JPEG', 'TIFF'],
                       help='Output format')
    parser.add_argument('--install', action='store_true', help='Install Windows printer')
    parser.add_argument('--uninstall', action='store_true', help='Uninstall Windows printer')
    
    args = parser.parse_args()
    
    # Create printer instance
    printer = VirtualPrinter(
        host=args.host,
        port=args.port,
        output_dir=args.output_dir,
        output_format=args.output_format
    )
    
    # Handle install/uninstall
    if args.install:
        if printer.install_printer():
            logger.info("Printer installed successfully")
            logger.info(f"Configure the printer to use Standard TCP/IP Port: {args.host}:{args.port}")
        else:
            logger.error("Failed to install printer")
        return
    
    if args.uninstall:
        if printer.uninstall_printer():
            logger.info("Printer uninstalled successfully")
        else:
            logger.error("Failed to uninstall printer")
        return
    
    # Always try to install/update the printer before starting the server
    try:
        logger.info("Installing/updating printer...")
        printer.install_printer()
        logger.info("Printer installation/update completed")
    except Exception as e:
        logger.warning(f"Failed to install printer (continuing anyway): {e}")
    
    # Run the server
    logger.info("Starting TCP Virtual Printer...")
    logger.info(f"Listening on {args.host}:{args.port}")
    logger.info(f"Output directory: {args.output_dir}")
    logger.info(f"Output format: {args.output_format}")
    logger.info("Press Ctrl+C to stop")
    
    printer.run()


if __name__ == '__main__':
    main()
