#!/usr/bin/env python3
"""
VirtPrint Windows Service
Runs the virtual printer as a Windows service.
"""

import sys
import os
from pathlib import Path

# CRITICAL: Add venv paths BEFORE any other imports
# This allows the service to find packages installed in the venv
SCRIPT_DIR = Path(__file__).parent.resolve()
VENV_DIR = SCRIPT_DIR / ".venv"

# Add venv site-packages to sys.path
if VENV_DIR.exists():
    venv_site_packages = VENV_DIR / "Lib" / "site-packages"
    if venv_site_packages.exists():
        sys.path.insert(0, str(venv_site_packages))
    
    # Also add the Scripts directory for any executables
    venv_scripts = VENV_DIR / "Scripts"
    if venv_scripts.exists():
        os.environ['PATH'] = str(venv_scripts) + os.pathsep + os.environ.get('PATH', '')

# Add the script directory to the path so we can import virtprint module
sys.path.insert(0, str(SCRIPT_DIR))

# Now import everything else
import logging
import win32serviceutil
import win32service
import win32event
import servicemanager

import settings
from virtprint import VirtualPrinter


class VirtPrintService(win32serviceutil.ServiceFramework):
    """Windows Service for VirtPrint virtual printer."""
    
    _svc_name_ = "VirtPrint"
    _svc_display_name_ = "VirtPrint Virtual Printer Service"
    _svc_description_ = "Virtual printer service that converts print jobs to PDF/image files"
    
    # Use the venv Python interpreter
    _exe_name_ = str(VENV_DIR / "Scripts" / "python.exe") if VENV_DIR.exists() else sys.executable
    _exe_args_ = f'"{Path(__file__)}"'
    
    def __init__(self, args):
        """Initialize the service."""
        win32serviceutil.ServiceFramework.__init__(self, args)
        
        # Create an event to listen for stop requests
        self.stop_event = win32event.CreateEvent(None, 0, 0, None)
        
        # Initialize basic attributes - don't do any heavy lifting here
        self.printer = None
        self.is_running = False
        self.logger = None
        
    def setup_logging(self):
        """Configure logging for the service."""
        try:
            # Use absolute path for log file
            log_file = Path(settings.LOG_FILE)
            if not log_file.is_absolute():
                log_file = SCRIPT_DIR / settings.LOG_FILE
            
            # Ensure log directory exists
            log_file.parent.mkdir(parents=True, exist_ok=True)
            
            # Configure logging
            logging.basicConfig(
                level=getattr(logging, settings.LOG_LEVEL, logging.INFO),
                format=settings.LOG_FORMAT,
                handlers=[
                    logging.FileHandler(str(log_file)),
                ],
                force=True  # Override any existing configuration
            )
            self.logger = logging.getLogger(__name__)
            self.logger.info(f"Logging initialized to: {log_file}")
        except Exception as e:
            # If logging fails, create a basic logger
            self.logger = logging.getLogger(__name__)
            try:
                servicemanager.LogErrorMsg(f"VirtPrint: Failed to setup logging: {e}")
            except:
                pass
        
    def SvcStop(self):
        """Handle stop request from Windows Service Manager."""
        try:
            if self.logger:
                self.logger.info("VirtPrint service stop requested")
            self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
            
            # Signal the stop event
            win32event.SetEvent(self.stop_event)
            
            # Stop the printer monitor
            self.is_running = False
            if self.printer:
                try:
                    self.printer.stop_monitoring()
                    if self.logger:
                        self.logger.info("Printer monitor stopped successfully")
                except Exception as e:
                    if self.logger:
                        self.logger.error(f"Error stopping printer monitor: {e}")
        except Exception as e:
            try:
                servicemanager.LogErrorMsg(f"VirtPrint SvcStop error: {e}")
            except:
                pass
                
    def SvcDoRun(self):
        """Main service execution method."""
        # CRITICAL: Report running status IMMEDIATELY before doing anything
        self.ReportServiceStatus(win32service.SERVICE_RUNNING)
        
        # Now setup logging (after reporting running)
        try:
            self.setup_logging()
        except Exception as e:
            try:
                servicemanager.LogErrorMsg(f"VirtPrint: Failed to setup logging: {e}")
            except:
                pass
        
        # Log service start to the Windows Event Log
        try:
            servicemanager.LogMsg(
                servicemanager.EVENTLOG_INFORMATION_TYPE,
                servicemanager.PYS_SERVICE_STARTED,
                (self._svc_name_, '')
            )
        except:
            pass
        
        # Start initialization in a separate thread to not block the service
        import threading
        init_thread = threading.Thread(target=self._initialize_printer)
        init_thread.daemon = False
        init_thread.start()
        
        # Wait for stop signal
        win32event.WaitForSingleObject(self.stop_event, win32event.INFINITE)
        
        if self.logger:
            self.logger.info("VirtPrint service stopped")
        
        # Log service stop to the Windows Event Log
        try:
            servicemanager.LogMsg(
                servicemanager.EVENTLOG_INFORMATION_TYPE,
                servicemanager.PYS_SERVICE_STOPPED,
                (self._svc_name_, '')
            )
        except:
            pass
    
    def _initialize_printer(self):
        """Initialize and start the printer in a background thread."""
        try:
            if self.logger:
                self.logger.info("="*60)
                self.logger.info("VirtPrint service initializing...")
                self.logger.info(f"Script directory: {SCRIPT_DIR}")
                self.logger.info(f"Printer name: {settings.PRINTER_NAME}")
                self.logger.info(f"Output directory: {settings.OUTPUT_DIR}")
                self.logger.info(f"Output format: {settings.OUTPUT_FORMAT}")
                self.logger.info("="*60)
            
            # Create VirtualPrinter instance
            self.is_running = True
            if self.logger:
                self.logger.info("Creating VirtualPrinter instance...")
            
            self.printer = VirtualPrinter(
                printer_name=settings.PRINTER_NAME,
                output_dir=settings.OUTPUT_DIR,
                output_format=settings.OUTPUT_FORMAT
            )
            
            # Install the printer if needed
            if self.logger:
                self.logger.info("Installing printer if needed...")
            
            if not self.printer.install_printer():
                if self.logger:
                    self.logger.error("Failed to install printer")
                servicemanager.LogErrorMsg("VirtPrint: Failed to install printer")
                return  # Exit the initialization thread
            
            # Start the printer monitoring
            if self.logger:
                self.logger.info("Starting printer monitor...")
            self.printer.start_monitoring()
            if self.logger:
                self.logger.info("Printer monitor is now running")
                
        except Exception as e:
            error_msg = f"Error initializing printer: {e}"
            if self.logger:
                self.logger.error(error_msg, exc_info=True)
            try:
                servicemanager.LogErrorMsg(f"VirtPrint initialization error: {e}")
            except:
                pass


if __name__ == '__main__':
    """
    Command-line interface for the service.
    
    Usage:
        python service.py install   - Install the service
        python service.py start     - Start the service
        python service.py stop      - Stop the service
        python service.py remove    - Remove the service
        python service.py debug     - Run in debug mode (console)
    """
    
    if len(sys.argv) == 1:
        # Run the service if no arguments provided
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(VirtPrintService)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        # Handle command-line arguments
        win32serviceutil.HandleCommandLine(VirtPrintService)
