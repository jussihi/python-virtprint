import subprocess
import sys
import textwrap

import settings


def run_powershell(ps_script: str) -> None:
    cmd = [
        "powershell.exe",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-Command",
        ps_script,
    ]
    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode != 0:
        print("PowerShell failed.\n--- STDOUT ---\n", proc.stdout, sep="")
        print("--- STDERR ---\n", proc.stderr, sep="")
        raise SystemExit(proc.returncode)
    if proc.stdout.strip():
        print(proc.stdout.strip())

def add_tcp_printer(printer_name: str = settings.PRINTER_NAME,
                    driver_name: str = settings.DRIVER_NAME,
                    port_name: str = settings.TCP_HOST + "_" + str(settings.TCP_PORT),
                    ip_address: str = settings.TCP_HOST,
                    tcp_port: int = settings.TCP_PORT) -> None:
    ps = textwrap.dedent(
        f"""
        $ErrorActionPreference = "Stop"

        $printerName = "{printer_name}"
        $driverName  = "{driver_name}"
        $portName    = "{port_name}"
        $ipAddress   = "{ip_address}"
        $portNumber  = {tcp_port}
        # Validate driver exists on this machine
        $drv = Get-PrinterDriver -Name $driverName -ErrorAction SilentlyContinue
        if (-not $drv) {{
            throw "Printer driver '$driverName' not found. Ensure the driver is installed on this system."
        }}

        # Create (or reuse) the Standard TCP/IP port (RAW 9100) without SNMP
        $existingPort = Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue
        if (-not $existingPort) {{
            Add-PrinterPort -Name $portName -PrinterHostAddress $ipAddress -PortNumber $portNumber
            $port = Get-WmiObject Win32_TCPIPPrinterPort -Filter "Name='$portName'"
            $port.SNMPEnabled = $false
            $port.Put()
        }} else {{
            # If the port exists, try to ensure settings match (best-effort; some fields are immutable)
            try {{
                Set-PrinterPort -Name $portName -PrinterHostAddress $ipAddress -PortNumber $portNumber
                $port = Get-WmiObject Win32_TCPIPPrinterPort -Filter "Name='$portName'"
                $port.SNMPEnabled = $false
                $port.Put()
            }} catch {{
                # ignore if Set-PrinterPort isn't supported for some fields in this environment
            }}
        }}

        # Create the printer if missing
        $p = Get-Printer -Name $printerName -ErrorAction SilentlyContinue
        if (-not $p) {{
            Add-Printer -Name $printerName -DriverName $driverName -PortName $portName
        }} else {{
            # Ensure it uses the right driver/port
            Set-Printer -Name $printerName -DriverName $driverName -PortName $portName
        }}

        # Disable bidirectional support (often trips "faulty/unknown state" for virtual backends)
        try {{
            Set-Printer -Name $printerName -EnableBidi $false
        }} catch {{
            # Some builds don't expose -EnableBidi; ignore if not available
        }}

        # Helpful confirmation output
        Get-Printer -Name $printerName | Format-List Name,DriverName,PortName,Shared,Published
        Get-PrinterPort -Name $portName | Format-List Name,PrinterHostAddress,PortNumber,SNMPEnabled,Protocol
        """
    ).strip()

    run_powershell(ps)