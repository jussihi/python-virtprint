"""
User-Definable Callbacks for virtprint
Customize this file to define what happens after a print job is processed.
"""

import logging

logger = logging.getLogger(__name__)


def on_print_job_complete(output_files, job_info):
    """
    Called when a print job has been intercepted and processed.
    
    Args:
        output_files (list or None): 
            - List of Path objects for successfully created files (ordered by page number)
            - For PDF: Single-element list with the PDF path
            - For images (PNG/JPEG/TIFF): List of image paths in page order
              e.g., ['output_page1.png', 'output_page2.png', 'output_page3.png']
            - None if the job failed to process
            
        job_info (dict): Information about the print job
            - 'job_id': Print job ID
            - 'document_name': Name of the document being printed
            - 'user_name': User who initiated the print job
            - 'machine_name': Computer name
            - 'pages': Number of pages (if available)
            - 'output_format': Format used (PDF, PNG, JPEG, TIFF)
    
    Example usage:
        - Send files to cloud storage
        - Email the files
        - Move files to a specific folder
        - Trigger other automation
        - Log to a database
    """
    
    # Default implementation - just log the result
    if output_files is None:
        logger.warning(f"Print job failed: {job_info.get('document_name', 'Unknown')}")
        print(f"❌ Print job FAILED: {job_info.get('document_name', 'Unknown')}")
    else:
        logger.info(f"Print job completed: {len(output_files)} file(s) created")
        print(f"\n✅ Print job completed successfully!")
        print(f"   Document: {job_info.get('document_name', 'Unknown')}")
        print(f"   User: {job_info.get('user_name', 'Unknown')}")
        print(f"   Format: {job_info.get('output_format', 'Unknown')}")
        print(f"   Files created ({len(output_files)}):")
        for i, file_path in enumerate(output_files, 1):
            print(f"      Page {i}: {file_path}")
        print()


# Example custom implementations (uncomment and modify as needed):

"""
def on_print_job_complete(output_files, job_info):
    # Example 1: Copy files to a network share
    import shutil
    from pathlib import Path
    
    if output_files:
        network_path = Path(r"\\server\share\scanned_documents")
        for file_path in output_files:
            dest = network_path / file_path.name
            shutil.copy2(file_path, dest)
            print(f"Copied {file_path.name} to network share")
"""

"""
def on_print_job_complete(output_files, job_info):
    # Example 2: Send email with attachments
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.base import MIMEBase
    from email.mime.text import MIMEText
    from email import encoders
    
    if output_files:
        msg = MIMEMultipart()
        msg['From'] = "printer@example.com"
        msg['To'] = "recipient@example.com"
        msg['Subject'] = f"Print: {job_info.get('document_name', 'Document')}"
        
        body = f"Print job from {job_info.get('user_name', 'Unknown')}"
        msg.attach(MIMEText(body, 'plain'))
        
        for file_path in output_files:
            with open(file_path, 'rb') as f:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename={file_path.name}')
                msg.attach(part)
        
        # Send email (configure your SMTP server)
        # server = smtplib.SMTP('smtp.example.com', 587)
        # server.starttls()
        # server.login("user", "password")
        # server.send_message(msg)
        # server.quit()
"""

"""
def on_print_job_complete(output_files, job_info):
    # Example 3: Upload to cloud storage (e.g., Dropbox, Google Drive)
    if output_files:
        for file_path in output_files:
            # Your cloud upload code here
            print(f"Would upload {file_path} to cloud storage")
"""

"""
def on_print_job_complete(output_files, job_info):
    # Example 4: Run OCR and save text
    import pytesseract
    from PIL import Image
    
    if output_files:
        for file_path in output_files:
            if file_path.suffix.lower() in ['.png', '.jpg', '.jpeg', '.tiff']:
                # Run OCR
                text = pytesseract.image_to_string(Image.open(file_path))
                
                # Save OCR text
                text_file = file_path.with_suffix('.txt')
                with open(text_file, 'w', encoding='utf-8') as f:
                    f.write(text)
                
                print(f"OCR completed: {text_file}")
"""

"""
def on_print_job_complete(output_files, job_info):
    # Example 5: Move files to dated folder structure
    from pathlib import Path
    from datetime import datetime
    import shutil
    
    if output_files:
        # Create dated folder: YYYY/MM/DD
        base_path = Path(r"C:\Documents\Printed")
        today = datetime.now()
        dest_folder = base_path / str(today.year) / f"{today.month:02d}" / f"{today.day:02d}"
        dest_folder.mkdir(parents=True, exist_ok=True)
        
        for file_path in output_files:
            dest = dest_folder / file_path.name
            shutil.move(str(file_path), str(dest))
            print(f"Moved to: {dest}")
"""
