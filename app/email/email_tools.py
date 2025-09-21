import smtplib #SMTP -- Simple Mail Transfer Protocol
import os
import mimetypes
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path
from typing import Optional, List, Dict
import logging

logger = logging.getLogger(__name__)

def send_gmail_with_attachment(
    to_email: str,
    subject: str,
    body: str,
    attachment_path: Optional[str] = None,
    sender_email: Optional[str] = None,
    sender_password: Optional[str] = None,
    cc_emails: Optional[List[str]] = None,
    bcc_emails: Optional[List[str]] = None
) -> Dict[str, any]:
    """
    Send an email with optional attachment via Gmail SMTP.
    
    Args:
        to_email (str): Recipient email address
        subject (str): Email subject
        body (str): Email body (supports HTML)
        attachment_path (str, optional): Path to file to attach
        sender_email (str, optional): Sender email (defaults to env GMAIL_EMAIL)
        sender_password (str, optional): App password (defaults to env GMAIL_APP_PASSWORD)
        cc_emails (List[str], optional): CC recipients
        bcc_emails (List[str], optional): BCC recipients
        
    Returns:
        Dict: Status and message about the email sending operation
    """
    try:
        # Get credentials from environment or parameters
        sender_email = sender_email or os.getenv('GMAIL_EMAIL')
        sender_password = sender_password or os.getenv('GMAIL_APP_PASSWORD')
        
        if not sender_email or not sender_password:
            return {
                "success": False,
                "error": "Gmail credentials not found. Please set GMAIL_EMAIL and GMAIL_APP_PASSWORD environment variables."
            }
        
        # Create message
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = to_email
        msg['Subject'] = subject
        
        # Add CC and BCC if provided
        if cc_emails:
            msg['Cc'] = ', '.join(cc_emails)
        if bcc_emails:
            msg['Bcc'] = ', '.join(bcc_emails)
        
        # Add body
        msg.attach(MIMEText(body, 'html' if '<' in body else 'plain'))
        
        # Add attachment if provided
        attachment_name = None
        if attachment_path and Path(attachment_path).exists():
            attachment_name = Path(attachment_path).name
            
            # Guess the content type based on the file's extension
            ctype, encoding = mimetypes.guess_type(attachment_path)
            if ctype is None or encoding is not None:
                ctype = 'application/octet-stream'
            
            maintype, subtype = ctype.split('/', 1)
            
            with open(attachment_path, 'rb') as fp:
                attachment = MIMEBase(maintype, subtype)
                attachment.set_payload(fp.read())
                encoders.encode_base64(attachment)
                attachment.add_header(
                    'Content-Disposition',
                    f'attachment; filename={attachment_name}'
                )
                msg.attach(attachment)
        
        # Prepare recipient list
        recipients = [to_email]
        if cc_emails:
            recipients.extend(cc_emails)
        if bcc_emails:
            recipients.extend(bcc_emails)
        
        # Send email
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, sender_password)
        text = msg.as_string()
        server.sendmail(sender_email, recipients, text)
        server.quit()
        
        logger.info(f"Email sent successfully to {to_email}")
        return {
            "success": True,
            "message": f"Email sent successfully to {to_email}",
            "attachment": attachment_name if attachment_name else None,
            "recipients": len(recipients)
        }
        
    except Exception as e:
        error_msg = f"Failed to send email: {str(e)}"
        logger.error(error_msg)
        return {
            "success": False,
            "error": error_msg
        }

def create_ppt_email_body(topic: str, ppt_path: str, generation_details: Dict) -> str:
    """
    Create a professional email body for sending PowerPoint presentations.
    
    Args:
        topic (str): The presentation topic
        ppt_path (str): Path to the PowerPoint file
        generation_details (Dict): Details about the generation process
        
    Returns:
        str: HTML formatted email body
    """
    ppt_filename = Path(ppt_path).name if ppt_path else "presentation.pptx"
    
    html_body = f"""
    <html>
    <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
        <h2 style="color: #2c3e50;">PowerPoint Presentation: {topic}</h2>
        
        <p>Hello,</p>
        
        <p>I'm pleased to share the PowerPoint presentation on <strong>"{topic}"</strong> that was automatically generated using our AI-powered presentation system.</p>
        
        <h3 style="color: #34495e;">Presentation Details:</h3>
        <ul>
            <li><strong>Topic:</strong> {topic}</li>
            <li><strong>File:</strong> {ppt_filename}</li>
            <li><strong>Generated:</strong> {generation_details.get('timestamp', 'Recently')}</li>
            <li><strong>Status:</strong> {'✅ Successfully generated' if generation_details.get('success') else '❌ Generation had issues'}</li>
        </ul>
        
        <p>The presentation includes:</p>
        <ul>
            <li>Research-backed content</li>
            <li>Professional slide layouts</li>
            <li>Relevant images and graphics</li>
            <li>Well-structured information flow</li>
        </ul>
        
        <p>Please find the PowerPoint presentation attached to this email.</p>
        
        <p>If you have any questions or need modifications to the presentation, please don't hesitate to reach out.</p>
        
        <p>Best regards,<br>
        <strong>AI Presentation Generator</strong></p>
        
        <hr style="border: 1px solid #eee; margin: 20px 0;">
        <p style="font-size: 12px; color: #666;">
            This presentation was automatically generated using AI technology. Please review the content for accuracy and appropriateness before use.
        </p>
    </body>
    </html>
    """
    
    return html_body

# Tool wrapper for LangChain integration
def send_ppt_email_tool(
    to_email: str,
    subject: str,
    ppt_path: str,
    topic: str,
    generation_details: Dict = None,
    custom_message: str = None
) -> str:
    """
    LangChain tool for sending PowerPoint presentations via email.
    
    Args:
        to_email (str): Recipient email address
        subject (str): Email subject
        ppt_path (str): Path to PowerPoint file
        topic (str): Presentation topic
        generation_details (Dict): Details about generation process
        custom_message (str): Custom message to include in email
        
    Returns:
        str: Result message
    """
    try:
        # Generate email body
        if custom_message:
            body = f"<p>{custom_message}</p><br>" + create_ppt_email_body(topic, ppt_path, generation_details or {})
        else:
            body = create_ppt_email_body(topic, ppt_path, generation_details or {})
        
        # Send email
        result = send_gmail_with_attachment(
            to_email=to_email,
            subject=subject,
            body=body,
            attachment_path=ppt_path
        )
        
        if result["success"]:
            return f"✅ Email sent successfully to {to_email}. Attachment: {result.get('attachment', 'None')}"
        else:
            return f"❌ Failed to send email: {result['error']}"
            
    except Exception as e:
        return f"❌ Error sending email: {str(e)}"

# Create LangChain tool
from langchain_core.tools import tool

@tool
def send_presentation_email(
    to_email: str,
    subject: str,
    ppt_path: str,
    topic: str,
    custom_message: str = None
) -> str:
    """
    Send a PowerPoint presentation via email.
    
    Args:
        to_email: Recipient email address
        subject: Email subject line
        ppt_path: Full path to the PowerPoint file to attach
        topic: The topic/title of the presentation
        custom_message: Optional custom message to include in email
        
    Returns:
        Status message about the email sending operation
    """
    return send_ppt_email_tool(
        to_email=to_email,
        subject=subject,
        ppt_path=ppt_path,
        topic=topic,
        custom_message=custom_message
    )

# Export the tool for use in agents
email_tools = [send_presentation_email]
