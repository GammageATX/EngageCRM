"""Module for analyzing Outlook emails from a specified folder."""

import win32com.client
import re


def extract_unit_from_email(email_address, domain_mappings=None):
    """Extract unit name from email address or domain.
    
    Args:
        email_address (str): Email address to analyze
        domain_mappings (dict): Optional mapping of domains to unit names
    
    Returns:
        str: Best guess at unit name or None
    """
    if not email_address:
        return None
        
    # Default mappings if none provided
    domain_mappings = domain_mappings or {
        'army.mil': 'Army',
        'navy.mil': 'Navy',
        'af.mil': 'Air Force',
        'marines.mil': 'Marines',
        'uscg.mil': 'Coast Guard'
    }
    
    # Try to extract domain
    match = re.search(r'@(.+)$', email_address)
    if match:
        domain = match.group(1).lower()
        # Check domain mappings
        for key, unit in domain_mappings.items():
            if key in domain:
                return unit
    
    return None


def convert_email_to_engagement(email_info):
    """Convert email data to engagement format."""
    # Extract potential unit from email domains
    unit = None
    recipients = []
    
    # Process To and CC fields
    for email_field in [email_info['to'], email_info['cc']]:
        if email_field:
            # Split multiple recipients
            emails = [e.strip() for e in email_field.split(';')]
            for email in emails:
                # Try to extract unit
                if not unit:
                    unit = extract_unit_from_email(email)
                # Get display name if present
                if '<' in email:
                    name = email.split('<')[0].strip()
                    recipients.append(name)
                else:
                    recipients.append(email.split('@')[0])
    
    # Add sender to participants
    participants = [email_info['sender']] + recipients
    # Remove duplicates while preserving order
    participants = list(dict.fromkeys(participants))
    
    # Create attachment summary if any
    attachment_summary = ""
    if email_info['attachments']:
        attachment_summary = "\n\nAttachments:\n" + "\n".join(
            f"- {att}" for att in email_info['attachments']
        )
    
    # Create engagement data
    engagement_data = {
        'date_time': email_info['received'],
        'type': determine_engagement_type(email_info['subject']),
        'unit': unit,
        'summary': (
            f"Subject: {email_info['subject']}\n"
            f"From: {email_info['sender']}\n"
            f"To: {email_info['to']}\n"
            f"CC: {email_info['cc']}\n\n"
            f"{email_info['body'][:500]}..."
            f"{attachment_summary}"
        ),
        'status': 'Completed',
        'action_items': extract_action_items(email_info['body']),
        'participants': participants,
        'attachments': email_info['attachments']
    }
    
    return engagement_data


def determine_engagement_type(subject):
    """Determine engagement type from subject."""
    subject = subject.lower()
    
    type_keywords = {
        'meeting': 'Meeting',
        'conference': 'Meeting',
        'training': 'Training',
        'workshop': 'Training',
        'demo': 'Demo',
        'demonstration': 'Demo',
        'review': 'Review',
        'brief': 'Briefing',
        'briefing': 'Briefing'
    }
    
    for keyword, eng_type in type_keywords.items():
        if keyword in subject:
            return eng_type
    
    return 'Email Communication'


def extract_action_items(body):
    """Extract potential action items from email body."""
    action_markers = [
        'action item',
        'action required',
        'todo',
        'to-do',
        'to do',
        'next steps',
        'follow up',
        'followup',
        'please',
        'request'
    ]
    
    lines = body.lower().split('\n')
    action_items = []
    
    for i, line in enumerate(lines):
        if any(marker in line for marker in action_markers):
            # Get the full line and next line for context
            if i < len(lines) - 1:
                action_items.append(f"- {lines[i].strip()}\n  {lines[i + 1].strip()}")
            else:
                action_items.append(f"- {lines[i].strip()}")
    
    return "\n".join(action_items) if action_items else ""


def get_outlook_folder_emails(folder_name="Python Emails"):
    """Get emails from a specific Outlook folder.
    
    Args:
        folder_name (str): Name of the Outlook folder to analyze
        
    Returns:
        list: List of dictionaries containing email information
    """
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    
    # Get the inbox
    inbox = namespace.GetDefaultFolder(6)  # 6 is the index for inbox
    
    # Try to find our specific folder
    try:
        target_folder = inbox.Folders[folder_name]
    except Exception as e:  # Specify exception type
        print(f"Folder '{folder_name}' not found. Creating it... Error: {e}")
        target_folder = inbox.Folders.Add(folder_name)
    
    # Get messages
    messages = target_folder.Items
    
    # Sort by received time
    messages.Sort("[ReceivedTime]", True)
    
    email_data = []
    
    # Process each email
    for msg in messages:
        email_info = {
            'subject': msg.Subject,
            'sender': msg.SenderName,
            'received': msg.ReceivedTime,
            'body': msg.Body,
            'categories': msg.Categories,
            'attachments': [att.FileName for att in msg.Attachments],
            'cc': msg.CC,
            'to': msg.To,
            'importance': msg.Importance,
        }
        email_data.append(email_info)
        
    return email_data


def import_emails_as_engagements(folder_name="Python Emails"):
    """Import emails as engagements and return engagement data list."""
    emails = get_outlook_folder_emails(folder_name)
    engagements = []
    
    for email in emails:
        engagement = convert_email_to_engagement(email)
        engagements.append(engagement)
    
    return engagements


def analyze_emails():
    """Analyze emails and generate some basic metrics.
    
    Prints total email count, top senders, and attachment statistics.
    """
    emails = get_outlook_folder_emails()
    
    # Basic analytics
    total_emails = len(emails)
    senders = {}
    subjects = []
    attachments = 0
    
    for email in emails:
        # Count emails per sender
        sender = email['sender']
        senders[sender] = senders.get(sender, 0) + 1
        
        # Collect subjects
        subjects.append(email['subject'])
        
        # Count attachments
        attachments += len(email['attachments'])
    
    # Print analysis
    print(f"\nTotal Emails: {total_emails}")
    print("\nTop Senders:")
    for sender, count in sorted(senders.items(), key=lambda x: x[1], reverse=True)[:5]:
        print(f"  {sender}: {count}")
    print(f"\nTotal Attachments: {attachments}")


if __name__ == "__main__":
    analyze_emails()
