
import requests
from datetime import datetime, timedelta
import re
from typing import Dict, Any, Optional, List
from datetime import timezone
from msgraph.generated.users.item.messages.messages_request_builder import MessagesRequestBuilder
from kiota_abstractions.base_request_configuration import RequestConfiguration

async def get_mail_inbox(graph, start_date=None, end_date=None, target_user_email=None):
    """
    Fetch emails using Microsoft Graph SDK, optionally filtering by date range.
    
    Parameters:
        graph: The Microsoft Graph client instance (from ctx.user_graph)
        start_date (str, optional): Start datetime in ISO 8601 format (e.g., '2024-04-01T00:00:00Z')
        end_date (str, optional): End datetime in ISO 8601 format (e.g., '2024-04-30T23:59:59Z')
        target_user_email (str, optional): Email of user to query. If None, uses current user (me)
    
    Returns:
        dict: {
            'success': bool,
            'emails': list,
            'error': str or None
        }
    """

    
    try:
        # Build the filter query
        filter_parts = []
        
        if start_date:
            # Ensure ISO 8601 format with Z suffix
            if not start_date.endswith('Z'):
                start_date = start_date + 'Z' if 'T' in start_date else start_date + 'T00:00:00Z'
            filter_parts.append(f"receivedDateTime ge {start_date}")
        
        if end_date:
            # Ensure ISO 8601 format with Z suffix
            if not end_date.endswith('Z'):
                end_date = end_date + 'Z' if 'T' in end_date else end_date + 'T23:59:59Z'
            filter_parts.append(f"receivedDateTime le {end_date}")
        
        # Create request configuration
        query_params = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters()
        
        # Apply filter if we have date constraints
        if filter_parts:
            query_params.filter = " and ".join(filter_parts)
        
        # Optional: add other query parameters
        # query_params.top = 50  # Limit results
        # query_params.orderby = ["receivedDateTime DESC"]  # Sort by date
        # query_params.select = ["subject", "from", "receivedDateTime", "bodyPreview"]  # Select specific fields
        
        request_config = RequestConfiguration(query_parameters=query_params)
        
        # Get messages
        if target_user_email:
            # Query specific user's mailbox
            messages_response = await graph.users.by_user_id(target_user_email).messages.get(
                request_configuration=request_config
            )
        else:
            # Query current user's mailbox
            messages_response = await graph.me.messages.get(
                request_configuration=request_config
            )
        
        # Extract messages from response
        emails = messages_response.value if messages_response and messages_response.value else []
        
        return {
            'success': True,
            'emails': emails,
            'error': None
        }
        
    except Exception as e:
        return {
            'success': False,
            'emails': [],
            'error': f"Error: {str(e)}"
        }

#def get_mail_inbox_1_day():
#    # Use UTC timezone for consistency
#    start_date = (datetime.now(timezone.utc) - timedelta(days=10)).isoformat().replace('+00:00', 'Z')
#    end_date = datetime.now(timezone.utc).isoformat().replace('+00:00', 'Z')
#    result = get_mail_inbox(start_date=start_date, end_date=end_date)
#    if result['success']:
#        return {
#            'success': True,
#            'emails': result['emails'],
#            'error': None
#        }
#    else:
#        return {
#            'success': False,
#            'emails': [],
#            'error': result['error']
#        }



def format_email(email: Dict[str, Any]) -> str:
    """
    Formats the email dictionary into a readable string.

    Args:
        email (Dict[str, Any]): The email information.

    Returns:
        str: Formatted email string.
    """
    try:
        lines = []
        lines.append(f"Reading email: {email.get('subject', '')}")
        lines.append(f"Subject: {email.get('subject', '')}")
        lines.append(f"From: {email.get('sender_name', '')} <{email.get('sender', '')}>")
        lines.append(f"To: {', '.join(email.get('to', []))}")
        if email.get('cc'):
            lines.append(f"CC: {', '.join(email.get('cc', []))}")
        lines.append(f"Received: {email.get('received_datetime', '')}")
        lines.append(f"Importance: {email.get('importance', '')}")
        lines.append(f"Read: {'Yes' if email.get('is_read', False) else 'No'}")
        lines.append(f"Has Attachments: {'Yes' if email.get('has_attachments', False) else 'No'}")
        lines.append("Email Body:")
        body = email.get('body', {}).get('content', '')
        body_clean = re.sub('<[^<]+?>', '', body)
        lines.append(body_clean)
        if email.get('attachments'):
            lines.append(f"\nAttachments ({len(email['attachments'])}):")
            for attachment in email['attachments']:
                lines.append(f"   - {attachment.get('name', 'Unknown')}")
        email_content = "\n".join(lines)
        return email_content
    except Exception as e:
        return f"Error formatting email: {str(e)}"






def format_inbox_summary(emails: List[Dict[str, Any]], max_length: int = 50) -> str:
    """
    Format a list of emails into a readable summary string.
    
    Args:
        emails: List of email dictionaries from get_delegated_inbox
        max_length: Maximum length of subject before truncating
    
    Returns:
        Formatted string with email summaries
    """
    if not emails:
        return "📧 Your inbox is empty!"
    
    lines = ["📧 **Your Recent Emails:**\n"]
    
    for i, email in enumerate(emails[:10], 1):
        subject = email.get('subject', '(No subject)')
        if len(subject) > max_length:
            subject = subject[:max_length] + "..."
        
        from_name = email.get('from_name', email.get('from', 'Unknown'))
        read_icon = "✓" if email.get('is_read') else "●"
        attach_icon = "📎" if email.get('has_attachments') else ""
        
        lines.append(f"{i}. {read_icon} **{subject}** - {from_name} {attach_icon}")
    
    return "\n".join(lines)


