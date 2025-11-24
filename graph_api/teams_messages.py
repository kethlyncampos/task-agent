"""
Module for fetching and formatting Microsoft Teams messages using Graph SDK.
"""

from typing import Dict, Any, Optional, List
from datetime import datetime, timedelta
import re
from msgraph.generated.users.item.chats.chats_request_builder import ChatsRequestBuilder
from kiota_abstractions.base_request_configuration import RequestConfiguration


async def get_recent_chats(graph, limit: int = 10):
    """
    Fetch recent chats for the current user using Microsoft Graph SDK.
    
    Parameters:
        graph: The Microsoft Graph client instance (from ctx.user_graph)
        limit (int): Maximum number of chats to retrieve (default: 10)
    
    Returns:
        dict: {
            'success': bool,
            'chats': list,  # List of chat objects
            'error': str or None
        }
    """
    try:
        # Create request configuration
        query_params = ChatsRequestBuilder.ChatsRequestBuilderGetQueryParameters()
        
        # Set query parameters
        query_params.top = limit
        query_params.orderby = ["lastMessagePreview/createdDateTime DESC"]
        query_params.expand = ["lastMessagePreview", "members"]
        
        request_config = RequestConfiguration(query_parameters=query_params)
        
        # Get user's chats
        chats_response = await graph.me.chats.get(request_configuration=request_config)
        
        # Extract chats from response
        chats = chats_response.value if chats_response and chats_response.value else []
        
        return {
            'success': True,
            'chats': chats,
            'error': None
        }
        
    except Exception as e:
        return {
            'success': False,
            'chats': [],
            'error': f"Error fetching chats: {str(e)}"
        }


async def get_chat_messages(graph, chat_id: str, limit: int = 50):
    """
    Fetch messages from a specific chat using Microsoft Graph SDK.
    
    Parameters:
        graph: The Microsoft Graph client instance
        chat_id (str): The ID of the chat
        limit (int): Maximum number of messages to retrieve (default: 50)
    
    Returns:
        dict: {
            'success': bool,
            'messages': list,  # List of message objects
            'error': str or None
        }
    """
    try:
        from msgraph.generated.chats.item.messages.messages_request_builder import MessagesRequestBuilder
        
        # Create request configuration
        query_params = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters()
        
        # Set query parameters
        query_params.top = limit
        query_params.orderby = ["createdDateTime DESC"]
        
        request_config = RequestConfiguration(query_parameters=query_params)
        
        # Get messages from the chat
        messages_response = await graph.chats.by_chat_id(chat_id).messages.get(
            request_configuration=request_config
        )
        
        # Extract messages from response
        messages = messages_response.value if messages_response and messages_response.value else []
        
        return {
            'success': True,
            'messages': messages,
            'error': None
        }
        
    except Exception as e:
        return {
            'success': False,
            'messages': [],
            'error': f"Error fetching messages: {str(e)}"
        }


async def get_all_recent_teams_messages(graph, num_chats: int = 5, messages_per_chat: int = 10):
    """
    Fetch recent messages across all user's chats.
    
    Parameters:
        graph: The Microsoft Graph client instance
        num_chats (int): Number of recent chats to check (default: 5)
        messages_per_chat (int): Max messages to fetch per chat (default: 10)
    
    Returns:
        dict: {
            'success': bool,
            'all_messages': list,  # List of all messages with chat context
            'error': str or None
        }
    """
    try:
        # First, get recent chats
        chats_result = await get_recent_chats(graph, limit=num_chats)
        
        if not chats_result['success']:
            return chats_result
        
        all_messages = []
        
        # For each chat, fetch recent messages
        for chat in chats_result['chats']:
            chat_id = chat.id
            chat_topic = chat.topic or "Unnamed Chat"
            
            # Get messages from this chat
            messages_result = await get_chat_messages(graph, chat_id, limit=messages_per_chat)
            
            if messages_result['success']:
                # Add chat context to each message
                for msg in messages_result['messages']:
                    all_messages.append({
                        'message': msg,
                        'chat_id': chat_id,
                        'chat_topic': chat_topic
                    })
        
        # Sort all messages by date (most recent first)
        all_messages.sort(
            key=lambda x: x['message'].created_date_time if x['message'].created_date_time else datetime.min,
            reverse=True
        )
        
        return {
            'success': True,
            'all_messages': all_messages,
            'error': None
        }
        
    except Exception as e:
        return {
            'success': False,
            'all_messages': [],
            'error': f"Error: {str(e)}"
        }


def format_teams_message(message_data: Dict[str, Any]) -> str:
    """
    Format a Teams message into a readable string.
    
    Args:
        message_data (dict): Dictionary containing 'message', 'chat_id', and 'chat_topic'
    
    Returns:
        str: Formatted message string
    """
    try:
        message = message_data.get('message')
        chat_topic = message_data.get('chat_topic', 'Unknown Chat')
        
        if not message:
            return "Error: No message data"
        
        lines = []
        lines.append(f"Chat: {chat_topic}")
        
        # Sender information
        if message.from_:
            if message.from_.user:
                sender_name = message.from_.user.display_name or "Unknown"
                lines.append(f"From: {sender_name}")
            elif message.from_.application:
                app_name = message.from_.application.display_name or "Bot/App"
                lines.append(f"From: {app_name} (Application)")
        else:
            lines.append("From: Unknown")
        
        # Message timestamp
        if message.created_date_time:
            timestamp = message.created_date_time.strftime("%Y-%m-%d %H:%M:%S")
            lines.append(f"Sent: {timestamp}")
        
        # Message type
        message_type = message.message_type.value if message.message_type else "message"
        lines.append(f"Type: {message_type}")
        
        # Importance
        if message.importance:
            importance = message.importance.value if hasattr(message.importance, 'value') else str(message.importance)
            if importance != "normal":
                lines.append(f"Importance: {importance}")
        
        # Message content
        if message.body:
            content = message.body.content or ""
            # Strip HTML tags
            content_clean = re.sub('<[^<]+?>', '', content).strip()
            if content_clean:
                lines.append("\nMessage:")
                lines.append(content_clean)
            else:
                lines.append("\nMessage: (no text content)")
        
        # Attachments
        if message.attachments and len(message.attachments) > 0:
            lines.append(f"\nAttachments: {len(message.attachments)}")
            for att in message.attachments[:3]:  # Show first 3
                att_name = att.name if hasattr(att, 'name') and att.name else "Unknown"
                lines.append(f"  - {att_name}")
        
        # Reactions
        if message.reactions and len(message.reactions) > 0:
            lines.append(f"\nReactions: {len(message.reactions)}")
        
        return "\n".join(lines)
        
    except Exception as e:
        return f"Error formatting message: {str(e)}"


def format_chat_summary(chat) -> str:
    """
    Format a chat object into a summary string.
    
    Args:
        chat: Chat object from Graph SDK
    
    Returns:
        str: Formatted chat summary
    """
    try:
        lines = []
        
        # Chat topic/title
        topic = chat.topic or "Unnamed Chat"
        lines.append(f"**{topic}**")
        
        # Chat type
        chat_type = chat.chat_type.value if chat.chat_type else "unknown"
        lines.append(f"  Type: {chat_type}")
        
        # Last message preview
        if chat.last_message_preview:
            preview = chat.last_message_preview
            if preview.created_date_time:
                timestamp = preview.created_date_time.strftime("%Y-%m-%d %H:%M")
                lines.append(f"  Last message: {timestamp}")
        
        # Members count
        if chat.members:
            lines.append(f"  Members: {len(chat.members)}")
        
        return "\n".join(lines)
        
    except Exception as e:
        return f"Error formatting chat: {str(e)}"

