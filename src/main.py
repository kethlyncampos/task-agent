"""
Copyright (c) Microsoft Corporation. All rights reserved.
Licensed under the MIT License.
"""

import asyncio
import logging
import os
import re
from datetime import datetime

from microsoft.teams.api import MessageActivity
from microsoft.teams.apps import ActivityContext, App, SignInEvent

from graph_api.mail_inbox import get_mail_inbox, format_email
from graph_api.teams_messages import get_all_recent_teams_messages, format_teams_message, get_recent_chats, format_chat_summary
from graph_api.todo_tasks import (
    get_todo_lists, get_tasks_from_list, create_task, complete_task,
    update_task, delete_task, create_task_list, format_task, format_task_list
)
from todo_list_generation.todo_list_generation import run_generate_todo_list, generate_single_task_from_user_message

logger = logging.getLogger(__name__)

# Create app with OAuth connection
app = App(default_connection_name=os.getenv("CONNECTION_NAME", "graph"))


def validate_and_format_date(date_str: str) -> str:
    """
    Validate and format date string to ISO 8601 format (YYYY-MM-DD).
    
    Args:
        date_str: Date string in various formats
        
    Returns:
        str: Date in YYYY-MM-DD format, or None if invalid
    """
    if not date_str or date_str == "null" or date_str.lower() == "none":
        return None
    
    # Common date formats to try
    date_formats = [
        "%Y-%m-%d",           # 2025-11-24
        "%d/%m/%Y",           # 24/11/2025
        "%d-%m-%Y",           # 24-11-2025
        "%Y/%m/%d",           # 2025/11/24
        "%d.%m.%Y",           # 24.11.2025
        "%Y-%m-%dT%H:%M:%S",  # 2025-11-24T10:00:00
        "%Y-%m-%d %H:%M:%S",  # 2025-11-24 10:00:00
    ]
    
    for fmt in date_formats:
        try:
            parsed_date = datetime.strptime(date_str.strip(), fmt)
            # Return in ISO 8601 format (YYYY-MM-DD)
            return parsed_date.strftime("%Y-%m-%d")
        except ValueError:
            continue
    
    # If no format matched, log and return None
    logger.warning(f"Could not parse date: {date_str}")
    return None


@app.on_message_pattern("signout")
async def handle_signout_command(ctx: ActivityContext[MessageActivity]):
    """Handle sign-out command."""
    if not ctx.is_signed_in:
        await ctx.send("ℹ️ You are not currently signed in.")
    else:
        await ctx.sign_out()
        await ctx.send("👋 You have been signed out successfully!")


@app.on_message_pattern("profile")
async def handle_profile_command(ctx: ActivityContext[MessageActivity]):
    """Handle profile command using Graph API with TokenProtocol pattern."""

    if not ctx.is_signed_in:
        await ctx.send("🔐 Please sign in first to access Microsoft Graph.")
        await ctx.sign_in()
        return

    graph = ctx.user_graph
    # Fetch user profile
    if graph:
        me = await graph.me.get()

    if me:
        profile_info = (
            f"👤 **Your Profile**\n\n"
            f"**Name:** {me.display_name or 'N/A'}\n\n"
            f"**Email:** {me.user_principal_name or 'N/A'}\n\n"
            f"**Job Title:** {me.job_title or 'N/A'}\n\n"
            f"**Department:** {me.department or 'N/A'}\n\n"
            f"**Office:** {me.office_location or 'N/A'}"
        )
        await ctx.send(profile_info)
    else:
        await ctx.send("❌ Could not retrieve your profile information.")



@app.on_message_pattern("emails")
async def handle_emails_command(ctx: ActivityContext[MessageActivity]):
    """Handle emails command to fetch and display user's emails."""
    
    if not ctx.is_signed_in:
        await ctx.send("🔐 Please sign in first to access Microsoft Graph.")
        await ctx.sign_in()
        return
    
    graph = ctx.user_graph
    
    if not graph:
        await ctx.send("❌ Could not access Microsoft Graph.")
        return
    
    # Example: Get emails from the last 7 days
    from datetime import datetime, timedelta
    end_date = datetime.utcnow().isoformat() + 'Z'
    start_date = (datetime.utcnow() - timedelta(days=7)).isoformat() + 'Z'
    
    # Get current user's emails
    result = await get_mail_inbox(graph, start_date=start_date, end_date=end_date)
    
    # Or get specific user's emails (requires admin permissions)
    # result = await get_mail_inbox(
    #     graph, 
    #     start_date=start_date, 
    #     end_date=end_date,
    #     target_user_email="kethlyncampos@synecticai.onmicrosoft.com"
    # )
    
    if result['success']:
        emails = result['emails']
        if emails:
            # Process and display first 2 emails with full formatting
            num_emails_to_show = min(2, len(emails))
            
            await ctx.send(f"📧 **Found {len(emails)} emails. Showing first {num_emails_to_show}:**\n")
            
            for i, email_obj in enumerate(emails[:num_emails_to_show], 1):
                # Convert SDK email object to dictionary format expected by format_email
                email_dict = {
                    'subject': email_obj.subject or '(No subject)',
                    'sender_name': email_obj.from_.email_address.name if email_obj.from_ and email_obj.from_.email_address else 'Unknown',
                    'sender': email_obj.from_.email_address.address if email_obj.from_ and email_obj.from_.email_address else 'Unknown',
                    'to': [recipient.email_address.address for recipient in (email_obj.to_recipients or []) if recipient.email_address],
                    'cc': [recipient.email_address.address for recipient in (email_obj.cc_recipients or []) if recipient.email_address],
                    'received_datetime': email_obj.received_date_time.strftime("%Y-%m-%d %H:%M:%S") if email_obj.received_date_time else 'Unknown',
                    'importance': email_obj.importance.value if email_obj.importance else 'normal',
                    'is_read': email_obj.is_read if email_obj.is_read is not None else False,
                    'has_attachments': email_obj.has_attachments if email_obj.has_attachments is not None else False,
                    'body': {
                        'content': email_obj.body.content if email_obj.body and email_obj.body.content else ''
                    },
                    'attachments': []
                }
                
                # Format and send the email
                formatted_email = format_email(email_dict)
                await ctx.send(f"\n{'='*60}\n**Email {i}**\n{'='*60}\n{formatted_email}\n")
            
            if len(emails) > 2:
                await ctx.send(f"\n📬 ... and {len(emails) - 2} more email(s) not shown.")
        else:
            await ctx.send("📭 No emails found in the specified date range.")
    else:
        await ctx.send(f"❌ Error fetching emails: {result['error']}")


@app.on_message_pattern("teams")
async def handle_teams_messages_command(ctx: ActivityContext[MessageActivity]):
    """Handle teams messages command to fetch and display recent Teams messages."""
    
    if not ctx.is_signed_in:
        await ctx.send("🔐 Please sign in first to access Microsoft Graph.")
        await ctx.sign_in()
        return
    
    graph = ctx.user_graph
    
    if not graph:
        await ctx.send("❌ Could not access Microsoft Graph.")
        return
    
    await ctx.send("🔍 Fetching your recent Teams messages...")
    
    # Fetch recent messages from the last 3 chats, 5 messages per chat
    result = await get_all_recent_teams_messages(graph, num_chats=3, messages_per_chat=5)
    
    if result['success']:
        messages = result['all_messages']
        
        if messages:
            num_to_show = min(2, len(messages))
            await ctx.send(f"💬 **Found {len(messages)} recent messages. Showing first {num_to_show}:**\n")
            
            for i, msg_data in enumerate(messages[:num_to_show], 1):
                formatted_msg = format_teams_message(msg_data)
                await ctx.send(f"\n{'='*60}\n**Message {i}**\n{'='*60}\n{formatted_msg}\n")
            
            if len(messages) > 2:
                await ctx.send(f"\n💬 ... and {len(messages) - 2} more message(s) not shown.")
        else:
            await ctx.send("💬 No recent Teams messages found.")
    else:
        await ctx.send(f"❌ Error fetching Teams messages: {result['error']}")


@app.on_message_pattern("chats")
async def handle_chats_command(ctx: ActivityContext[MessageActivity]):
    """Handle chats command to list user's recent chats."""
    
    if not ctx.is_signed_in:
        await ctx.send("🔐 Please sign in first to access Microsoft Graph.")
        await ctx.sign_in()
        return
    
    graph = ctx.user_graph
    
    if not graph:
        await ctx.send("❌ Could not access Microsoft Graph.")
        return
    
    await ctx.send("🔍 Fetching your recent chats...")
    
    # Fetch recent chats
    result = await get_recent_chats(graph, limit=5)
    
    if result['success']:
        chats = result['chats']
        
        if chats:
            message = f"💬 **Your {len(chats)} most recent chats:**\n\n"
            
            for i, chat in enumerate(chats, 1):
                chat_summary = format_chat_summary(chat)
                message += f"{i}. {chat_summary}\n\n"
            
            await ctx.send(message)
        else:
            await ctx.send("💬 No chats found.")
    else:
        await ctx.send(f"❌ Error fetching chats: {result['error']}")


@app.on_message_pattern("todo lists")
async def handle_todo_lists_command(ctx: ActivityContext[MessageActivity]):
    """Handle todo lists command to display all To Do task lists."""
    
    if not ctx.is_signed_in:
        await ctx.send("🔐 Please sign in first to access Microsoft Graph.")
        await ctx.sign_in()
        return
    
    graph = ctx.user_graph
    
    if not graph:
        await ctx.send("❌ Could not access Microsoft Graph.")
        return
    
    await ctx.send("🔍 Fetching your To Do lists...")
    
    # Fetch all task lists
    result = await get_todo_lists(graph)
    
    if result['success']:
        lists = result['lists']
        
        if lists:
            message = f"📋 **Your To Do Lists ({len(lists)}):**\n\n"
            
            for i, task_list in enumerate(lists, 1):
                list_summary = format_task_list(task_list)
                list_id_short = task_list.id[:8] if task_list.id else "unknown"
                message += f"{i}. {list_summary}\n   ID: {list_id_short}...\n\n"
            
            message += "\n💡 Use **'todo tasks [list-name]'** to view tasks from a specific list"
            await ctx.send(message)
        else:
            await ctx.send("📋 No To Do lists found.")
    else:
        await ctx.send(f"❌ Error fetching To Do lists: {result['error']}")


@app.on_message_pattern("todo tasks")
async def handle_todo_tasks_command(ctx: ActivityContext[MessageActivity]):
    """Handle todo tasks command to display tasks from a list."""
    
    if not ctx.is_signed_in:
        await ctx.send("🔐 Please sign in first to access Microsoft Graph.")
        await ctx.sign_in()
        return
    
    graph = ctx.user_graph
    
    if not graph:
        await ctx.send("❌ Could not access Microsoft Graph.")
        return
    
    await ctx.send("🔍 Fetching your tasks...")
    
    # First, get all lists to find the default one
    lists_result = await get_todo_lists(graph)
    
    if not lists_result['success']:
        await ctx.send(f"❌ Error fetching lists: {lists_result['error']}")
        return
    
    lists = lists_result['lists']
    
    if not lists:
        await ctx.send("📋 No To Do lists found. Create one first!")
        return
    
    # Use the first list (usually the default "Tasks" list)
    default_list = lists[0]
    list_name = default_list.display_name
    list_id = default_list.id
    
    # Fetch tasks from the list
    tasks_result = await get_tasks_from_list(graph, list_id)
    
    if tasks_result['success']:
        tasks = tasks_result['tasks']
        
        if tasks:
            # Separate completed and incomplete tasks
            incomplete_tasks = [t for t in tasks if t.status != "completed"]
            completed_tasks = [t for t in tasks if t.status == "completed"]
            
            message = f"📋 **Tasks in '{list_name}' ({len(tasks)} total)**\n\n"
            
            if incomplete_tasks:
                message += f"**Pending ({len(incomplete_tasks)}):**\n"
                for task in incomplete_tasks[:5]:
                    formatted = format_task(task)
                    message += f"{formatted}\n\n"
                
                if len(incomplete_tasks) > 5:
                    message += f"... and {len(incomplete_tasks) - 5} more pending task(s)\n\n"
            
            if completed_tasks:
                message += f"\n**Completed ({len(completed_tasks)}):**\n"
                for task in completed_tasks[:3]:
                    formatted = format_task(task)
                    message += f"{formatted}\n\n"
                
                if len(completed_tasks) > 3:
                    message += f"... and {len(completed_tasks) - 3} more completed task(s)\n"
            
            await ctx.send(message)
        else:
            await ctx.send(f"📋 No tasks found in '{list_name}'.\n\n💡 Use **'todo create'** to add a task!")
    else:
        await ctx.send(f"❌ Error fetching tasks: {tasks_result['error']}")


@app.on_message_pattern(re.compile(r"^add\s+task\s*$", re.IGNORECASE))
async def handle_add_task_help(ctx: ActivityContext[MessageActivity]):
    """Handle 'add task' without text - show help."""
    await ctx.send(
        "❌ Por favor, forneça uma descrição da tarefa.\n\n"
        "**Uso:** add task {descrição da sua tarefa}\n\n"
        "**Exemplos:**\n"
        "• add task Revisar relatório do Q4 até sexta-feira\n"
        "• add task Ligar para cliente sobre proposta\n"
        "• add task Preparar apresentação para reunião de segunda"
    )


@app.on_message_pattern(re.compile(r"^add\s+task\s+(.+)", re.IGNORECASE))
async def handle_add_task_command(ctx: ActivityContext[MessageActivity]):
    """Handle 'add task {text}' command to create a detailed task from user message using AI."""
    
    if not ctx.is_signed_in:
        await ctx.send("🔐 Please sign in first to access Microsoft Graph.")
        await ctx.sign_in()
        return
    
    graph = ctx.user_graph
    
    if not graph:
        await ctx.send("❌ Could not access Microsoft Graph.")
        return
    
    # Extract task text from the message using regex
    message_text = ctx.activity.text.strip()
    match = re.search(r"^add\s+task\s+(.+)", message_text, re.IGNORECASE)
    
    if match:
        task_text = match.group(1).strip()
    else:
        # This shouldn't happen due to the regex pattern, but handle it just in case
        await ctx.send(
            "❌ Por favor, forneça uma descrição da tarefa.\n\n"
            "**Uso:** add task {descrição da sua tarefa}\n\n"
            "**Exemplo:** add task Revisar relatório do Q4 até sexta-feira"
        )
        return
    
    # Get the first (default) task list
    lists_result = await get_todo_lists(graph)
    
    if not lists_result['success'] or not lists_result['lists']:
        await ctx.send("❌ No To Do lists found. Please create a list first.")
        return
    
    default_list = lists_result['lists'][0]
    list_id = default_list.id
    list_name = default_list.display_name
    
    await ctx.send("🤖 Gerando tarefa detalhada com IA...")
    
    try:
        # Generate detailed task using AI
        task_entry = await generate_single_task_from_user_message(task_text)
        
        if not task_entry or not task_entry.get('task'):
            await ctx.send("❌ Could not generate task. Please try again with a clearer message.")
            return
        
        # Extract task details
        task_title = task_entry.get('task', task_text)
        priority = task_entry.get('priority', 'normal')
        comments = task_entry.get('comments', '')
        due_date_raw = task_entry.get('due_date', None)
        
        # Validate and format the due date
        due_date = None
        if due_date_raw:
            due_date = validate_and_format_date(str(due_date_raw))
            if not due_date:
                logger.warning(f"Invalid due_date from AI: {due_date_raw}, ignoring it")
        
        # Map priority to importance
        importance = "normal"
        if priority and 'high' in priority.lower():
            importance = "high"
        elif priority and 'low' in priority.lower():
            importance = "low"
        
        # Create the task with AI-generated details
        result = await create_task(
            graph,
            list_id=list_id,
            title=task_title,
            body=comments if comments else None,
            due_date=due_date,
            importance=importance
        )
        
        if result['success']:
            task = result['task']
            await ctx.send(
                f"✅ **Tarefa Criada com Sucesso!**\n\n"
                f"Lista: {list_name}\n"
                f"{format_task(task)}\n\n"
                f"💡 Use **'todo tasks'** para ver todas as suas tarefas"
            )
        else:
            await ctx.send(f"❌ Erro ao criar tarefa: {result['error']}")
            
    except Exception as e:
        logger.error(f"Error in add task command: {str(e)}", exc_info=True)
        await ctx.send(
            f"❌ **Erro ao gerar tarefa:** {str(e)}\n\n"
            "Criando tarefa simples sem IA..."
        )
        
        # Fallback: create simple task without AI
        result = await create_task(
            graph,
            list_id=list_id,
            title=task_text,
            body=None,
            due_date=None,
            importance="normal"
        )
        
        if result['success']:
            task = result['task']
            await ctx.send(
                f"✅ **Tarefa Criada!**\n\n"
                f"Lista: {list_name}\n"
                f"{format_task(task)}"
            )
        else:
            await ctx.send(f"❌ Erro ao criar tarefa: {result['error']}")


@app.on_message_pattern("todo create")
async def handle_todo_create_command(ctx: ActivityContext[MessageActivity]):
    """Handle todo create command to create a new task."""
    
    if not ctx.is_signed_in:
        await ctx.send("🔐 Please sign in first to access Microsoft Graph.")
        await ctx.sign_in()
        return
    
    graph = ctx.user_graph
    
    if not graph:
        await ctx.send("❌ Could not access Microsoft Graph.")
        return
    
    # Get the first (default) task list
    lists_result = await get_todo_lists(graph)
    
    if not lists_result['success'] or not lists_result['lists']:
        await ctx.send("❌ No To Do lists found. Please create a list first.")
        return
    
    default_list = lists_result['lists'][0]
    list_id = default_list.id
    list_name = default_list.display_name
    
    # Example: Create a demo task
    from datetime import datetime, timedelta
    due_date = (datetime.utcnow() + timedelta(days=7)).strftime("%Y-%m-%d")
    
    result = await create_task(
        graph,
        list_id=list_id,
        title="Demo Task - Review Project Documentation",
        body="This is a sample task created via Microsoft Graph API. Please review and update all project documentation.",
        due_date=due_date,
        importance="high"
    )
    
    if result['success']:
        task = result['task']
        await ctx.send(
            f"✅ **Task Created Successfully!**\n\n"
            f"List: {list_name}\n"
            f"{format_task(task)}\n\n"
            f"💡 Use **'todo tasks'** to view all your tasks"
        )
    else:
        await ctx.send(f"❌ Error creating task: {result['error']}")


@app.on_message_pattern("todo new list")
async def handle_create_list_command(ctx: ActivityContext[MessageActivity]):
    """Handle todo new list command to create a new task list."""
    
    if not ctx.is_signed_in:
        await ctx.send("🔐 Please sign in first to access Microsoft Graph.")
        await ctx.sign_in()
        return
    
    graph = ctx.user_graph
    
    if not graph:
        await ctx.send("❌ Could not access Microsoft Graph.")
        return
    
    # Example: Create a new task list
    from datetime import datetime
    list_name = f"Project Tasks - {datetime.now().strftime('%Y-%m-%d')}"
    
    result = await create_task_list(graph, list_name)
    
    if result['success']:
        created_list = result['list']
        await ctx.send(
            f"✅ **Task List Created!**\n\n"
            f"{format_task_list(created_list)}\n\n"
            f"💡 Use **'todo lists'** to view all your lists"
        )
    else:
        await ctx.send(f"❌ Error creating list: {result['error']}")


@app.on_message_pattern("generate todo")
async def handle_generate_todo_command(ctx: ActivityContext[MessageActivity]):
    """Handle generate todo command - analyzes emails and Teams messages to create tasks."""
    
    if not ctx.is_signed_in:
        await ctx.send("🔐 Please sign in first to access Microsoft Graph.")
        await ctx.sign_in()
        return
    
    graph = ctx.user_graph
    
    if not graph:
        await ctx.send("❌ Could not access Microsoft Graph.")
        return
    
    await ctx.send("🤖 **AI Todo Generation Started**\n\nThis may take a moment...\n")
    
    try:
        # Step 1: Fetch emails from the last 40 days
        from datetime import datetime, timedelta
        end_date = datetime.utcnow().isoformat() + 'Z'
        start_date = (datetime.utcnow() - timedelta(days=40)).isoformat() + 'Z'
        
        await ctx.send("📧 Fetching emails from the last 40 days...")
        emails_result = await get_mail_inbox(graph, start_date=start_date, end_date=end_date)
        
        if not emails_result['success']:
            await ctx.send(f"⚠️ Warning: Could not fetch emails: {emails_result['error']}")
            emails = []
        else:
            emails = emails_result['emails']
            await ctx.send(f"✅ Found {len(emails)} emails")
        
        # Step 2: Fetch recent Teams messages
        await ctx.send("💬 Fetching recent Teams messages...")
        teams_result = await get_all_recent_teams_messages(graph, num_chats=5, messages_per_chat=10)
        
        if not teams_result['success']:
            await ctx.send(f"⚠️ Warning: Could not fetch Teams messages: {teams_result['error']}")
            teams_messages = []
        else:
            teams_messages = teams_result['all_messages']
            await ctx.send(f"✅ Found {len(teams_messages)} Teams messages")
        
        # Step 3: Run AI todo list generation
        await ctx.send("🧠 Analyzing content with AI to extract tasks...")
        
        todo_result = await run_generate_todo_list(
            raw_emails=emails,
            raw_teams_messages=teams_messages
        )
        
        all_tasks = todo_result.get('all_tasks', [])
        email_task_count = len(todo_result.get('email_tasks', []))
        teams_task_count = len(todo_result.get('teams_tasks', []))
        
        if not all_tasks:
            await ctx.send("ℹ️ No actionable tasks were found in your emails and Teams messages.")
            return
        
        await ctx.send(
            f"✅ **AI Analysis Complete!**\n\n"
            f"Found {len(all_tasks)} tasks:\n"
            f"• {email_task_count} from emails\n"
            f"• {teams_task_count} from Teams messages\n\n"
            f"Now creating tasks in Microsoft To Do..."
        )
        
        # Step 4: Get the default To Do list
        lists_result = await get_todo_lists(graph)
        
        if not lists_result['success'] or not lists_result['lists']:
            await ctx.send("❌ No To Do lists found. Please create a list first using **'todo new list'**")
            return
        
        default_list = lists_result['lists'][0]
        list_id = default_list.id
        list_name = default_list.display_name
        
        await ctx.send(f"📋 Adding tasks to: **{list_name}**\n")
        
        # Step 5: Create tasks in Microsoft To Do
        created_count = 0
        failed_count = 0
        
        for task_entry in all_tasks:
            task_title = task_entry.get('task', 'Untitled Task')
            priority = task_entry.get('priority', 'normal')
            comments = task_entry.get('comments', '')
            due_date = task_entry.get('due_date', None)
            
            # Map priority to importance
            importance = "normal"
            if priority and 'high' in priority.lower():
                importance = "high"
            elif priority and 'low' in priority.lower():
                importance = "low"
            
            # Combine comments into body
            body = comments if comments else None
            
            # Create the task
            create_result = await create_task(
                graph,
                list_id=list_id,
                title=task_title,
                body=body,
                due_date=due_date,
                importance=importance
            )
            
            if create_result['success']:
                created_count += 1
            else:
                failed_count += 1
                logger.warning(f"Failed to create task '{task_title}': {create_result['error']}")
        
        # Step 6: Show results
        success_msg = f"✅ **Successfully created {created_count} tasks in '{list_name}'!**\n\n"
        
        if failed_count > 0:
            success_msg += f"⚠️ {failed_count} task(s) could not be created.\n\n"
        
        # Show a preview of the first few tasks
        if created_count > 0:
            success_msg += "**Sample of created tasks:**\n"
            for i, task_entry in enumerate(all_tasks[:3], 1):
                task_title = task_entry.get('task', 'Untitled Task')
                priority = task_entry.get('priority', 'neutral')
                priority_icon = "🔴" if 'high' in priority.lower() else "🟢" if 'low' in priority.lower() else "🟡"
                success_msg += f"{i}. {priority_icon} {task_title}\n"
            
            if len(all_tasks) > 3:
                success_msg += f"... and {len(all_tasks) - 3} more\n"
        
        success_msg += f"\n💡 Use **'todo tasks'** to view all your tasks"
        
        await ctx.send(success_msg)
        
    except Exception as e:
        logger.error(f"Error in generate todo command: {str(e)}", exc_info=True)
        await ctx.send(f"❌ **Error generating tasks:** {str(e)}\n\nPlease try again or contact support if the issue persists.")


@app.on_message
async def handle_default_message(ctx: ActivityContext[MessageActivity]):
    """Handle default message - trigger signin."""
    if ctx.is_signed_in:
        await ctx.send(
            "✅ You are already signed in!\n\n"
            "**Available commands:**\n\n"
            "📧 **Email & Teams:**\n"
            "• **emails** - View recent emails\n"
            "• **teams** - View recent Teams messages\n"
            "• **chats** - List your recent chats\n\n"
            "✅ **To Do Tasks:**\n"
            "• **add task {text}** - ⚡ Quick: Create a task from your message\n"
            "• **generate todo** - 🤖 AI-powered: Analyze emails & Teams to create tasks\n"
            "• **todo lists** - View all your task lists\n"
            "• **todo tasks** - View tasks from default list\n"
            "• **todo create** - Create a sample task\n"
            "• **todo new list** - Create a new task list\n\n"
            "👤 **Profile & Auth:**\n"
            "• **profile** - View your profile\n"
            "• **signout** - Sign out when done"
        )
    else:
        await ctx.send("🔐 Please sign in to access Microsoft Graph...")
        await ctx.sign_in()


@app.event("sign_in")
async def handle_sign_in_event(event: SignInEvent):
    """Handle successful sign-in events."""
    await event.activity_ctx.send(
        "✅ **Successfully signed in!**\n\n"
        "**Available commands:**\n\n"
        "📧 **Email & Teams:**\n"
        "• **emails** - View recent emails\n"
        "• **teams** - View recent Teams messages\n"
        "• **chats** - List your recent chats\n\n"
        "✅ **To Do Tasks:**\n"
        "• **add task {text}** - ⚡ Quick: Create a task from your message\n"
        "• **generate todo** - 🤖 AI-powered: Analyze emails & Teams to create tasks\n"
        "• **todo lists** - View all your task lists\n"
        "• **todo tasks** - View tasks from default list\n"
        "• **todo create** - Create a sample task\n"
        "• **todo new list** - Create a new task list\n\n"
        "👤 **Profile & Auth:**\n"
        "• **profile** - View your profile\n"
        "• **signout** - Sign out when done"
    )


def main():
    asyncio.run(app.start())


if __name__ == "__main__":
    main()
