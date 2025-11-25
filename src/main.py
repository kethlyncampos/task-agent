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
    update_task, delete_task, create_task_list, format_task, format_task_list,
    get_incomplete_tasks_from_list
)
from todo_list_generation.todo_list_generation import (
    run_generate_todo_list, 
    generate_single_task_from_user_message,
    complete_task_generation_workflow,
    format_task_body
)

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
    
    # Obtenha a primeira (padrão) lista de tarefas
    lists_result = await get_todo_lists(graph)
    
    if not lists_result['success'] or not lists_result['lists']:
        await ctx.send("❌ Nenhuma lista do To Do encontrada. Por favor, crie uma lista primeiro.")
        return
    
    default_list = lists_result['lists'][0]
    list_id = default_list.id
    list_name = default_list.display_name
    
    await ctx.send("🤖 Gerando tarefa detalhada com IA...")
    
    try:
        # Gerar tarefa detalhada com IA
        task_entry = await generate_single_task_from_user_message(task_text)
        
        if not task_entry or not task_entry.get('task'):
            await ctx.send("❌ Não foi possível gerar a tarefa. Por favor, tente novamente com uma mensagem mais clara.")
            return
        
        # Extrair detalhes da tarefa
        task_title = task_entry.get('task', task_text)
        priority = task_entry.get('priority', 'normal')
        comments = task_entry.get('comments', '')
        person_envolved = task_entry.get('person_envolved', '')
        due_date_raw = task_entry.get('due_date', None)
        
        # Validar e formatar a data de vencimento
        due_date = None
        if due_date_raw:
            due_date = validate_and_format_date(str(due_date_raw))
            if not due_date:
                logger.warning(f"Data de vencimento inválida da IA: {due_date_raw}, ignorando")
        
        # Mapear prioridade para importância
        importance = "normal"
        if priority and 'high' in priority.lower():
            importance = "high"
        elif priority and 'low' in priority.lower():
            importance = "low"
        
        # Formatar corpo da tarefa com comentários e pessoa envolvida
        task_body = format_task_body(comments=comments, person_envolved=person_envolved)
        
        # Criar a tarefa com os detalhes gerados pela IA
        result = await create_task(
            graph,
            list_id=list_id,
            title=task_title,
            body=task_body,
            due_date=due_date,
            importance=importance
        )
        
        if result['success']:
            task = result['task']
            await ctx.send(
                f"✅ **Tarefa criada com sucesso!**\n\n"
                f"Lista: {list_name}\n"
                f"{format_task(task)}\n\n"
                f"💡 Use **'todo tasks'** para ver todas as suas tarefas"
            )
        else:
            await ctx.send(f"❌ Erro ao criar a tarefa: {result['error']}")
            
    except Exception as e:
        logger.error(f"Erro no comando adicionar tarefa: {str(e)}", exc_info=True)
        await ctx.send(
            f"❌ **Erro ao gerar tarefa:** {str(e)}\n\n"
            "Criando tarefa simples sem IA..."
        )
        
        # Alternativa: criar tarefa simples sem IA
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
                f"✅ **Tarefa criada!**\n\n"
                f"Lista: {list_name}\n"
                f"{format_task(task)}"
            )
        else:
            await ctx.send(f"❌ Erro ao criar a tarefa: {result['error']}")


@app.on_message_pattern("todo create")
async def handle_todo_create_command(ctx: ActivityContext[MessageActivity]):
    """Lida com o comando 'todo create' para criar uma nova tarefa."""
    
    if not ctx.is_signed_in:
        await ctx.send("🔐 Por favor, faça login primeiro para acessar o Microsoft Graph.")
        await ctx.sign_in()
        return
    
    graph = ctx.user_graph
    
    if not graph:
        await ctx.send("❌ Não foi possível acessar o Microsoft Graph.")
        return
    
    # Obtenha a primeira (padrão) lista de tarefas
    lists_result = await get_todo_lists(graph)
    
    if not lists_result['success'] or not lists_result['lists']:
        await ctx.send("❌ Nenhuma lista do To Do encontrada. Por favor, crie uma lista primeiro.")
        return
    
    default_list = lists_result['lists'][0]
    list_id = default_list.id
    list_name = default_list.display_name
    
    # Exemplo: Criar uma tarefa de demonstração
    from datetime import datetime, timedelta
    due_date = (datetime.utcnow() + timedelta(days=7)).strftime("%Y-%m-%d")
    
    result = await create_task(
        graph,
        list_id=list_id,
        title="Tarefa de demonstração - Revisar documentação do projeto",
        body="Esta é uma tarefa modelo criada via Microsoft Graph API. Por favor, revise e atualize toda a documentação do projeto.",
        due_date=due_date,
        importance="high"
    )
    
    if result['success']:
        task = result['task']
        await ctx.send(
            f"✅ **Tarefa criada com sucesso!**\n\n"
            f"Lista: {list_name}\n"
            f"{format_task(task)}\n\n"
            f"💡 Use **'todo tasks'** para ver todas as suas tarefas"
        )
    else:
        await ctx.send(f"❌ Erro ao criar a tarefa: {result['error']}")


@app.on_message_pattern("todo new list")
async def handle_create_list_command(ctx: ActivityContext[MessageActivity]):
    """Lida com o comando 'todo new list' para criar uma nova lista de tarefas."""
    
    if not ctx.is_signed_in:
        await ctx.send("🔐 Por favor, faça login primeiro para acessar o Microsoft Graph.")
        await ctx.sign_in()
        return
    
    graph = ctx.user_graph
    
    if not graph:
        await ctx.send("❌ Não foi possível acessar o Microsoft Graph.")
        return
    
    # Exemplo: Criar uma nova lista de tarefas
    from datetime import datetime
    list_name = f"Tarefas do Projeto - {datetime.now().strftime('%Y-%m-%d')}"
    
    result = await create_task_list(graph, list_name)
    
    if result['success']:
        created_list = result['list']
        await ctx.send(
            f"✅ **Lista de tarefas criada!**\n\n"
            f"{format_task_list(created_list)}\n\n"
            f"💡 Use **'todo lists'** para ver todas as suas listas"
        )
    else:
        await ctx.send(f"❌ Erro ao criar a lista: {result['error']}")


@app.on_message_pattern("generate todo")
async def handle_generate_todo_command(ctx: ActivityContext[MessageActivity]):
    """Lida com o comando generate todo - analisa emails e mensagens do Teams para criar tarefas."""
    
    if not ctx.is_signed_in:
        await ctx.send("🔐 Por favor, faça login primeiro para acessar o Microsoft Graph.")
        await ctx.sign_in()
        return
    
    graph = ctx.user_graph
    
    if not graph:
        await ctx.send("❌ Não foi possível acessar o Microsoft Graph.")
        return
    
    await ctx.send("🤖 **Geração de tarefas com IA iniciada**\n\nIsso pode levar um momento...\n")
    
    try:
        # Etapa 1: Buscar emails dos últimos 40 dias
        from datetime import datetime, timedelta
        end_date = datetime.utcnow().isoformat() + 'Z'
        start_date = (datetime.utcnow() - timedelta(days=40)).isoformat() + 'Z'
        
        await ctx.send("📧 Buscando emails dos últimos 40 dias...")
        emails_result = await get_mail_inbox(graph, start_date=start_date, end_date=end_date)
        
        if not emails_result['success']:
            await ctx.send(f"⚠️ Aviso: Não foi possível buscar emails: {emails_result['error']}")
            emails = []
        else:
            emails = emails_result['emails']
            await ctx.send(f"✅ {len(emails)} emails encontrados")
        
        # Etapa 2: Buscar mensagens recentes do Teams
        await ctx.send("💬 Buscando mensagens recentes do Teams...")
        teams_result = await get_all_recent_teams_messages(graph, num_chats=5, messages_per_chat=10)
        
        if not teams_result['success']:
            await ctx.send(f"⚠️ Aviso: Não foi possível buscar mensagens do Teams: {teams_result['error']}")
            teams_messages = []
        else:
            teams_messages = teams_result['all_messages']
            await ctx.send(f"✅ {len(teams_messages)} mensagens do Teams encontradas")
        
        # Etapa 3: Rodar geração de tarefas com IA
        await ctx.send("🧠 Analisando o conteúdo com IA para extrair tarefas...")
        
        todo_result = await run_generate_todo_list(
            raw_emails=emails,
            raw_teams_messages=teams_messages
        )
        
        all_tasks = todo_result.get('all_tasks', [])
        email_task_count = len(todo_result.get('email_tasks', []))
        teams_task_count = len(todo_result.get('teams_tasks', []))
        
        if not all_tasks:
            await ctx.send("ℹ️ Nenhuma tarefa acionável foi encontrada nos seus emails e mensagens do Teams.")
            return
        
        await ctx.send(
            f"✅ **Análise da IA concluída!**\n\n"
            f"Foram encontradas {len(all_tasks)} tarefas:\n"
            f"• {email_task_count} dos emails\n"
            f"• {teams_task_count} das mensagens do Teams\n\n"
            f"Agora criando tarefas no Microsoft To Do..."
        )
        
        # Etapa 4: Obter a lista padrão do To Do
        lists_result = await get_todo_lists(graph)
        
        if not lists_result['success'] or not lists_result['lists']:
            await ctx.send("❌ Nenhuma lista do To Do encontrada. Por favor, crie uma lista primeiro, usando **'todo new list'**")
            return
        
        default_list = lists_result['lists'][0]
        list_id = default_list.id
        list_name = default_list.display_name
        
        await ctx.send(f"📋 Adicionando tarefas em: **{list_name}**\n")
        
        # Etapa 5: Criar tarefas no Microsoft To Do
        created_count = 0
        failed_count = 0
        
        for task_entry in all_tasks:
            task_title = task_entry.get('task', 'Tarefa sem título')
            priority = task_entry.get('priority', 'normal')
            comments = task_entry.get('comments', '')
            person_envolved = task_entry.get('person_envolved', '')
            due_date = task_entry.get('due_date', None)
            
            # Mapear prioridade para importância
            importance = "normal"
            if priority and 'high' in priority.lower():
                importance = "high"
            elif priority and 'low' in priority.lower():
                importance = "low"
            
            # Formatar corpo da tarefa com comentários e pessoa envolvida
            task_body = format_task_body(comments=comments, person_envolved=person_envolved)
            
            # Criar a tarefa
            create_result = await create_task(
                graph,
                list_id=list_id,
                title=task_title,
                body=task_body,
                due_date=due_date,
                importance=importance
            )
            
            if create_result['success']:
                created_count += 1
            else:
                failed_count += 1
                logger.warning(f"Falha ao criar tarefa '{task_title}': {create_result['error']}")
        
        # Etapa 6: Mostrar resultados
        success_msg = f"✅ **{created_count} tarefas criadas com sucesso na lista '{list_name}'!**\n\n"
        
        if failed_count > 0:
            success_msg += f"⚠️ {failed_count} tarefa(s) não puderam ser criadas.\n\n"
        
        # Mostrar amostra das primeiras tarefas criadas
        if created_count > 0:
            success_msg += "**Exemplo de tarefas criadas:**\n"
            for i, task_entry in enumerate(all_tasks[:3], 1):
                task_title = task_entry.get('task', 'Tarefa sem título')
                priority = task_entry.get('priority', 'neutra')
                priority_icon = "🔴" if 'high' in priority.lower() else "🟢" if 'low' in priority.lower() else "🟡"
                success_msg += f"{i}. {priority_icon} {task_title}\n"
            
            if len(all_tasks) > 3:
                success_msg += f"... e mais {len(all_tasks) - 3}\n"
        
        success_msg += f"\n💡 Use **'todo tasks'** para ver todas as suas tarefas"
        
        await ctx.send(success_msg)
        
    except Exception as e:
        logger.error(f"Erro no comando generate todo: {str(e)}", exc_info=True)
        await ctx.send(f"❌ **Erro ao gerar tarefas:** {str(e)}\n\nPor favor, tente novamente ou entre em contato com o suporte se o problema persistir.")


@app.on_message_pattern("generate new tasks")
async def handle_generate_new_tasks_command(ctx: ActivityContext[MessageActivity]):
    """
    Lida com o comando generate new tasks - analisa emails e mensagens do Teams
    para criar tarefas com deduplicação inteligente contra tarefas já existentes e incompletas.
    """
    
    if not ctx.is_signed_in:
        await ctx.send("🔐 Por favor, faça login primeiro para acessar o Microsoft Graph.")
        await ctx.sign_in()
        return
    
    graph = ctx.user_graph
    
    if not graph:
        await ctx.send("❌ Não foi possível acessar o Microsoft Graph.")
        return
    
    await ctx.send(
        "🚀 **Geração de tarefas com IA e deduplicação inteligente**\n\n"
        "Isso irá:\n"
        "1️⃣ Analisar seus emails e mensagens recentes do Teams\n"
        "2️⃣ Extrair tarefas acionáveis usando IA\n"
        "3️⃣ Comparar com suas tarefas incompletas existentes\n"
        "4️⃣ Criar apenas tarefas NOVAS e não duplicadas\n\n"
        "⏳ Processando... Isso pode demorar um pouco..."
    )
    
    try:
        # Etapa 1: Buscar emails dos últimos 2 dias
        from datetime import datetime, timedelta, timezone
        end_date = datetime.now(timezone.utc).isoformat().replace('+00:00', 'Z')
        start_date = (datetime.now(timezone.utc) - timedelta(days=2)).isoformat().replace('+00:00', 'Z')
        
        await ctx.send("📧 Buscando emails dos últimos 2 dias...")
        emails_result = await get_mail_inbox(graph, start_date=start_date, end_date=end_date)
        
        if not emails_result['success']:
            await ctx.send(f"⚠️ Aviso: Não foi possível buscar emails: {emails_result['error']}")
            emails = []
        else:
            emails = emails_result['emails']
            await ctx.send(f"✅ {len(emails)} emails encontrados")
        
        # Etapa 2: Buscar mensagens recentes do Teams
        await ctx.send("💬 Buscando mensagens recentes do Teams...")
        teams_result = await get_all_recent_teams_messages(graph, num_chats=5, messages_per_chat=10)
        
        if not teams_result['success']:
            await ctx.send(f"⚠️ Aviso: Não foi possível buscar mensagens do Teams: {teams_result['error']}")
            teams_messages = []
        else:
            teams_messages = teams_result['all_messages']
            await ctx.send(f"✅ {len(teams_messages)} mensagens do Teams encontradas")
        
        # Conferir se há conteúdo para analisar
        if not emails and not teams_messages:
            await ctx.send(
                "ℹ️ **Nenhum conteúdo encontrado para análise**\n\n"
                "Nenhum email ou mensagem do Teams foi encontrado no período especificado.\n"
                "Tente novamente mais tarde, quando houver novas comunicações."
            )
            return
        
        # Etapa 3: Buscar a lista padrão do To Do para criação das tarefas
        lists_result = await get_todo_lists(graph)
        
        if not lists_result['success'] or not lists_result['lists']:
            await ctx.send("❌ Nenhuma lista do To Do encontrada. Por favor, crie uma lista primeiro, usando **'todo new list'**")
            return
        
        default_list = lists_result['lists'][0]
        list_id = default_list.id
        list_name = default_list.display_name
        
        await ctx.send(f"📋 Lista alvo: **{list_name}**")
        
        # Etapa 4: Rodar workflow completo com deduplicação
        await ctx.send("🧠 Executando análise com IA e deduplicação...")
        
        workflow_result = await complete_task_generation_workflow(
            graph_client=graph,
            raw_emails=emails,
            raw_teams_messages=teams_messages,
            target_list_id=list_id,
            user_id=ctx.activity.from_.id if hasattr(ctx.activity, 'from_') and ctx.activity.from_ else "desconhecido",
            session_id=ctx.activity.conversation.id if hasattr(ctx.activity, 'conversation') and ctx.activity.conversation else "desconhecido"
        )
        
        # Etapa 5: Exibir resultados detalhados
        if workflow_result['success']:
            created_tasks = workflow_result.get('created_tasks', [])
            duplicate_tasks = workflow_result.get('duplicate_tasks', [])
            all_generated_tasks = workflow_result.get('all_generated_tasks', [])
            creation_errors = workflow_result.get('creation_errors', [])
            
            # Montar resposta detalhada
            response = f"""
✅ **Geração de tarefas concluída!**

{workflow_result['summary']}
"""
            
            # Amostra das tarefas criadas
            if created_tasks:
                response += "\n\n**📝 Novas tarefas criadas:**\n"
                for i, task_info in enumerate(created_tasks[:5], 1):
                    response += f"{i}. {task_info.get('title', 'Sem título')}\n"
                
                if len(created_tasks) > 5:
                    response += f"... e mais {len(created_tasks) - 5}\n"
            
            # Exibir duplicatas puladas
            if duplicate_tasks:
                response += f"\n\n**🔄 Duplicatas ignoradas ({len(duplicate_tasks)}):**\n"
                for i, task_data in enumerate(duplicate_tasks[:3], 1):
                    task_title = task_data.get('task', 'Desconhecido')[:60]
                    response += f"{i}. {task_title}\n"
                
                if len(duplicate_tasks) > 3:
                    response += f"... e mais {len(duplicate_tasks) - 3}\n"
            
            # Exibir erros, se presentes
            if creation_errors:
                response += f"\n\n⚠️ **Erros ({len(creation_errors)}):**\n"
                for error_info in creation_errors[:3]:
                    if isinstance(error_info, dict):
                        task_name = error_info.get('task', 'Desconhecido')[:40]
                        error_msg = error_info.get('error', 'Erro desconhecido')[:60]
                        response += f"• {task_name}: {error_msg}\n"
                    else:
                        response += f"• {str(error_info)[:80]}\n"
            
            response += f"\n\n💡 Use **'todo tasks'** para ver todas as suas tarefas"
            
            await ctx.send(response)
            
        else:
            error_msg = workflow_result.get('summary', 'Erro desconhecido')
            await ctx.send(f"❌ **Falha no workflow:**\n\n{error_msg}\n\nPor favor, tente novamente ou entre em contato com o suporte.")
        
    except Exception as e:
        logger.error(f"Erro no comando generate new tasks: {str(e)}", exc_info=True)
        await ctx.send(
            f"❌ **Erro ao gerar tarefas:**\n\n{str(e)}\n\n"
            "Por favor, tente novamente ou entre em contato com o suporte se o problema persistir."
        )


@app.on_message_pattern("delete all tasks")
async def handle_delete_all_tasks_command(ctx: ActivityContext[MessageActivity]):
    """
    Lida com o comando delete all tasks - deleta todas as tarefas de todas as listas.
    Requer confirmação do usuário por segurança.
    """
    
    if not ctx.is_signed_in:
        await ctx.send("🔐 Por favor, faça login primeiro para acessar o Microsoft Graph.")
        await ctx.sign_in()
        return
    
    graph = ctx.user_graph
    
    if not graph:
        await ctx.send("❌ Não foi possível acessar o Microsoft Graph.")
        return
    
    await ctx.send(
        "⚠️ **AVISO: Operação Perigosa!**\n\n"
        "Você está prestes a **DELETAR TODAS AS TAREFAS** de todas as suas listas no Microsoft To Do.\n\n"
        "Esta ação **NÃO PODE SER DESFEITA**!\n\n"
        "Para confirmar, responda com: **CONFIRMAR EXCLUSÃO**\n"
        "Para cancelar, responda com qualquer outra coisa."
    )


@app.on_message_pattern(re.compile(r"^CONFIRMAR EXCLUSÃO$", re.IGNORECASE))
async def handle_confirm_delete_all_tasks(ctx: ActivityContext[MessageActivity]):
    """Confirma e executa a exclusão de todas as tarefas."""
    
    if not ctx.is_signed_in:
        await ctx.send("🔐 Por favor, faça login primeiro.")
        return
    
    graph = ctx.user_graph
    
    if not graph:
        await ctx.send("❌ Não foi possível acessar o Microsoft Graph.")
        return
    
    await ctx.send("🗑️ **Iniciando exclusão de todas as tarefas...**\n\nIsso pode levar alguns momentos...")
    
    try:
        # Etapa 1: Buscar todas as listas
        lists_result = await get_todo_lists(graph)
        
        if not lists_result['success']:
            await ctx.send(f"❌ Erro ao buscar listas: {lists_result['error']}")
            return
        
        lists = lists_result['lists']
        
        if not lists:
            await ctx.send("ℹ️ Nenhuma lista de tarefas encontrada.")
            return
        
        await ctx.send(f"📋 Encontradas {len(lists)} lista(s). Processando...")
        
        # Contadores
        total_deleted = 0
        total_failed = 0
        lists_processed = 0
        
        # Etapa 2: Para cada lista, buscar e deletar todas as tarefas
        for task_list in lists:
            list_id = task_list.id
            list_name = task_list.display_name or "Lista sem nome"
            
            # Buscar todas as tarefas da lista
            tasks_result = await get_tasks_from_list(graph, list_id)
            
            if not tasks_result['success']:
                await ctx.send(f"⚠️ Erro ao buscar tarefas de '{list_name}': {tasks_result['error']}")
                continue
            
            tasks = tasks_result['tasks']
            
            if not tasks:
                await ctx.send(f"✓ '{list_name}': 0 tarefas")
                lists_processed += 1
                continue
            
            await ctx.send(f"🗑️ Deletando {len(tasks)} tarefa(s) de '{list_name}'...")
            
            # Deletar cada tarefa
            deleted_count = 0
            failed_count = 0
            
            for task in tasks:
                task_id = task.id
                task_title = task.title or "Sem título"
                
                delete_result = await delete_task(graph, list_id, task_id)
                
                if delete_result['success']:
                    deleted_count += 1
                    total_deleted += 1
                else:
                    failed_count += 1
                    total_failed += 1
                    logger.warning(f"Falha ao deletar tarefa '{task_title}': {delete_result['error']}")
            
            lists_processed += 1
            await ctx.send(
                f"✓ '{list_name}': {deleted_count} deletada(s)"
                + (f", {failed_count} falha(s)" if failed_count > 0 else "")
            )
        
        # Etapa 3: Resumo final
        summary_msg = f"""
✅ **Exclusão Concluída!**

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
📊 **Resumo:**
• Listas processadas: {lists_processed}
• Tarefas deletadas: {total_deleted}
"""
        
        if total_failed > 0:
            summary_msg += f"• Falhas: {total_failed}\n"
        
        summary_msg += "━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
        
        if total_deleted > 0:
            summary_msg += "\n\n🎉 Todas as tarefas foram removidas com sucesso!"
        
        await ctx.send(summary_msg)
        
    except Exception as e:
        logger.error(f"Erro no comando delete all tasks: {str(e)}", exc_info=True)
        await ctx.send(
            f"❌ **Erro ao deletar tarefas:**\n\n{str(e)}\n\n"
            "Algumas tarefas podem ter sido deletadas antes do erro."
        )


@app.on_message_pattern("delete incomplete tasks")
async def handle_delete_incomplete_tasks_command(ctx: ActivityContext[MessageActivity]):
    """
    Lida com o comando delete incomplete tasks - deleta apenas tarefas incompletas.
    """
    
    if not ctx.is_signed_in:
        await ctx.send("🔐 Por favor, faça login primeiro para acessar o Microsoft Graph.")
        await ctx.sign_in()
        return
    
    graph = ctx.user_graph
    
    if not graph:
        await ctx.send("❌ Não foi possível acessar o Microsoft Graph.")
        return
    
    await ctx.send(
        "⚠️ **AVISO!**\n\n"
        "Você está prestes a **DELETAR TODAS AS TAREFAS INCOMPLETAS**.\n\n"
        "Tarefas concluídas não serão afetadas.\n\n"
        "Para confirmar, responda com: **CONFIRMAR EXCLUSÃO INCOMPLETAS**\n"
        "Para cancelar, responda com qualquer outra coisa."
    )


@app.on_message_pattern(re.compile(r"^CONFIRMAR EXCLUSÃO INCOMPLETAS$", re.IGNORECASE))
async def handle_confirm_delete_incomplete_tasks(ctx: ActivityContext[MessageActivity]):
    """Confirma e executa a exclusão de tarefas incompletas."""
    
    if not ctx.is_signed_in:
        await ctx.send("🔐 Por favor, faça login primeiro.")
        return
    
    graph = ctx.user_graph
    
    if not graph:
        await ctx.send("❌ Não foi possível acessar o Microsoft Graph.")
        return
    
    await ctx.send("🗑️ **Deletando tarefas incompletas...**")
    
    try:
        # Buscar todas as listas
        lists_result = await get_todo_lists(graph)
        
        if not lists_result['success']:
            await ctx.send(f"❌ Erro ao buscar listas: {lists_result['error']}")
            return
        
        lists = lists_result['lists']
        
        if not lists:
            await ctx.send("ℹ️ Nenhuma lista de tarefas encontrada.")
            return
        
        total_deleted = 0
        total_failed = 0
        
        for task_list in lists:
            list_id = task_list.id
            list_name = task_list.display_name or "Lista sem nome"
            
            # Buscar apenas tarefas incompletas
            tasks_result = await get_incomplete_tasks_from_list(graph, list_id)
            
            if not tasks_result['success']:
                continue
            
            tasks = tasks_result['tasks']
            
            if not tasks:
                continue
            
            await ctx.send(f"🗑️ '{list_name}': deletando {len(tasks)} tarefa(s) incompleta(s)...")
            
            for task in tasks:
                delete_result = await delete_task(graph, list_id, task.id)
                
                if delete_result['success']:
                    total_deleted += 1
                else:
                    total_failed += 1
        
        await ctx.send(
            f"✅ **Concluído!**\n\n"
            f"• Tarefas incompletas deletadas: {total_deleted}\n"
            + (f"• Falhas: {total_failed}\n" if total_failed > 0 else "") +
            f"\n💡 Use **'todo tasks'** para verificar as tarefas restantes"
        )
        
    except Exception as e:
        logger.error(f"Erro ao deletar tarefas incompletas: {str(e)}", exc_info=True)
        await ctx.send(f"❌ **Erro:** {str(e)}")


@app.on_message
async def handle_default_message(ctx: ActivityContext[MessageActivity]):
    """Lida com mensagens padrão - aciona login."""
    if ctx.is_signed_in:
        await ctx.send(
            "✅ Você já está conectado!\n\n"
            "**Comandos disponíveis:**\n\n"
            "📧 **Email & Teams:**\n"
            "• **emails** - Ver emails recentes\n"
            "• **teams** - Ver mensagens recentes do Teams\n"
            "• **chats** - Listar seus chats recentes\n\n"
            "✅ **Tarefas do To Do:**\n"
            "• **add task {texto}** - ⚡ Rápido: crie uma tarefa a partir da sua mensagem\n"
            "• **generate new tasks** - 🚀 IA: Deduplicação inteligente (Recomendado!)\n"
            "• **generate todo** - 🤖 IA: Analisa emails e Teams para criar tarefas\n"
            "• **todo lists** - Ver todas as suas listas de tarefas\n"
            "• **todo tasks** - Ver tarefas da lista padrão\n"
            "• **todo create** - Criar uma tarefa de exemplo\n"
            "• **todo new list** - Criar uma nova lista de tarefas\n"
            "• **delete all tasks** - 🗑️ Deletar TODAS as tarefas (requer confirmação)\n"
            "• **delete incomplete tasks** - 🗑️ Deletar apenas tarefas incompletas\n\n"
            "👤 **Perfil & Autenticação:**\n"
            "• **profile** - Ver seu perfil\n"
            "• **signout** - Sair da conta"
        )
    else:
        await ctx.send("🔐 Por favor, faça login para acessar o Microsoft Graph...")
        await ctx.sign_in()


@app.event("sign_in")
async def handle_sign_in_event(event: SignInEvent):
    """Lida com eventos de login bem-sucedido."""
    await event.activity_ctx.send(
        "✅ **Login realizado com sucesso!**\n\n"
        "**Comandos disponíveis:**\n\n"
        "📧 **Email & Teams:**\n"
        "• **emails** - Ver emails recentes\n"
        "• **teams** - Ver mensagens recentes do Teams\n"
        "• **chats** - Listar seus chats recentes\n\n"
        "✅ **Tarefas do To Do:**\n"
        "• **add task {texto}** - ⚡ Rápido: crie uma tarefa a partir da sua mensagem\n"
        "• **generate new tasks** - 🚀 IA: Deduplicação inteligente (Recomendado!)\n"
        "• **generate todo** - 🤖 IA: Analisa emails e Teams para criar tarefas\n"
        "• **todo lists** - Ver todas as suas listas de tarefas\n"
        "• **todo tasks** - Ver tarefas da lista padrão\n"
        "• **todo create** - Criar uma tarefa de exemplo\n"
        "• **todo new list** - Criar uma nova lista de tarefas\n"
        "• **delete all tasks** - 🗑️ Deletar TODAS as tarefas (requer confirmação)\n"
        "• **delete incomplete tasks** - 🗑️ Deletar apenas tarefas incompletas\n\n"
        "👤 **Perfil & Autenticação:**\n"
        "• **profile** - Ver seu perfil\n"
        "• **signout** - Sair da conta"
    )


def main():
    asyncio.run(app.start())


if __name__ == "__main__":
    main()
