"""
Todo List Generation Workflow

AI-powered todo list generation system that processes text content
to extract and organize actionable tasks using LangGraph workflow orchestration.

"""

from graph_api.mail_inbox import format_email
from graph_api.teams_messages import format_teams_message
from typing import TypedDict

from langchain_core.messages import HumanMessage
from langchain_core.runnables import RunnableConfig
from langgraph.graph import StateGraph, START, END

from todo_list_generation.prompts import todo_list_generation_prompt, single_task_generation_prompt
from todo_list_generation.todo_list_model import TodoList

from langchain_openai import ChatOpenAI

import os
from logging import getLogger
from src.settings import get_settings


l = getLogger(__name__)

# =============================================================================
# MODEL CONFIGURATIONS
# =============================================================================

MODEL_CONFIGS = {
    "todo_list_generation": {
        "model": "gpt-4o",
        "temperature": 0.0,
        "max_tokens": 10000
    }
}


def format_task_body(comments: str = None, person_envolved: str = None) -> str:
    """
    Format task body by combining comments and person involved.
    
    Args:
        comments: Task comments/notes
        person_envolved: Person involved in the task
    
    Returns:
        Formatted task body string
    """
    body_parts = []
    
    if comments:
        body_parts.append(comments.strip())
    
    if person_envolved:
        body_parts.append(f"\n👤 Pessoa envolvida: {person_envolved.strip()}")
    
    return "\n".join(body_parts) if body_parts else None



class GraphState(TypedDict):
    content: str
    todo_entries: list
    emails: list
    email_tasks: list
    teams_messages: list
    teams_tasks: list
    raw_emails: list  # Store raw email objects
    raw_teams_messages: list  # Store raw teams messages

async def get_emails(state: GraphState, config: RunnableConfig) -> dict:
    """Get emails from state (already fetched by the command)"""
    raw_emails = state.get("raw_emails", [])
    
    parsed_emails = []
    for email in raw_emails:
        # Convert SDK email object to dictionary format expected by format_email
        email_dict = {
            'subject': email.subject or '(No subject)',
            'sender_name': email.from_.email_address.name if email.from_ and email.from_.email_address else 'Unknown',
            'sender': email.from_.email_address.address if email.from_ and email.from_.email_address else 'Unknown',
            'to': [recipient.email_address.address for recipient in (email.to_recipients or []) if recipient.email_address],
            'cc': [recipient.email_address.address for recipient in (email.cc_recipients or []) if recipient.email_address],
            'received_datetime': email.received_date_time.strftime("%Y-%m-%d %H:%M:%S") if email.received_date_time else 'Unknown',
            'importance': email.importance.value if email.importance else 'normal',
            'is_read': email.is_read if email.is_read is not None else False,
            'has_attachments': email.has_attachments if email.has_attachments is not None else False,
            'body': {
                'content': email.body.content if email.body and email.body.content else ''
            },
            'attachments': []
        }
        formatted = format_email(email_dict)
        parsed_emails.append(formatted)

    return {
        "emails": parsed_emails
    }

async def generate_email_tasks(state: GraphState, config: RunnableConfig) -> dict:

    emails = state.get("emails")
    todo_entries = state.get("todo_entries", [])
    email_tasks = state.get("email_tasks", [])

    model = ChatOpenAI(openai_api_key=get_settings().OPENAI_API_KEY, 
                       model=MODEL_CONFIGS["todo_list_generation"]["model"],
                       temperature=MODEL_CONFIGS["todo_list_generation"]["temperature"],
                       max_tokens=MODEL_CONFIGS["todo_list_generation"]["max_tokens"])

    model = model.with_structured_output(TodoList)

    content = emails.pop(0) if emails else ""

    try:
        model_answer = await model.ainvoke(
            [HumanMessage(content=todo_list_generation_prompt.format(
                input=content,
            ))],
            config
        )
        
        if model_answer is None:
            raise Exception("Model returned None")
            
    except Exception as e:
        l.error(f"Todo list generation failed: {str(e)}")
        raise
    try:    
        result = model_answer.model_dump()

        email_tasks = email_tasks + result.get("entries", [])

        return {
            "email_tasks": email_tasks, "emails": emails
        }
    except Exception as e:
        l.error(f"Todo list generation failed: {str(e)}")
        return {}
    
def check_emails(state) -> str:
    emails = state.get("emails", [])

    if len(emails) > 0:
        return "generate_email_tasks"
    else:
        return "get_teams_messages"

async def get_teams_messages(state: GraphState, config: RunnableConfig) -> dict:
    """Get teams messages from state (already fetched by the command)"""
    raw_teams_messages = state.get("raw_teams_messages", [])
    
    parsed_teams_messages = []
    for msg_data in raw_teams_messages:
        formatted = format_teams_message(msg_data)
        parsed_teams_messages.append(formatted)
    
    return {
        "teams_messages": parsed_teams_messages
    }

async def generate_teams_tasks(state: GraphState, config: RunnableConfig) -> dict:
    teams_messages = state.get("teams_messages")
    teams_tasks = state.get("teams_tasks", [])

    model = ChatOpenAI(
        openai_api_key=get_settings().OPENAI_API_KEY,
        model=MODEL_CONFIGS["todo_list_generation"]["model"],
        temperature=MODEL_CONFIGS["todo_list_generation"]["temperature"],
        max_tokens=MODEL_CONFIGS["todo_list_generation"]["max_tokens"]
    )

    model = model.with_structured_output(TodoList)

    content = teams_messages.pop(0) if teams_messages else ""
    
    try:
        model_answer = await model.ainvoke(
            [HumanMessage(content=todo_list_generation_prompt.format(
                input=content,
            ))],
            config
        )
        if model_answer is None:
            raise Exception("Model returned None")
        
    except Exception as e:
        l.error(f"Teams todo list generation failed: {str(e)}")
        raise
    try:
        result = model_answer.model_dump()
        teams_tasks = teams_tasks + result.get("entries", [])

        return {
            "teams_tasks": teams_tasks,
            "teams_messages": teams_messages
        }
    except Exception as e:
        l.error(f"Teams todo list generation failed: {str(e)}")
        return {}


def check_teams_messages(state) -> str:
    teams_messages = state.get("teams_messages", [])
    if len(teams_messages) > 0:
        return "generate_teams_tasks"
    else:
        return END

def create_graph():
    """
    Create and configure the todo list generation workflow graph.
    
    Returns:
        Compiled LangGraph workflow ready for execution
    """
    graph_builder = StateGraph(GraphState)
    
    graph_builder.add_node("generate_email_tasks", generate_email_tasks)
    graph_builder.add_node("get_emails", get_emails)
    graph_builder.add_node("generate_teams_tasks", generate_teams_tasks)
    graph_builder.add_node("get_teams_messages", get_teams_messages)
    

    graph_builder.add_edge(START, "get_emails")
    graph_builder.add_conditional_edges("get_emails", check_emails)
    graph_builder.add_conditional_edges("generate_email_tasks", check_emails)

    graph_builder.add_conditional_edges("get_teams_messages", check_teams_messages)
    graph_builder.add_conditional_edges("generate_teams_tasks", check_teams_messages)
    
    return graph_builder.compile(name="Todo List Generation Workflow")

async def run_generate_todo_list(
                                raw_emails: list = None,
                                raw_teams_messages: list = None,
                                user_id: str = "",
                                session_id: str = "") -> dict:
    """
    Execute the complete todo list generation workflow.
    
    Args:
        raw_emails: List of raw email objects from Microsoft Graph
        raw_teams_messages: List of raw Teams message objects
        user_id: User identifier for tracking (optional)
        session_id: Session identifier for tracking (optional)
        
    Returns:
        Dictionary containing:
        - email_tasks: List of TodoListEntry objects from emails
        - teams_tasks: List of TodoListEntry objects from Teams messages
    """
    graph = create_graph()
    
    try:
        result = await graph.ainvoke(
            {
                "content": "",
                "raw_emails": raw_emails or [],
                "raw_teams_messages": raw_teams_messages or [],
                "emails": [],
                "email_tasks": [],
                "teams_messages": [],
                "teams_tasks": [],
            },
            config={"recursion_limit": 200}
        )

        return {
            "email_tasks": result.get("email_tasks", []),
            "teams_tasks": result.get("teams_tasks", []),
            "all_tasks": result.get("email_tasks", []) + result.get("teams_tasks", [])
        }

    except Exception as e:
        l.error(f"Todo list generation workflow failed: {str(e)} | user: {user_id} | session: {session_id}")
        raise



class SingleTaskState(TypedDict):
    """State for single task generation from user message"""
    user_message: str
    task_entry: dict


async def generate_task_from_message(state: SingleTaskState, config: RunnableConfig) -> dict:
    """Generate a detailed task from user message in Portuguese"""
    user_message = state.get("user_message", "")
    
    if not user_message:
        return {"task_entry": {}}
    
    model = ChatOpenAI(
        openai_api_key=get_settings().OPENAI_API_KEY,
        model=MODEL_CONFIGS["todo_list_generation"]["model"],
        temperature=MODEL_CONFIGS["todo_list_generation"]["temperature"],
        max_tokens=MODEL_CONFIGS["todo_list_generation"]["max_tokens"]
    )
    
    model = model.with_structured_output(TodoList)
    
    try:
        model_answer = await model.ainvoke(
            [HumanMessage(content=single_task_generation_prompt.format(
                user_message=user_message,
            ))],
            config
        )
        
        if model_answer is None:
            raise Exception("Model returned None")
        
        result = model_answer.model_dump()
        entries = result.get("entries", [])
        
        # Get the first (and should be only) task entry
        task_entry = entries[0] if entries else {}
        
        return {"task_entry": task_entry}
        
    except Exception as e:
        l.error(f"Single task generation failed: {str(e)}")
        raise


def create_single_task_graph():
    """Create a simple graph for single task generation"""
    graph_builder = StateGraph(SingleTaskState)
    
    graph_builder.add_node("generate_task", generate_task_from_message)
    
    graph_builder.add_edge(START, "generate_task")
    graph_builder.add_edge("generate_task", END)
    
    return graph_builder.compile(name="Single Task Generation Workflow")


async def generate_single_task_from_user_message(user_message: str) -> dict:
    """
    Generate a single detailed task from a user message in Portuguese.
    
    Args:
        user_message: The user's message describing what they want to do
        
    Returns:
        Dictionary containing the task entry with fields:
        - task: The detailed task title
        - priority: Priority level (high/normal/low)
        - comments: Additional context or steps
        - due_date: Optional due date
    """
    graph = create_single_task_graph()
    
    try:
        result = await graph.ainvoke(
            {
                "user_message": user_message,
                "task_entry": {},
            },
            config={"recursion_limit": 10}
        )
        
        return result.get("task_entry", {})
        
    except Exception as e:
        l.error(f"Single task generation workflow failed: {str(e)}")
        raise


class DeduplicationState(TypedDict):
    """State for task deduplication workflow"""
    incomplete_tasks: list  # Existing incomplete tasks from Microsoft To Do
    new_tasks: list  # New tasks generated from emails/teams
    deduplicated_tasks: list  # Tasks that don't exist yet
    duplicate_tasks: list  # Tasks that are duplicates
    graph_client: object  # Microsoft Graph client
    target_list_id: str  # ID of the To Do list to add tasks to


async def fetch_incomplete_tasks(state: DeduplicationState, config: RunnableConfig) -> dict:
    """Fetch all incomplete tasks from Microsoft To Do"""
    from graph_api.todo_tasks import get_all_incomplete_tasks
    
    graph_client = state.get("graph_client")
    
    if not graph_client:
        l.error("No graph client provided")
        return {"incomplete_tasks": []}
    
    try:
        result = await get_all_incomplete_tasks(graph_client)
        
        if result['success']:
            # Flatten the tasks for easier comparison
            incomplete_tasks = []
            for task_data in result['all_tasks']:
                task = task_data['task']
                incomplete_tasks.append({
                    'title': task.title or '',
                    'body': task.body.content if task.body and task.body.content else '',
                    'due_date': task.due_date_time.date_time if task.due_date_time else None,
                    'importance': str(task.importance) if task.importance else 'normal',
                    'list_name': task_data['list_name'],
                    'list_id': task_data['list_id'],
                    'task_id': task.id
                })
            
            return {"incomplete_tasks": incomplete_tasks}
        else:
            l.error(f"Failed to fetch incomplete tasks: {result['error']}")
            return {"incomplete_tasks": []}
            
    except Exception as e:
        l.error(f"Error fetching incomplete tasks: {str(e)}")
        return {"incomplete_tasks": []}


async def deduplicate_tasks(state: DeduplicationState, config: RunnableConfig) -> dict:
    """Use AI to identify which new tasks are duplicates of existing ones"""
    incomplete_tasks = state.get("incomplete_tasks", [])
    new_tasks = state.get("new_tasks", [])
    
    if not new_tasks:
        return {"deduplicated_tasks": [], "duplicate_tasks": []}
    
    if not incomplete_tasks:
        # No existing tasks, all new tasks are unique
        return {"deduplicated_tasks": new_tasks, "duplicate_tasks": []}
    
    # Prepare context for the AI
    existing_tasks_text = "EXISTING INCOMPLETE TASKS:\n"
    for i, task in enumerate(incomplete_tasks, 1):
        existing_tasks_text += f"{i}. {task['title']}\n"
        if task['body']:
            existing_tasks_text += f"   Details: {task['body'][:200]}\n"
        if task['due_date']:
            existing_tasks_text += f"   Due: {task['due_date']}\n"
    
    new_tasks_text = "\nNEW TASKS TO EVALUATE:\n"
    for i, task in enumerate(new_tasks, 1):
        new_tasks_text += f"{i}. {task.get('task', '')}\n"
        if task.get('comments'):
            new_tasks_text += f"   Details: {task.get('comments', '')}\n"
        if task.get('due_date'):
            new_tasks_text += f"   Due: {task.get('due_date', '')}\n"
    
    deduplication_prompt = f"""
    Você é um especialista em análise de tarefas. Sua função é identificar se as novas tarefas já existem na lista de tarefas incompletas.
    
    INSTRUÇÕES:
    - Compare cada nova tarefa com as tarefas existentes
    - Considere uma tarefa como DUPLICADA se:
      * O objetivo principal é o mesmo (mesmo que a redação seja diferente)
      * Refere-se ao mesmo assunto ou ação
      * Tem contexto ou prazo similar
    - Considere uma tarefa como ÚNICA se:
      * É uma ação diferente ou adicional
      * Refere-se a um aspecto distinto do trabalho
      * É um follow-up ou próximo passo de uma tarefa existente
    
    {existing_tasks_text}
    {new_tasks_text}
    
    Para cada nova tarefa, retorne:
    - "unique": se a tarefa NÃO existe na lista atual
    - "duplicate": se a tarefa JÁ existe na lista atual
    
    Retorne uma lista com o status de cada nova tarefa no formato JSON:
    {{"task_status": [{{"task_number": 1, "status": "unique"}}, {{"task_number": 2, "status": "duplicate"}}]}}
    """
    
    model = ChatOpenAI(
        openai_api_key=get_settings().OPENAI_API_KEY,
        model="gpt-4o",
        temperature=0.0
    )
    
    try:
        from pydantic import BaseModel
        
        class TaskStatus(BaseModel):
            task_number: int
            status: str  # "unique" or "duplicate"
        
        class TaskStatusList(BaseModel):
            task_status: list[TaskStatus]
        
        model = model.with_structured_output(TaskStatusList)
        
        model_answer = await model.ainvoke(
            [HumanMessage(content=deduplication_prompt)],
            config
        )
        
        result = model_answer.model_dump()
        task_status_list = result.get("task_status", [])
        
        deduplicated_tasks = []
        duplicate_tasks = []
        
        for status_item in task_status_list:
            task_number = status_item['task_number']
            status = status_item['status']
            
            if 1 <= task_number <= len(new_tasks):
                task = new_tasks[task_number - 1]
                if status == "unique":
                    deduplicated_tasks.append(task)
                else:
                    duplicate_tasks.append(task)
        
        l.info(f"Deduplication complete: {len(deduplicated_tasks)} unique, {len(duplicate_tasks)} duplicates")
        
        return {
            "deduplicated_tasks": deduplicated_tasks,
            "duplicate_tasks": duplicate_tasks
        }
        
    except Exception as e:
        l.error(f"Error during deduplication: {str(e)}")
        # If deduplication fails, treat all tasks as unique to be safe
        return {
            "deduplicated_tasks": new_tasks,
            "duplicate_tasks": []
        }


async def create_new_tasks(state: DeduplicationState, config: RunnableConfig) -> dict:
    """Create the deduplicated tasks in Microsoft To Do"""
    from graph_api.todo_tasks import create_task, get_todo_lists
    
    graph_client = state.get("graph_client")
    deduplicated_tasks = state.get("deduplicated_tasks", [])
    target_list_id = state.get("target_list_id")
    
    if not graph_client:
        l.error("No graph client provided")
        return {"created_tasks": [], "creation_errors": []}
    
    if not deduplicated_tasks:
        l.info("No tasks to create")
        return {"created_tasks": [], "creation_errors": []}
    
    # If no target list specified, try to find or create a default one
    if not target_list_id:
        lists_result = await get_todo_lists(graph_client)
        if lists_result['success'] and lists_result['lists']:
            # Use the first list (usually the default "Tasks" list)
            target_list_id = lists_result['lists'][0].id
            l.info(f"Using default list: {lists_result['lists'][0].display_name}")
        else:
            l.error("No To Do lists available and no target list specified")
            return {"created_tasks": [], "creation_errors": ["No target list available"]}
    
    created_tasks = []
    creation_errors = []
    
    for task_data in deduplicated_tasks:
        try:
            # Map priority from our format to Microsoft To Do format
            priority = task_data.get('priority', 'neutral').lower()
            if 'high' in priority:
                importance = 'high'
            elif 'low' in priority:
                importance = 'low'
            else:
                importance = 'normal'
            
            # Format body with comments and person involved
            task_body = format_task_body(
                comments=task_data.get('comments'),
                person_envolved=task_data.get('person_envolved')
            )
            
            result = await create_task(
                graph=graph_client,
                list_id=target_list_id,
                title=task_data.get('task', 'Untitled Task'),
                body=task_body,
                due_date=task_data.get('due_date'),
                importance=importance
            )
            
            if result['success']:
                created_tasks.append({
                    'title': task_data.get('task'),
                    'task_id': result['task'].id,
                    'list_id': target_list_id
                })
                l.info(f"Created task: {task_data.get('task')}")
            else:
                creation_errors.append({
                    'task': task_data.get('task'),
                    'error': result['error']
                })
                l.error(f"Failed to create task: {result['error']}")
                
        except Exception as e:
            creation_errors.append({
                'task': task_data.get('task', 'Unknown'),
                'error': str(e)
            })
            l.error(f"Error creating task: {str(e)}")
    
    return {
        "created_tasks": created_tasks,
        "creation_errors": creation_errors
    }


def create_upsert_graph():
    """Create the graph for upserting tasks with deduplication"""
    graph_builder = StateGraph(DeduplicationState)
    
    graph_builder.add_node("fetch_incomplete_tasks", fetch_incomplete_tasks)
    graph_builder.add_node("deduplicate_tasks", deduplicate_tasks)
    graph_builder.add_node("create_new_tasks", create_new_tasks)
    
    graph_builder.add_edge(START, "fetch_incomplete_tasks")
    graph_builder.add_edge("fetch_incomplete_tasks", "deduplicate_tasks")
    graph_builder.add_edge("deduplicate_tasks", "create_new_tasks")
    graph_builder.add_edge("create_new_tasks", END)
    
    return graph_builder.compile(name="Task Upsert with Deduplication Workflow")


async def upsert_new_tasks(
    graph_client,
    new_tasks: list,
    target_list_id: str = None
) -> dict:
    """
    Insere novas tarefas verificando tarefas incompletas existentes e criando apenas tarefas únicas.
    
    Args:
        graph_client: Instância do cliente Microsoft Graph
        new_tasks: Lista de dicionários de tarefas (de email_tasks, teams_tasks, etc.)
                   Cada tarefa deve ter: task, priority, comments, due_date
        target_list_id: Opcional - ID da lista To Do para adicionar as tarefas.
                        Se None, usa a lista padrão.
    
    Returns:
        Dicionário contendo:
        - success: bool
        - created_tasks: Lista de tarefas criadas com sucesso
        - duplicate_tasks: Lista de tarefas que eram duplicadas
        - creation_errors: Lista de erros na criação de tarefas
        - summary: Resumo em linguagem natural
    """
    graph = create_upsert_graph()
    
    try:
        result = await graph.ainvoke(
            {
                "incomplete_tasks": [],
                "new_tasks": new_tasks,
                "deduplicated_tasks": [],
                "duplicate_tasks": [],
                "graph_client": graph_client,
                "target_list_id": target_list_id,
                "created_tasks": [],
                "creation_errors": []
            },
            config={"recursion_limit": 50}
        )
        
        created_tasks = result.get("created_tasks", [])
        duplicate_tasks = result.get("duplicate_tasks", [])
        creation_errors = result.get("creation_errors", [])
        
        summary = f"""
✅ Inserção de Tarefas Concluída:
- {len(created_tasks)} novas tarefas criadas
- {len(duplicate_tasks)} duplicatas ignoradas
- {len(creation_errors)} erros
        """.strip()
        
        return {
            "success": True,
            "created_tasks": created_tasks,
            "duplicate_tasks": duplicate_tasks,
            "creation_errors": creation_errors,
            "summary": summary
        }
        
    except Exception as e:
        l.error(f"Falha no fluxo de inserção de tarefas: {str(e)}")
        return {
            "success": False,
            "created_tasks": [],
            "duplicate_tasks": [],
            "creation_errors": [str(e)],
            "summary": f"❌ Falha na inserção de tarefas: {str(e)}"
        }


async def complete_task_generation_workflow(
    graph_client,
    raw_emails: list = None,
    raw_teams_messages: list = None,
    target_list_id: str = None,
    user_id: str = "",
    session_id: str = ""
) -> dict:
    """
    Fluxo de trabalho completo de ponta a ponta:
    1. Gerar tarefas a partir de emails e mensagens do Teams
    2. Buscar tarefas incompletas existentes
    3. Deduplicar novas tarefas com as existentes
    4. Criar apenas tarefas únicas no Microsoft To Do
    
    Esta é a principal função para o fluxo completo.
    
    Args:
        graph_client: Instância do cliente Microsoft Graph
        raw_emails: Lista de objetos de email brutos do Microsoft Graph
        raw_teams_messages: Lista de objetos de mensagens Teams brutas
        target_list_id: Opcional - ID da lista To Do para adicionar tarefas
        user_id: Identificador do usuário para rastreamento (opcional)
        session_id: Identificador da sessão para rastreamento (opcional)
    
    Returns:
        Dicionário contendo:
        - success: bool
        - email_tasks: Lista de tarefas geradas de emails
        - teams_tasks: Lista de tarefas geradas de mensagens do Teams
        - all_generated_tasks: Lista combinada de todas as tarefas geradas
        - created_tasks: Lista de tarefas criadas com sucesso
        - duplicate_tasks: Lista de tarefas que eram duplicadas
        - creation_errors: Lista de erros na criação de tarefas
        - summary: Resumo em linguagem natural
    """
    try:
        # Etapa 1: Gerar tarefas a partir de emails e mensagens do Teams
        l.info("Etapa 1: Gerando tarefas a partir de emails e mensagens do Teams...")
        generation_result = await run_generate_todo_list(
            raw_emails=raw_emails,
            raw_teams_messages=raw_teams_messages,
            user_id=user_id,
            session_id=session_id
        )
        
        email_tasks = generation_result.get("email_tasks", [])
        teams_tasks = generation_result.get("teams_tasks", [])
        all_generated_tasks = generation_result.get("all_tasks", [])
        
        l.info(f"{len(email_tasks)} tarefas geradas a partir de emails, {len(teams_tasks)} tarefas a partir do Teams")
        
        if not all_generated_tasks:
            return {
                "success": True,
                "email_tasks": [],
                "teams_tasks": [],
                "all_generated_tasks": [],
                "created_tasks": [],
                "duplicate_tasks": [],
                "creation_errors": [],
                "summary": "ℹ️ Nenhuma tarefa encontrada em emails ou mensagens do Teams"
            }
        
        # Etapa 2: Inserir tarefas (com deduplicação)
        l.info("Etapa 2: Removendo duplicatas e criando tarefas...")
        upsert_result = await upsert_new_tasks(
            graph_client=graph_client,
            new_tasks=all_generated_tasks,
            target_list_id=target_list_id
        )
        
        # Combina resultados
        return {
            "success": True,
            "email_tasks": email_tasks,
            "teams_tasks": teams_tasks,
            "all_generated_tasks": all_generated_tasks,
            "created_tasks": upsert_result.get("created_tasks", []),
            "duplicate_tasks": upsert_result.get("duplicate_tasks", []),
            "creation_errors": upsert_result.get("creation_errors", []),
            "summary": f"""
📊 Resumo Completo do Fluxo:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━
📧 Tarefas de Email Geradas: {len(email_tasks)}
💬 Tarefas do Teams Geradas: {len(teams_tasks)}
📝 Total de Tarefas Geradas: {len(all_generated_tasks)}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━
✅ Novas Tarefas Criadas: {len(upsert_result.get('created_tasks', []))}
🔄 Duplicatas Ignoradas: {len(upsert_result.get('duplicate_tasks', []))}
❌ Erros: {len(upsert_result.get('creation_errors', []))}
            """.strip()
        }
        
    except Exception as e:
        l.error(f"Falha no fluxo completo: {str(e)} | usuário: {user_id} | sessão: {session_id}")
        return {
            "success": False,
            "email_tasks": [],
            "teams_tasks": [],
            "all_generated_tasks": [],
            "created_tasks": [],
            "duplicate_tasks": [],
            "creation_errors": [str(e)],
            "summary": f"❌ Falha no fluxo: {str(e)}"
        }