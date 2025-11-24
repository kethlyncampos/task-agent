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