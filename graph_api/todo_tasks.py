"""
Module for managing Microsoft To Do tasks using Graph SDK.
"""

from typing import Dict, Any, Optional, List
from datetime import datetime
from msgraph.generated.users.item.todo.lists.lists_request_builder import ListsRequestBuilder
from msgraph.generated.users.item.todo.lists.item.tasks.tasks_request_builder import TasksRequestBuilder
from kiota_abstractions.base_request_configuration import RequestConfiguration
from msgraph.generated.models.todo_task import TodoTask
from msgraph.generated.models.todo_task_list import TodoTaskList
from msgraph.generated.models.item_body import ItemBody
from msgraph.generated.models.body_type import BodyType
from msgraph.generated.models.importance import Importance
from msgraph.generated.models.date_time_time_zone import DateTimeTimeZone


async def get_todo_lists(graph):
    """
    Fetch all To Do task lists for the current user.
    
    Parameters:
        graph: The Microsoft Graph client instance
    
    Returns:
        dict: {
            'success': bool,
            'lists': list,  # List of task list objects
            'error': str or None
        }
    """
    try:
        # Get all task lists
        lists_response = await graph.me.todo.lists.get()
        
        # Extract lists from response
        lists = lists_response.value if lists_response and lists_response.value else []
        
        return {
            'success': True,
            'lists': lists,
            'error': None
        }
        
    except Exception as e:
        return {
            'success': False,
            'lists': [],
            'error': f"Error fetching To Do lists: {str(e)}"
        }


async def get_tasks_from_list(graph, list_id: str):
    """
    Fetch all tasks from a specific To Do list.
    
    Parameters:
        graph: The Microsoft Graph client instance
        list_id (str): The ID of the task list
    
    Returns:
        dict: {
            'success': bool,
            'tasks': list,  # List of task objects
            'error': str or None
        }
    """
    try:
        # Create request configuration
        query_params = TasksRequestBuilder.TasksRequestBuilderGetQueryParameters()
        query_params.orderby = ["createdDateTime DESC"]
        
        request_config = RequestConfiguration(query_parameters=query_params)
        
        # Get tasks from the list
        tasks_response = await graph.me.todo.lists.by_todo_task_list_id(list_id).tasks.get(
            request_configuration=request_config
        )
        
        # Extract tasks from response
        tasks = tasks_response.value if tasks_response and tasks_response.value else []
        
        return {
            'success': True,
            'tasks': tasks,
            'error': None
        }
        
    except Exception as e:
        return {
            'success': False,
            'tasks': [],
            'error': f"Error fetching tasks: {str(e)}"
        }


async def get_incomplete_tasks_from_list(graph, list_id: str):
    """
    Fetch all incomplete (not completed) tasks from a specific To Do list.
    
    Parameters:
        graph: The Microsoft Graph client instance
        list_id (str): The ID of the task list
    
    Returns:
        dict: {
            'success': bool,
            'tasks': list,  # List of incomplete task objects
            'error': str or None
        }
    """
    try:
        # Create request configuration with filter for incomplete tasks
        query_params = TasksRequestBuilder.TasksRequestBuilderGetQueryParameters()
        query_params.filter = "status ne 'completed'"
        query_params.orderby = ["createdDateTime DESC"]
        
        request_config = RequestConfiguration(query_parameters=query_params)
        
        # Get incomplete tasks from the list
        tasks_response = await graph.me.todo.lists.by_todo_task_list_id(list_id).tasks.get(
            request_configuration=request_config
        )
        
        # Extract tasks from response
        tasks = tasks_response.value if tasks_response and tasks_response.value else []
        
        return {
            'success': True,
            'tasks': tasks,
            'count': len(tasks),
            'error': None
        }
        
    except Exception as e:
        return {
            'success': False,
            'tasks': [],
            'count': 0,
            'error': f"Error fetching incomplete tasks: {str(e)}"
        }


async def get_all_incomplete_tasks(graph):
    """
    Fetch all incomplete (not completed) tasks from all To Do lists.
    
    Parameters:
        graph: The Microsoft Graph client instance
    
    Returns:
        dict: {
            'success': bool,
            'tasks_by_list': dict,  # Dictionary mapping list_name -> list of tasks
            'all_tasks': list,  # Flat list of all incomplete tasks
            'total_count': int,  # Total number of incomplete tasks
            'error': str or None
        }
    """
    try:
        tasks_by_list = {}
        all_tasks = []
        
        # First, get all task lists
        lists_result = await get_todo_lists(graph)
        
        if not lists_result['success']:
            return {
                'success': False,
                'tasks_by_list': {},
                'all_tasks': [],
                'total_count': 0,
                'error': lists_result['error']
            }
        
        # For each list, get incomplete tasks
        for task_list in lists_result['lists']:
            list_id = task_list.id
            list_name = task_list.display_name or "(Unnamed list)"
            
            # Get incomplete tasks from this list
            tasks_result = await get_incomplete_tasks_from_list(graph, list_id)
            
            if tasks_result['success'] and tasks_result['tasks']:
                # Store tasks organized by list
                tasks_by_list[list_name] = {
                    'list_id': list_id,
                    'tasks': tasks_result['tasks']
                }
                
                # Add to flat list with list name info
                for task in tasks_result['tasks']:
                    task_with_list = {
                        'list_name': list_name,
                        'list_id': list_id,
                        'task': task
                    }
                    all_tasks.append(task_with_list)
        
        return {
            'success': True,
            'tasks_by_list': tasks_by_list,
            'all_tasks': all_tasks,
            'total_count': len(all_tasks),
            'error': None
        }
        
    except Exception as e:
        return {
            'success': False,
            'tasks_by_list': {},
            'all_tasks': [],
            'total_count': 0,
            'error': f"Error fetching all incomplete tasks: {str(e)}"
        }


async def create_task(graph, list_id: str, title: str, body: str = None, 
                     due_date: str = None, importance: str = "normal"):
    """
    Create a new task in a To Do list.
    
    Parameters:
        graph: The Microsoft Graph client instance
        list_id (str): The ID of the task list
        title (str): Task title
        body (str, optional): Task description/notes
        due_date (str, optional): Due date in ISO 8601 format (e.g., '2024-12-31')
        importance (str, optional): Task importance ('low', 'normal', 'high')
    
    Returns:
        dict: {
            'success': bool,
            'task': object or None,  # Created task object
            'error': str or None
        }
    """
    try:
        # Create new task object
        new_task = TodoTask()
        new_task.title = title
        
        # Add body/description if provided
        if body:
            task_body = ItemBody()
            task_body.content_type = BodyType.Text
            task_body.content = body
            new_task.body = task_body
        
        # Add due date if provided
        if due_date:
            due_datetime = DateTimeTimeZone()
            due_datetime.date_time = due_date if 'T' in due_date else f"{due_date}T00:00:00"
            due_datetime.time_zone = "UTC"
            new_task.due_date_time = due_datetime
        
        # Set importance
        if importance.lower() == "high":
            new_task.importance = Importance.High
        elif importance.lower() == "low":
            new_task.importance = Importance.Low
        else:
            new_task.importance = Importance.Normal
        
        # Create the task
        created_task = await graph.me.todo.lists.by_todo_task_list_id(list_id).tasks.post(new_task)
        
        return {
            'success': True,
            'task': created_task,
            'error': None
        }
        
    except Exception as e:
        return {
            'success': False,
            'task': None,
            'error': f"Error creating task: {str(e)}"
        }


async def complete_task(graph, list_id: str, task_id: str):
    """
    Mark a task as completed.
    
    Parameters:
        graph: The Microsoft Graph client instance
        list_id (str): The ID of the task list
        task_id (str): The ID of the task to complete
    
    Returns:
        dict: {
            'success': bool,
            'task': object or None,  # Updated task object
            'error': str or None
        }
    """
    try:
        # Create task update with completed status
        task_update = TodoTask()
        task_update.status = "completed"
        
        # Update the task
        updated_task = await graph.me.todo.lists.by_todo_task_list_id(list_id).tasks.by_todo_task_id(task_id).patch(task_update)
        
        return {
            'success': True,
            'task': updated_task,
            'error': None
        }
        
    except Exception as e:
        return {
            'success': False,
            'task': None,
            'error': f"Error completing task: {str(e)}"
        }


async def update_task(graph, list_id: str, task_id: str, title: str = None, 
                     body: str = None, due_date: str = None, importance: str = None):
    """
    Update an existing task.
    
    Parameters:
        graph: The Microsoft Graph client instance
        list_id (str): The ID of the task list
        task_id (str): The ID of the task to update
        title (str, optional): New task title
        body (str, optional): New task description
        due_date (str, optional): New due date in ISO 8601 format
        importance (str, optional): New importance ('low', 'normal', 'high')
    
    Returns:
        dict: {
            'success': bool,
            'task': object or None,  # Updated task object
            'error': str or None
        }
    """
    try:
        # Create task update object
        task_update = TodoTask()
        
        if title:
            task_update.title = title
        
        if body:
            task_body = ItemBody()
            task_body.content_type = BodyType.Text
            task_body.content = body
            task_update.body = task_body
        
        if due_date:
            due_datetime = DateTimeTimeZone()
            due_datetime.date_time = due_date if 'T' in due_date else f"{due_date}T00:00:00"
            due_datetime.time_zone = "UTC"
            task_update.due_date_time = due_datetime
        
        if importance:
            if importance.lower() == "high":
                task_update.importance = Importance.High
            elif importance.lower() == "low":
                task_update.importance = Importance.Low
            else:
                task_update.importance = Importance.Normal
        
        # Update the task
        updated_task = await graph.me.todo.lists.by_todo_task_list_id(list_id).tasks.by_todo_task_id(task_id).patch(task_update)
        
        return {
            'success': True,
            'task': updated_task,
            'error': None
        }
        
    except Exception as e:
        return {
            'success': False,
            'task': None,
            'error': f"Error updating task: {str(e)}"
        }


async def delete_task(graph, list_id: str, task_id: str):
    """
    Delete a task from a To Do list.
    
    Parameters:
        graph: The Microsoft Graph client instance
        list_id (str): The ID of the task list
        task_id (str): The ID of the task to delete
    
    Returns:
        dict: {
            'success': bool,
            'error': str or None
        }
    """
    try:
        # Delete the task
        await graph.me.todo.lists.by_todo_task_list_id(list_id).tasks.by_todo_task_id(task_id).delete()
        
        return {
            'success': True,
            'error': None
        }
        
    except Exception as e:
        return {
            'success': False,
            'error': f"Error deleting task: {str(e)}"
        }


async def create_task_list(graph, list_name: str):
    """
    Create a new To Do task list.
    
    Parameters:
        graph: The Microsoft Graph client instance
        list_name (str): Name of the new task list
    
    Returns:
        dict: {
            'success': bool,
            'list': object or None,  # Created list object
            'error': str or None
        }
    """
    try:
        # Create new task list object
        new_list = TodoTaskList()
        new_list.display_name = list_name
        
        # Create the list
        created_list = await graph.me.todo.lists.post(new_list)
        
        return {
            'success': True,
            'list': created_list,
            'error': None
        }
        
    except Exception as e:
        return {
            'success': False,
            'list': None,
            'error': f"Error creating task list: {str(e)}"
        }


def format_task(task) -> str:
    """
    Format a To Do task into a readable string.
    
    Args:
        task: Task object from Graph SDK
    
    Returns:
        str: Formatted task string
    """
    try:
        lines = []
        
        # Task title
        title = task.title or "(No title)"
        status_icon = "✅" if task.status == "completed" else "⬜"
        lines.append(f"{status_icon} **{title}**")
        
        # Importance
        if task.importance and str(task.importance).lower() != "normal":
            importance = str(task.importance).replace("Importance.", "")
            lines.append(f"   Priority: {importance}")
        
        # Due date
        if task.due_date_time:
            due_date = task.due_date_time.date_time
            if due_date:
                # Parse and format the date
                try:
                    if isinstance(due_date, str):
                        dt = datetime.fromisoformat(due_date.replace('Z', '+00:00'))
                        lines.append(f"   Due: {dt.strftime('%Y-%m-%d')}")
                    else:
                        lines.append(f"   Due: {due_date}")
                except:
                    lines.append(f"   Due: {due_date}")
        
        # Body/description
        if task.body and task.body.content:
            content = task.body.content.strip()
            if content and len(content) > 0:
                # Truncate if too long
                if len(content) > 100:
                    content = content[:100] + "..."
                lines.append(f"   Note: {content}")
        
        # Created date
        if task.created_date_time:
            created = task.created_date_time.strftime("%Y-%m-%d")
            lines.append(f"   Created: {created}")
        
        return "\n".join(lines)
        
    except Exception as e:
        return f"Error formatting task: {str(e)}"


def format_task_list(task_list) -> str:
    """
    Format a To Do task list into a readable string.
    
    Args:
        task_list: TaskList object from Graph SDK
    
    Returns:
        str: Formatted list string
    """
    try:
        name = task_list.display_name or "(Unnamed list)"
        is_owner = task_list.is_owner if hasattr(task_list, 'is_owner') else True
        owner_text = " (Owner)" if is_owner else " (Shared)"
        
        return f"📋 **{name}**{owner_text}"
        
    except Exception as e:
        return f"Error formatting list: {str(e)}"


def format_incomplete_tasks_summary(tasks_by_list: Dict[str, Any]) -> str:
    """
    Format incomplete tasks grouped by list into a readable summary.
    
    Args:
        tasks_by_list: Dictionary from get_all_incomplete_tasks()
    
    Returns:
        str: Formatted summary string
    """
    if not tasks_by_list:
        return "✅ No incomplete tasks found! Great job!"
    
    lines = ["📝 **Incomplete Tasks:**\n"]
    total_tasks = 0
    
    for list_name, list_data in tasks_by_list.items():
        tasks = list_data['tasks']
        task_count = len(tasks)
        total_tasks += task_count
        
        lines.append(f"📋 **{list_name}** ({task_count} task{'s' if task_count != 1 else ''})")
        
        for i, task in enumerate(tasks[:10], 1):  # Show first 10 tasks per list
            title = task.title or "(No title)"
            
            # Add importance indicator
            importance_icon = ""
            if task.importance:
                importance_str = str(task.importance).lower()
                if "high" in importance_str:
                    importance_icon = "❗"
                elif "low" in importance_str:
                    importance_icon = "🔽"
            
            # Add due date if available
            due_info = ""
            if task.due_date_time and task.due_date_time.date_time:
                try:
                    due_date = task.due_date_time.date_time
                    if isinstance(due_date, str):
                        dt = datetime.fromisoformat(due_date.replace('Z', '+00:00'))
                        # Check if overdue
                        if dt.date() < datetime.now().date():
                            due_info = f" ⚠️ (Overdue: {dt.strftime('%Y-%m-%d')})"
                        else:
                            due_info = f" 📅 (Due: {dt.strftime('%Y-%m-%d')})"
                except:
                    pass
            
            lines.append(f"   {i}. {importance_icon}⬜ {title}{due_info}")
        
        if task_count > 10:
            lines.append(f"   ... and {task_count - 10} more task{'s' if task_count - 10 != 1 else ''}")
        
        lines.append("")  # Empty line between lists
    
    lines.insert(1, f"**Total: {total_tasks} incomplete task{'s' if total_tasks != 1 else ''}**\n")
    
    return "\n".join(lines)

