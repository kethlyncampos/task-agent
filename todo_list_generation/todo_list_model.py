from typing import List, Dict, Optional
from pydantic import BaseModel, Field

class TodoListEntry(BaseModel):
    task: str = Field(..., description="Task to be done")
    priority: str = Field(
        ...,
        description="Priority level of the task. Must be one of: 'high priority', 'low priority', or 'neutral'."
    )
    comments: Optional[str] = Field(..., description="Comments about the task")
    due_date: Optional[str] = Field(..., description="Due date of the task, if any")
    person_envolved: Optional[str] = Field(..., description="Person envolved in the task, if any in the format Name LastName")

class TodoList(BaseModel):
    entries: List[TodoListEntry] = Field(..., description="List of todo list entries")