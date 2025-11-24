from pathlib import Path
from docx import Document
from datetime import datetime, timedelta, timezone
from typing import Dict, Any, Optional, List
from kiota_abstractions.base_request_configuration import RequestConfiguration
from msgraph.generated.users.item.online_meetings.online_meetings_request_builder import OnlineMeetingsRequestBuilder
from msgraph.generated.communications.call_records.call_records_request_builder import CallRecordsRequestBuilder

def read_transcripts_fake(transcripts_folder: str = None) -> list[str]:
    """
    Reads all .docx files from the transcripts folder and returns their contents.
    
    Args:
        transcripts_folder: Path to the transcripts folder. 
                           If None, defaults to the 'transcripts' folder in the project root.
    
    Returns:
        A list of strings containing the text content of each .docx file.
    """
    if transcripts_folder is None:
        # Default to the transcripts folder in the project root
        transcripts_folder = Path(__file__).parent.parent / "transcripts"
    else:
        transcripts_folder = Path(transcripts_folder)
    
    contents = []
    
    # Get all .docx files in the transcripts folder
    docx_files = sorted(transcripts_folder.glob("*.docx"))
    
    for docx_file in docx_files:
        try:
            # Open and read the docx file
            doc = Document(docx_file)
            
            # Extract all paragraphs
            text_content = "\n".join([paragraph.text for paragraph in doc.paragraphs])
            contents.append(text_content)
            
        except Exception as e:
            print(f"Error reading {docx_file.name}: {e}")
    
    return contents


async def get_teams_call_recordings(graph, days: int = 2, target_user_email: Optional[str] = None) -> Dict[str, Any]:
    """
    Fetch Teams call recordings from recorded meetings using Microsoft Graph SDK.
    
    Parameters:
        graph: The Microsoft Graph client instance (from ctx.user_graph)
        days (int): Number of days to look back (default: 2)
        target_user_email (str, optional): Email of user to query. If None, uses current user (me)
    
    Returns:
        dict: {
            'success': bool,
            'recordings': list,
            'error': str or None
        }
    """
    
    try:
        recordings_list = []
        
        # Calculate date range
        start_date = (datetime.now(timezone.utc) - timedelta(days=days)).isoformat().replace('+00:00', 'Z')
        end_date = datetime.now(timezone.utc).isoformat().replace('+00:00', 'Z')
        
        # Step 1: Get online meetings from the last N days
        try:
            
            
            # Build filter for meetings in date range
            filter_query = f"creationDateTime ge {start_date} and creationDateTime le {end_date}"
            
            query_params = OnlineMeetingsRequestBuilder.OnlineMeetingsRequestBuilderGetQueryParameters()
            query_params.filter = filter_query
            
            request_config = RequestConfiguration(query_parameters=query_params)
            
            # Get meetings for the specified user or current user
            if target_user_email:
                meetings_response = await graph.users.by_user_id(target_user_email).online_meetings.get(
                    request_configuration=request_config
                )
            else:
                meetings_response = await graph.me.online_meetings.get(
                    request_configuration=request_config
                )
            
            meetings = meetings_response.value if meetings_response and meetings_response.value else []
            
            # Step 2: For each meeting, try to get recordings
            for meeting in meetings:
                meeting_id = meeting.id
                
                try:
                    # Get recordings for this meeting
                    if target_user_email:
                        recordings_response = await graph.users.by_user_id(target_user_email).online_meetings.by_online_meeting_id(meeting_id).recordings.get()
                    else:
                        recordings_response = await graph.me.online_meetings.by_online_meeting_id(meeting_id).recordings.get()
                    
                    if recordings_response and recordings_response.value:
                        for recording in recordings_response.value:
                            recording_info = {
                                'meeting_id': meeting_id,
                                'meeting_subject': meeting.subject if hasattr(meeting, 'subject') else 'N/A',
                                'meeting_start_time': meeting.start_date_time.isoformat() if hasattr(meeting, 'start_date_time') and meeting.start_date_time else 'N/A',
                                'meeting_end_time': meeting.end_date_time.isoformat() if hasattr(meeting, 'end_date_time') and meeting.end_date_time else 'N/A',
                                'recording_id': recording.id if hasattr(recording, 'id') else 'N/A',
                                'recording_created': recording.created_date_time.isoformat() if hasattr(recording, 'created_date_time') and recording.created_date_time else 'N/A',
                                'recording_content_url': recording.content if hasattr(recording, 'content') else None,
                                'recording_created_by': recording.created_by.user.display_name if hasattr(recording, 'created_by') and recording.created_by and hasattr(recording.created_by, 'user') else 'Unknown'
                            }
                            recordings_list.append(recording_info)
                
                except Exception as recording_error:
                    # Some meetings may not have recordings or we may not have permission
                    continue
        
        except ImportError:
            # Fallback: Try using call records API
            try:
                
                
                # Build filter for call records
                filter_query = f"startDateTime ge {start_date} and startDateTime le {end_date}"
                
                query_params = CallRecordsRequestBuilder.CallRecordsRequestBuilderGetQueryParameters()
                query_params.filter = filter_query
                
                request_config = RequestConfiguration(query_parameters=query_params)
                
                call_records_response = await graph.communications.call_records.get(
                    request_configuration=request_config
                )
                
                call_records = call_records_response.value if call_records_response and call_records_response.value else []
                
                for record in call_records:
                    recording_info = {
                        'call_record_id': record.id if hasattr(record, 'id') else 'N/A',
                        'start_time': record.start_date_time.isoformat() if hasattr(record, 'start_date_time') and record.start_date_time else 'N/A',
                        'end_time': record.end_date_time.isoformat() if hasattr(record, 'end_date_time') and record.end_date_time else 'N/A',
                        'organizer': record.organizer.user.display_name if hasattr(record, 'organizer') and record.organizer and hasattr(record.organizer, 'user') else 'Unknown'
                    }
                    recordings_list.append(recording_info)
                    
            except Exception as callrecord_error:
                pass
        
        return {
            'success': True,
            'recordings': recordings_list,
            'count': len(recordings_list),
            'error': None
        }
        
    except Exception as e:
        return {
            'success': False,
            'recordings': [],
            'count': 0,
            'error': f"Error fetching call recordings: {str(e)}"
        }


def format_recording_summary(recordings: List[Dict[str, Any]]) -> str:
    """
    Format a list of call recordings into a readable summary string.
    
    Args:
        recordings: List of recording dictionaries from get_teams_call_recordings
    
    Returns:
        Formatted string with recording summaries
    """
    if not recordings:
        return "🎥 No call recordings found in the specified time period."
    
    lines = [f"🎥 **Teams Call Recordings ({len(recordings)} found):**\n"]
    
    for i, recording in enumerate(recordings, 1):
        subject = recording.get('meeting_subject', 'Unknown Meeting')
        start_time = recording.get('meeting_start_time', recording.get('start_time', 'N/A'))
        created_by = recording.get('recording_created_by', recording.get('organizer', 'Unknown'))
        
        # Format datetime if it's a full ISO string
        if 'T' in str(start_time):
            try:
                dt = datetime.fromisoformat(start_time.replace('Z', '+00:00'))
                start_time = dt.strftime('%Y-%m-%d %H:%M UTC')
            except:
                pass
        
        lines.append(f"{i}. **{subject}**")
        lines.append(f"   📅 {start_time}")
        lines.append(f"   👤 {created_by}")
        
        if recording.get('recording_content_url'):
            lines.append(f"   🔗 Recording available")
        
        lines.append("")  # Empty line between recordings
    
    return "\n".join(lines)

