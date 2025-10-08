"""
Google Slides MCP Tools

This module provides MCP tools for interacting with Google Slides API.
"""

import logging
import asyncio
from typing import List, Dict, Any


from auth.service_decorator import require_google_service
from core.server import server
from core.utils import handle_http_errors
from core.comments import create_comment_tools

logger = logging.getLogger(__name__)


@server.tool()
@handle_http_errors("create_presentation", service_type="slides")
@require_google_service("slides", "slides")
async def create_presentation(
    service,
    user_google_email: str,
    title: str = "Untitled Presentation"
) -> str:
    """
    Create a new Google Slides presentation.

    Args:
        user_google_email (str): The user's Google email address. Required.
        title (str): The title for the new presentation. Defaults to "Untitled Presentation".

    Returns:
        str: Details about the created presentation including ID and URL.
    """
    logger.info(f"[create_presentation] Invoked. Email: '{user_google_email}', Title: '{title}'")

    body = {
        'title': title
    }

    result = await asyncio.to_thread(
        service.presentations().create(body=body).execute
    )

    presentation_id = result.get('presentationId')
    presentation_url = f"https://docs.google.com/presentation/d/{presentation_id}/edit"

    confirmation_message = f"""Presentation Created Successfully for {user_google_email}:
- Title: {title}
- Presentation ID: {presentation_id}
- URL: {presentation_url}
- Slides: {len(result.get('slides', []))} slide(s) created"""

    logger.info(f"Presentation created successfully for {user_google_email}")
    return confirmation_message


@server.tool()
@handle_http_errors("get_presentation", is_read_only=True, service_type="slides")
@require_google_service("slides", "slides_read")
async def get_presentation(
    service,
    user_google_email: str,
    presentation_id: str
) -> str:
    """
    Get details about a Google Slides presentation.

    Args:
        user_google_email (str): The user's Google email address. Required.
        presentation_id (str): The ID of the presentation to retrieve.

    Returns:
        str: Details about the presentation including title, slides count, and metadata.
    """
    logger.info(f"[get_presentation] Invoked. Email: '{user_google_email}', ID: '{presentation_id}'")

    result = await asyncio.to_thread(
        service.presentations().get(presentationId=presentation_id).execute
    )

    title = result.get('title', 'Untitled')
    slides = result.get('slides', [])
    page_size = result.get('pageSize', {})

    slides_info = []
    for i, slide in enumerate(slides, 1):
        slide_id = slide.get('objectId', 'Unknown')
        page_elements = slide.get('pageElements', [])
        slides_info.append(f"  Slide {i}: ID {slide_id}, {len(page_elements)} element(s)")

    confirmation_message = f"""Presentation Details for {user_google_email}:
- Title: {title}
- Presentation ID: {presentation_id}
- URL: https://docs.google.com/presentation/d/{presentation_id}/edit
- Total Slides: {len(slides)}
- Page Size: {page_size.get('width', {}).get('magnitude', 'Unknown')} x {page_size.get('height', {}).get('magnitude', 'Unknown')} {page_size.get('width', {}).get('unit', '')}

Slides Breakdown:
{chr(10).join(slides_info) if slides_info else '  No slides found'}"""

    logger.info(f"Presentation retrieved successfully for {user_google_email}")
    return confirmation_message


@server.tool()
@handle_http_errors("batch_update_presentation", service_type="slides")
@require_google_service("slides", "slides")
async def batch_update_presentation(
    service,
    user_google_email: str,
    presentation_id: str,
    requests: List[Dict[str, Any]]
) -> str:
    """
    Apply batch updates to a Google Slides presentation.

    Args:
        user_google_email (str): The user's Google email address. Required.
        presentation_id (str): The ID of the presentation to update.
        requests (List[Dict[str, Any]]): List of update requests to apply.

    Returns:
        str: Details about the batch update operation results.
    """
    logger.info(f"[batch_update_presentation] Invoked. Email: '{user_google_email}', ID: '{presentation_id}', Requests: {len(requests)}")

    body = {
        'requests': requests
    }

    result = await asyncio.to_thread(
        service.presentations().batchUpdate(
            presentationId=presentation_id,
            body=body
        ).execute
    )

    replies = result.get('replies', [])

    confirmation_message = f"""Batch Update Completed for {user_google_email}:
- Presentation ID: {presentation_id}
- URL: https://docs.google.com/presentation/d/{presentation_id}/edit
- Requests Applied: {len(requests)}
- Replies Received: {len(replies)}"""

    if replies:
        confirmation_message += "\n\nUpdate Results:"
        for i, reply in enumerate(replies, 1):
            if 'createSlide' in reply:
                slide_id = reply['createSlide'].get('objectId', 'Unknown')
                confirmation_message += f"\n  Request {i}: Created slide with ID {slide_id}"
            elif 'createShape' in reply:
                shape_id = reply['createShape'].get('objectId', 'Unknown')
                confirmation_message += f"\n  Request {i}: Created shape with ID {shape_id}"
            else:
                confirmation_message += f"\n  Request {i}: Operation completed"

    logger.info(f"Batch update completed successfully for {user_google_email}")
    return confirmation_message


@server.tool()
@handle_http_errors("get_page", is_read_only=True, service_type="slides")
@require_google_service("slides", "slides_read")
async def get_page(
    service,
    user_google_email: str,
    presentation_id: str,
    page_object_id: str
) -> str:
    """
    Get details about a specific page (slide) in a presentation.

    Args:
        user_google_email (str): The user's Google email address. Required.
        presentation_id (str): The ID of the presentation.
        page_object_id (str): The object ID of the page/slide to retrieve.

    Returns:
        str: Details about the specific page including elements and layout.
    """
    logger.info(f"[get_page] Invoked. Email: '{user_google_email}', Presentation: '{presentation_id}', Page: '{page_object_id}'")

    result = await asyncio.to_thread(
        service.presentations().pages().get(
            presentationId=presentation_id,
            pageObjectId=page_object_id
        ).execute
    )

    page_type = result.get('pageType', 'Unknown')
    page_elements = result.get('pageElements', [])

    elements_info = []
    for element in page_elements:
        element_id = element.get('objectId', 'Unknown')
        if 'shape' in element:
            shape_type = element['shape'].get('shapeType', 'Unknown')
            elements_info.append(f"  Shape: ID {element_id}, Type: {shape_type}")
        elif 'table' in element:
            table = element['table']
            rows = table.get('rows', 0)
            cols = table.get('columns', 0)
            elements_info.append(f"  Table: ID {element_id}, Size: {rows}x{cols}")
        elif 'line' in element:
            line_type = element['line'].get('lineType', 'Unknown')
            elements_info.append(f"  Line: ID {element_id}, Type: {line_type}")
        else:
            elements_info.append(f"  Element: ID {element_id}, Type: Unknown")

    confirmation_message = f"""Page Details for {user_google_email}:
- Presentation ID: {presentation_id}
- Page ID: {page_object_id}
- Page Type: {page_type}
- Total Elements: {len(page_elements)}

Page Elements:
{chr(10).join(elements_info) if elements_info else '  No elements found'}"""

    logger.info(f"Page retrieved successfully for {user_google_email}")
    return confirmation_message


@server.tool()
@handle_http_errors("get_page_thumbnail", is_read_only=True, service_type="slides")
@require_google_service("slides", "slides_read")
async def get_page_thumbnail(
    service,
    user_google_email: str,
    presentation_id: str,
    page_object_id: str,
    thumbnail_size: str = "MEDIUM"
) -> str:
    """
    Generate a thumbnail URL for a specific page (slide) in a presentation.

    Args:
        user_google_email (str): The user's Google email address. Required.
        presentation_id (str): The ID of the presentation.
        page_object_id (str): The object ID of the page/slide.
        thumbnail_size (str): Size of thumbnail ("LARGE", "MEDIUM", "SMALL"). Defaults to "MEDIUM".

    Returns:
        str: URL to the generated thumbnail image.
    """
    logger.info(f"[get_page_thumbnail] Invoked. Email: '{user_google_email}', Presentation: '{presentation_id}', Page: '{page_object_id}', Size: '{thumbnail_size}'")

    result = await asyncio.to_thread(
        service.presentations().pages().getThumbnail(
            presentationId=presentation_id,
            pageObjectId=page_object_id,
            thumbnailProperties_thumbnailSize=thumbnail_size,
            thumbnailProperties_mimeType='PNG'
        ).execute
    )

    thumbnail_url = result.get('contentUrl', '')

    confirmation_message = f"""Thumbnail Generated for {user_google_email}:
- Presentation ID: {presentation_id}
- Page ID: {page_object_id}
- Thumbnail Size: {thumbnail_size}
- Thumbnail URL: {thumbnail_url}

You can view or download the thumbnail using the provided URL."""

    logger.info(f"Thumbnail generated successfully for {user_google_email}")
    return confirmation_message


@server.tool()
@handle_http_errors("get_slide_content", is_read_only=True, service_type="slides")
@require_google_service("slides", "slides_read")
async def get_slide_content(
    service,
    user_google_email: str,
    presentation_id: str,
    page_object_id: str
) -> str:
    """
    Get detailed content from a specific slide, including text, tables, and image/media descriptions.

    Args:
        user_google_email (str): The user's Google email address. Required.
        presentation_id (str): The ID of the presentation.
        page_object_id (str): The object ID of the page/slide to retrieve content from.

    Returns:
        str: A detailed summary of the slide's content, listing text, tables, images, and media elements.
    """
    logger.info(f"[get_slide_content] Invoked. Email: '{user_google_email}', Presentation: '{presentation_id}', Page: '{page_object_id}'")

    result = await asyncio.to_thread(
        service.presentations().pages().get(
            presentationId=presentation_id,
            pageObjectId=page_object_id
        ).execute
    )

    page_elements = result.get('pageElements', [])

    content_blocks = []
    for element in page_elements:
        element_id = element.get('objectId', 'Unknown')
        if 'shape' in element:
            shape = element['shape']
            shape_type = shape.get('shapeType', 'Unknown')
            text_content = ""
            if 'text' in shape:
                text_elements = shape['text'].get('textElements', [])
                for te in text_elements:
                    if 'textRun' in te:
                        text_content += te['textRun'].get('content', "")
            content_blocks.append(f"[Shape] ID: {element_id}, Type: {shape_type}\n  Text: {text_content.strip() if text_content else '(empty)'}")
        elif 'table' in element:
            table = element['table']
            rows = table.get('rows', 0)
            cols = table.get('columns', 0)
            table_content = []
            cells = table.get('tableRows', [])
            for row_idx, row in enumerate(cells):
                cell_row = []
                for col_idx, cell in enumerate(row.get('tableCells', [])):
                    cell_text = ""
                    text_elements = cell.get('text', {}).get('textElements', [])
                    for te in text_elements:
                        if 'textRun' in te:
                            cell_text += te['textRun'].get('content', "")
                    cell_row.append(cell_text.strip())
                table_content.append(f"    Row {row_idx+1}: {cell_row}")
            content_blocks.append(f"[Table] ID: {element_id}, Size: {rows}x{cols}\n{chr(10).join(table_content) if table_content else '    (empty table)'}")
        elif 'image' in element:
            image = element['image']
            description = image.get('description', '')
            content_blocks.append(f"[Image] ID: {element_id}\n  Description: {description or '(no description)'}")
        elif 'video' in element:
            video = element['video']
            video_url = video.get('url', '(unknown)')
            content_blocks.append(f"[Video] ID: {element_id}\n  URL: {video_url}")
        elif 'line' in element:
            line_type = element['line'].get('lineType', 'Unknown')
            content_blocks.append(f"[Line] ID: {element_id}, Type: {line_type}")
        else:
            content_blocks.append(f"[Element] ID: {element_id}, Type: Unknown")

    confirmation_message = f"""Slide Content Details for {user_google_email}:
- Presentation ID: {presentation_id}
- Page ID: {page_object_id}
- Total Elements: {len(page_elements)}

Content Summary:
{chr(10).join(content_blocks) if content_blocks else '  No content found on this slide.'}
"""
    logger.info(f"Slide content retrieved successfully for {user_google_email}")
    return confirmation_message



@server.tool()
@handle_http_errors("update_slide_content", service_type="slides")
@require_google_service("slides", "slides")
async def update_slide_content(
    service,
    user_google_email: str,
    presentation_id: str,
    page_object_id: str,
    updates: List[Dict[str, Any]]
) -> str:
    """
    Update the content of a specific slide (page) in a Google Slides presentation.
    
    This tool allows updating text in shapes, table cells, or replacing/setting image descriptions on a slide.
    You must specify one or more updates, each targeting a page element by objectId and providing new content.

    Args:
        user_google_email (str): The user's Google email address. Required.
        presentation_id (str): The ID of the presentation.
        page_object_id (str): The object ID of the page/slide to update.
        updates (List[Dict[str, Any]]): List of updates to apply. Each update must contain:
            - "element_id": The objectId of the element to update (shape, table cell, image, etc.).
            - "update_type": One of "shape_text", "table_cell", "image_description".
            - For "shape_text": Must include "text" (str) to set.
            - For "table_cell": Must include "row" (int), "column" (int), and "text" (str).
            - For "image_description": Must include "description" (str).

    Returns:
        str: Details about the update operation results.
    """
    logger.info(f"[update_slide_content] Invoked. Email: '{user_google_email}', Presentation: '{presentation_id}', Page: '{page_object_id}', Updates: {len(updates)}")

    # Build batchUpdate requests
    requests = []
    for upd in updates:
        element_id = upd.get("element_id")
        update_type = upd.get("update_type")
        if update_type == "shape_text":
            text = upd.get("text", "")
            requests.append({
                "deleteText": {
                    "objectId": element_id,
                    "textRange": {"type": "ALL"}
                }
            })
            requests.append({
                "insertText": {
                    "objectId": element_id,
                    "insertionIndex": 0,
                    "text": text
                }
            })
        elif update_type == "table_cell":
            row = upd.get("row")
            col = upd.get("column")
            text = upd.get("text", "")
            # Table cells use a cell location
            requests.append({
                "deleteText": {
                    "objectId": element_id,
                    "cellLocation": {
                        "rowIndex": row,
                        "columnIndex": col
                    },
                    "textRange": {"type": "ALL"}
                }
            })
            requests.append({
                "insertText": {
                    "objectId": element_id,
                    "cellLocation": {
                        "rowIndex": row,
                        "columnIndex": col
                    },
                    "insertionIndex": 0,
                    "text": text
                }
            })
        elif update_type == "image_description":
            description = upd.get("description", "")
            requests.append({
                "updateImageProperties": {
                    "objectId": element_id,
                    "fields": "description",
                    "imageProperties": {
                        "description": description
                    }
                }
            })
        else:
            logger.warning(f"Unknown update_type '{update_type}' for element {element_id}")

    if not requests:
        return f"No valid updates were provided."

    # Execute the requests as a batch update
    result = await asyncio.to_thread(
        service.presentations().batchUpdate(
            presentationId=presentation_id,
            body={"requests": requests}
        ).execute
    )

    replies = result.get("replies", [])

    summary = []
    update_idx = 0
    for upd in updates:
        update_type = upd.get("update_type")
        element_id = upd.get("element_id")
        if update_type == "shape_text":
            summary.append(f"Shape (ID: {element_id}): Text updated")
            update_idx += 2
        elif update_type == "table_cell":
            row = upd.get("row")
            col = upd.get("column")
            summary.append(f"Table (ID: {element_id}), Cell ({row}, {col}): Text updated")
            update_idx += 2
        elif update_type == "image_description":
            summary.append(f"Image (ID: {element_id}): Description updated")
            update_idx += 1
        else:
            summary.append(f"Element (ID: {element_id}): Skipped (unknown type)")

    confirmation_message = f"""Slide Content Update Results for {user_google_email}:
- Presentation ID: {presentation_id}
- Page ID: {page_object_id}
- Updates Requested: {len(updates)}
- Requests Sent: {len(requests)}
- Replies Received: {len(replies)}

Update Summary:
{chr(10).join(summary)}
"""

    logger.info(f"Slide content updated for {user_google_email} - {len(updates)} updates.")
    return confirmation_message


@server.tool()
@handle_http_errors("copy_slide", service_type="slides")
@require_google_service("slides", "slides")
async def copy_slide(
    service,
    user_google_email: str,
    presentation_id: str,
    source_page_object_id: str,
    updates: List[Dict[str, Any]] = None
) -> str:
    """
    Copy (clone) a slide from an existing slide in a presentation, inserting it right after the source slide,
    and optionally update the new slide's content.

    Args:
        user_google_email (str): The user's Google email address. Required.
        presentation_id (str): The ID of the presentation.
        source_page_object_id (str): The object ID of the slide to copy.
        updates (List[Dict[str, Any]], optional): List of updates to apply to the new slide (see update_slide_content for format).

    Returns:
        str: Details about the copy operation and any updates applied.
    """
    logger.info(f"[copy_slide] Invoked. Email: '{user_google_email}', Presentation: '{presentation_id}', Source Page: '{source_page_object_id}', Updates: {len(updates) if updates else 0}")

    # Retrieve the presentation to find the insertion index
    presentation = await asyncio.to_thread(
        service.presentations().get(presentationId=presentation_id).execute
    )
    slides = presentation.get("slides", [])
    insertion_index = None
    for idx, slide in enumerate(slides):
        if slide.get("objectId") == source_page_object_id:
            insertion_index = idx + 1
            break
    if insertion_index is None:
        return f"Source slide with object ID '{source_page_object_id}' not found."

    # Prepare the duplicateObject request to always insert right after the source slide
    duplicate_request = {
        "duplicateObject": {
            "objectId": source_page_object_id,
            "insertionIndex": insertion_index
        }
    }

    requests = [duplicate_request]

    # Execute the duplicate request and get the new slide id
    batch_result = await asyncio.to_thread(
        service.presentations().batchUpdate(
            presentationId=presentation_id,
            body={"requests": requests}
        ).execute
    )
    reply = batch_result.get("replies", [{}])[0]
    new_slide_id = reply.get("duplicateObject", {}).get("objectId")
    if not new_slide_id:
        return f"Failed to duplicate slide {source_page_object_id}."

    update_summary = ""
    # If updates are provided, apply them to the new slide
    if updates:
        update_requests = []
        for upd in updates:
            element_id = upd.get("element_id")
            update_type = upd.get("update_type")
            if update_type == "shape_text":
                text = upd.get("text", "")
                update_requests.append({
                    "deleteText": {
                        "objectId": element_id,
                        "textRange": {"type": "ALL"}
                    }
                })
                update_requests.append({
                    "insertText": {
                        "objectId": element_id,
                        "insertionIndex": 0,
                        "text": text
                    }
                })
            elif update_type == "table_cell":
                row = upd.get("row")
                col = upd.get("column")
                text = upd.get("text", "")
                update_requests.append({
                    "deleteText": {
                        "objectId": element_id,
                        "cellLocation": {
                            "rowIndex": row,
                            "columnIndex": col
                        },
                        "textRange": {"type": "ALL"}
                    }
                })
                update_requests.append({
                    "insertText": {
                        "objectId": element_id,
                        "cellLocation": {
                            "rowIndex": row,
                            "columnIndex": col
                        },
                        "insertionIndex": 0,
                        "text": text
                    }
                })
            elif update_type == "image_description":
                description = upd.get("description", "")
                update_requests.append({
                    "updateImageProperties": {
                        "objectId": element_id,
                        "fields": "description",
                        "imageProperties": {
                            "description": description
                        }
                    }
                })
        if update_requests:
            update_result = await asyncio.to_thread(
                service.presentations().batchUpdate(
                    presentationId=presentation_id,
                    body={"requests": update_requests}
                ).execute
            )
            update_summary = f"\n- Updates Applied: {len(update_requests)} request(s) sent, {len(update_result.get('replies', []))} replies received."

    confirmation_message = f"""Slide Copied Successfully for {user_google_email}:
- Presentation ID: {presentation_id}
- Source Page ID: {source_page_object_id}
- New Slide ID: {new_slide_id}
- Insertion Index: {insertion_index}
{update_summary}
- URL: https://docs.google.com/presentation/d/{presentation_id}/edit#slide=id.{new_slide_id}
"""
    logger.info(f"Slide copied successfully for {user_google_email} - New ID: {new_slide_id}")
    return confirmation_message


@server.tool()
@handle_http_errors("create_slide", service_type="slides")
@require_google_service("slides", "slides")
async def create_slide(
    service,
    user_google_email: str,
    presentation_id: str,
    insertion_index: int = None,
    layout_id: str = None
) -> str:
    """
    Create a new (empty) slide in a Google Slides presentation.

    Args:
        user_google_email (str): The user's Google email address. Required.
        presentation_id (str): The ID of the presentation.
        insertion_index (int, optional): The zero-based position where the new slide will be inserted. If None, inserts at the end.
        layout_id (str, optional): The layout object ID to use for the new slide. If None, uses default layout.

    Returns:
        str: Details about the created slide, including new slide ID and position.
    """
    logger.info(f"[create_slide] Invoked. Email: '{user_google_email}', Presentation: '{presentation_id}', Insertion Index: {insertion_index}, Layout ID: {layout_id}")

    request = {
        "createSlide": {}
    }
    if insertion_index is not None:
        request["createSlide"]["insertionIndex"] = insertion_index
    if layout_id is not None:
        request["createSlide"]["slideLayoutReference"] = {"layoutId": layout_id}

    requests = [request]

    # Execute the createSlide request and get the new slide id
    result = await asyncio.to_thread(
        service.presentations().batchUpdate(
            presentationId=presentation_id,
            body={"requests": requests}
        ).execute
    )

    reply = result.get("replies", [{}])[0]
    new_slide_id = reply.get("createSlide", {}).get("objectId")
    if not new_slide_id:
        return f"Failed to create a new slide."

    confirmation_message = f"""Slide Created Successfully for {user_google_email}:
- Presentation ID: {presentation_id}
- New Slide ID: {new_slide_id}
- Insertion Index: {insertion_index if insertion_index is not None else 'end'}
- Layout ID: {layout_id if layout_id else 'default'}
- URL: https://docs.google.com/presentation/d/{presentation_id}/edit#slide=id.{new_slide_id}
"""
    logger.info(f"Slide created successfully for {user_google_email} - New ID: {new_slide_id}")
    return confirmation_message



# Create comment management tools for slides
_comment_tools = create_comment_tools("presentation", "presentation_id")
read_presentation_comments = _comment_tools['read_comments']
create_presentation_comment = _comment_tools['create_comment']
reply_to_presentation_comment = _comment_tools['reply_to_comment']
resolve_presentation_comment = _comment_tools['resolve_comment']

# Aliases for backwards compatibility and intuitive naming
read_slide_comments = read_presentation_comments
create_slide_comment = create_presentation_comment
reply_to_slide_comment = reply_to_presentation_comment
resolve_slide_comment = resolve_presentation_comment