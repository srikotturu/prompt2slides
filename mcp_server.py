from fastmcp import FastMCP
import pptx_services

# Initialize FastMCP server
mcp = FastMCP("powerpoint")


@mcp.tool()
async def create_presentation(
    filename: str = "presentation"
) -> str:
    """
    Creates a new PowerPoint presentation.
    
    Args:
        filename (str, optional): Name of the presentation file without extension. Defaults to "presentation".
    
    Returns:
        str: Presentation ID for use in other operations.
    """
    return await pptx_services.create_presentation(filename=filename)


@mcp.tool()
async def add_slide(
    presentation_id: str,
    layout_index: int = 1,
    title: str | None = None,
    background_color: tuple | None = None,
) -> str:
    """
    Adds a new slide to the presentation.
    
    Args:
        presentation_id (str): ID of the presentation.
        layout_index (int, optional): Index of slide layout to use. Defaults to 1 (Title and Content).
            Available layouts:
            0 - Title Slide
            1 - Title and Content
            2 - Section Header
            3 - Two Content
            4 - Comparison
            5 - Title Only
            6 - Blank
            7 - Content with Caption
            8 - Picture with Caption
        title (str, optional): Slide title.
        background_color (tuple, optional): RGB tuple for background color (e.g., (255, 255, 255)).
    
    Returns:
        str: Slide ID for use in other operations.
    """
    return await pptx_services.add_slide(
        presentation_id=presentation_id,
        layout_index=layout_index,
        title=title,
        background_color=background_color
    )


@mcp.tool()
async def add_text(
    presentation_id: str,
    slide_id: str,
    text: str,
    left: float = 1.0,
    top: float = 2.0,
    width: float = 8.0,
    height: float = 1.0,
    font_size: int = 18,
    font_name: str = "Calibri",
    color: tuple = (0, 0, 0),
    bold: bool = False,
    italic: bool = False,
    underline: bool = False,
    alignment: str = "LEFT",
) -> str:
    """
    Adds text to a slide.
    
    Args:
        presentation_id (str): ID of the presentation.
        slide_id (str): ID of the slide.
        text (str): Text content.
        left (float, optional): Left position in inches. Defaults to 1.0.
        top (float, optional): Top position in inches. Defaults to 2.0.
        width (float, optional): Width of text box in inches. Defaults to 8.0.
        height (float, optional): Height of text box in inches. Defaults to 1.0.
        font_size (int, optional): Font size in points. Defaults to 18.
        font_name (str, optional): Font name. Defaults to "Calibri".
        color (tuple, optional): RGB color tuple (e.g., (255, 0, 0) for red). Defaults to (0, 0, 0).
        bold (bool, optional): Whether text should be bold. Defaults to False.
        italic (bool, optional): Whether text should be italic. Defaults to False.
        underline (bool, optional): Whether text should be underlined. Defaults to False.
        alignment (str, optional): Text alignment (LEFT, CENTER, RIGHT). Defaults to "LEFT".
    
    Returns:
        str: Text shape ID.
    """
    return await pptx_services.add_text(
        presentation_id=presentation_id,
        slide_id=slide_id,
        text=text,
        left=left,
        top=top,
        width=width,
        height=height,
        font_size=font_size,
        font_name=font_name,
        color=color,
        bold=bold,
        italic=italic,
        underline=underline,
        alignment=alignment,
    )


@mcp.tool()
async def add_paragraph(
    presentation_id: str,
    slide_id: str,
    text: str,
    left: float = 1.0,
    top: float = 2.0,
    width: float = 8.0,
    height: float = 3.0,
    font_size: int = 16,
    font_name: str = "Calibri",
    color: tuple = (0, 0, 0),
    alignment: str = "LEFT",
    line_spacing: float = 1.0,
) -> str:
    """
    Adds a paragraph (multi-line text) to a slide.
    
    Args:
        presentation_id (str): ID of the presentation.
        slide_id (str): ID of the slide.
        text (str): Paragraph text content.
        left (float, optional): Left position in inches. Defaults to 1.0.
        top (float, optional): Top position in inches. Defaults to 2.0.
        width (float, optional): Width of text box in inches. Defaults to 8.0.
        height (float, optional): Height of text box in inches. Defaults to 3.0.
        font_size (int, optional): Font size in points. Defaults to 16.
        font_name (str, optional): Font name. Defaults to "Calibri".
        color (tuple, optional): RGB color tuple. Defaults to (0, 0, 0).
        alignment (str, optional): Text alignment (LEFT, CENTER, RIGHT). Defaults to "LEFT".
        line_spacing (float, optional): Line spacing multiplier. Defaults to 1.0.
    
    Returns:
        str: Text box shape ID.
    """
    return await pptx_services.add_paragraph(
        presentation_id=presentation_id,
        slide_id=slide_id,
        text=text,
        left=left,
        top=top,
        width=width,
        height=height,
        font_size=font_size,
        font_name=font_name,
        color=color,
        alignment=alignment,
        line_spacing=line_spacing,
    )


@mcp.tool()
async def add_bullet_list(
    presentation_id: str,
    slide_id: str,
    items: list,
    left: float = 1.0,
    top: float = 2.0,
    width: float = 8.0,
    height: float = 3.0,
    font_size: int = 16,
    font_name: str = "Calibri",
    color: tuple = (0, 0, 0),
    level: int = 0,
    bullet_character: str | None = None,
) -> str:
    """
    Adds a bulleted list to a slide.
    
    Args:
        presentation_id (str): ID of the presentation.
        slide_id (str): ID of the slide.
        items (list): List of text items.
        left (float, optional): Left position in inches. Defaults to 1.0.
        top (float, optional): Top position in inches. Defaults to 2.0.
        width (float, optional): Width of text box in inches. Defaults to 8.0.
        height (float, optional): Height of text box in inches. Defaults to 3.0.
        font_size (int, optional): Font size in points. Defaults to 16.
        font_name (str, optional): Font name. Defaults to "Calibri".
        color (tuple, optional): RGB color tuple. Defaults to (0, 0, 0).
        level (int, optional): Indentation level (0 for top level). Defaults to 0.
        bullet_character (str, optional): Custom bullet character. Defaults to None.
    
    Returns:
        str: Text box shape ID.
    """
    return await pptx_services.add_bullet_list(
        presentation_id=presentation_id,
        slide_id=slide_id,
        items=items,
        left=left,
        top=top,
        width=width,
        height=height,
        font_size=font_size,
        font_name=font_name,
        color=color,
        level=level,
        bullet_character=bullet_character,
    )


@mcp.tool()
async def add_image(
    presentation_id: str,
    slide_id: str,
    image_path: str,
    left: float = 1.0,
    top: float = 2.0,
    width: float | None = None,
    height: float | None = None,
) -> str:
    """
    Adds an image to a slide.
    
    Args:
        presentation_id (str): ID of the presentation.
        slide_id (str): ID of the slide.
        image_path (str): Path to the image file.
        left (float, optional): Left position in inches. Defaults to 1.0.
        top (float, optional): Top position in inches. Defaults to 2.0.
        width (float, optional): Width of image in inches. Defaults to None (maintains aspect ratio).
        height (float, optional): Height of image in inches. Defaults to None (maintains aspect ratio).
    
    Returns:
        str: Image shape ID.
    """
    return await pptx_services.add_image(
        presentation_id=presentation_id,
        slide_id=slide_id,
        image_path=image_path,
        left=left,
        top=top,
        width=width,
        height=height,
    )


@mcp.tool()
async def add_chart(
    presentation_id: str,
    slide_id: str,
    chart_type: str,
    categories: list,
    data_series: list,
    left: float = 1.0,
    top: float = 2.0,
    width: float = 8.0,
    height: float = 4.0,
    chart_title: str | None = None,
) -> str:
    """
    Adds a chart to a slide.
    
    Args:
        presentation_id (str): ID of the presentation.
        slide_id (str): ID of the slide.
        chart_type (str): Type of chart (BAR_CLUSTERED, BAR_STACKED, COLUMN_CLUSTERED, 
                          LINE, PIE, SCATTER, AREA, RADAR, etc.)
        categories (list): List of category labels.
        data_series (list): List of tuples (series_name, values).
        left (float, optional): Left position in inches. Defaults to 1.0.
        top (float, optional): Top position in inches. Defaults to 2.0.
        width (float, optional): Width of chart in inches. Defaults to 8.0.
        height (float, optional): Height of chart in inches. Defaults to 4.0.
        chart_title (str, optional): Chart title. Defaults to None.
    
    Returns:
        str: Chart shape ID.
    """
    return await pptx_services.add_chart(
        presentation_id=presentation_id,
        slide_id=slide_id,
        chart_type=chart_type,
        categories=categories,
        data_series=data_series,
        left=left,
        top=top,
        width=width,
        height=height,
        chart_title=chart_title,
    )


@mcp.tool()
async def add_table(
    presentation_id: str,
    slide_id: str,
    data: list,
    left: float = 1.0,
    top: float = 2.0,
    width: float | None = None,
    height: float | None = None,
    column_widths: list | None = None,
    first_row_is_header: bool = True,
) -> str:
    """
    Adds a table to a slide.
    
    Args:
        presentation_id (str): ID of the presentation.
        slide_id (str): ID of the slide.
        data (list): 2D array/list of data.
        left (float, optional): Left position in inches. Defaults to 1.0.
        top (float, optional): Top position in inches. Defaults to 2.0.
        width (float, optional): Total width of table in inches. Defaults to None.
        height (float, optional): Total height of table in inches. Defaults to None.
        column_widths (list, optional): List of column widths in inches. Defaults to None.
        first_row_is_header (bool, optional): Whether to format first row as header. Defaults to True.
    
    Returns:
        str: Table shape ID.
    """
    return await pptx_services.add_table(
        presentation_id=presentation_id,
        slide_id=slide_id,
        data=data,
        left=left,
        top=top,
        width=width,
        height=height,
        column_widths=column_widths,
        first_row_is_header=first_row_is_header,
    )


@mcp.tool()
async def add_shape(
    presentation_id: str,
    slide_id: str,
    shape_type: str,
    left: float = 2.0,
    top: float = 2.0,
    width: float = 2.0,
    height: float = 2.0,
    fill_color: tuple | None = None,
    line_color: tuple | None = None,
    line_width: float = 1.0,
) -> str:
    """
    Adds a shape to a slide.
    
    Args:
        presentation_id (str): ID of the presentation.
        slide_id (str): ID of the slide.
        shape_type (str): Type of shape (RECTANGLE, OVAL, TRIANGLE, ARROW, etc.)
        left (float, optional): Left position in inches. Defaults to 2.0.
        top (float, optional): Top position in inches. Defaults to 2.0.
        width (float, optional): Width of shape in inches. Defaults to 2.0.
        height (float, optional): Height of shape in inches. Defaults to 2.0.
        fill_color (tuple, optional): RGB fill color. Defaults to None.
        line_color (tuple, optional): RGB line color. Defaults to None.
        line_width (float, optional): Line width in points. Defaults to 1.0.
    
    Returns:
        str: Shape ID.
    """
    return await pptx_services.add_shape(
        presentation_id=presentation_id,
        slide_id=slide_id,
        shape_type=shape_type,
        left=left,
        top=top,
        width=width,
        height=height,
        fill_color=fill_color,
        line_color=line_color,
        line_width=line_width,
    )


@mcp.tool()
async def add_header_footer(
    presentation_id: str,
    header_text: str | None = None,
    footer_text: str | None = None,
    slide_number: bool = True,
    date: bool = True,
) -> str:
    """
    Adds headers and footers to all slides.
    
    Args:
        presentation_id (str): ID of the presentation.
        header_text (str, optional): Header text. Defaults to None.
        footer_text (str, optional): Footer text. Defaults to None.
        slide_number (bool, optional): Whether to show slide numbers. Defaults to True.
        date (bool, optional): Whether to show date. Defaults to True.
    
    Returns:
        str: Result of the operation.
    """
    return await pptx_services.add_header_footer(
        presentation_id=presentation_id,
        header_text=header_text,
        footer_text=footer_text,
        slide_number=slide_number,
        date=date,
    )


@mcp.tool()
async def save_presentation(
    presentation_id: str,
    filename: str = "presentation",
) -> str:
    """
    Saves the presentation to a file.
    
    Args:
        presentation_id (str): ID of the presentation.
        filename (str, optional): Name of the file (without extension). Defaults to "presentation".
    
    Returns:
        str: Path to the saved file.
    """
    return await pptx_services.save_presentation(
        presentation_id=presentation_id,
        filename=filename,
    )


@mcp.tool()
async def create_complete_presentation(
    title: str,
    slides_content: list,
    subtitle: str | None = None,
    header: str | None = None,
    footer: str | None = None,
    filename: str | None = None,
    title_slide_color: tuple | None = None,
) -> str:
    """
    Creates a complete PowerPoint presentation with multiple slides.
    
    Args:
        title (str): Presentation title.
        slides_content (list): List of slide content dictionaries.
            Each dictionary may contain:
            - title (str): Slide title
            - text (str): Paragraph text
            - bullets (list): List of bullet points
            - image (str): Path to image file
            - chart_type (str): Type of chart
            - chart_categories (list): Categories for chart
            - chart_data (list): Data series for chart
            - table_data (list): 2D array of table data
        subtitle (str, optional): Subtitle for title slide. Defaults to None.
        header (str, optional): Header text for all slides. Defaults to None.
        footer (str, optional): Footer text for all slides. Defaults to None.
        filename (str, optional): Name of the file (without extension). Defaults to None.
        title_slide_color (tuple, optional): RGB color for title slide. Defaults to None.
    
    Returns:
        str: ID of the created presentation.
    """
    options = {
        "subtitle": subtitle,
        "header": header,
        "footer": footer,
        "filename": filename,
        "title_slide_color": title_slide_color,
    }
    
    return await pptx_services.create_complete_presentation(
        title=title,
        slides_content=slides_content,
        options=options,
    )


if __name__ == "__main__":
    # Run MCP server
    mcp.run(transport="stdio")