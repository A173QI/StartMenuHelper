class StyleSheet:
    """
    Stylesheet definitions for the application, following Windows design language.
    
    Colors:
    - Primary: #0078D4 (Windows blue)
    - Secondary: #107C10 (success green)
    - Warning: #D83B01 (alert orange)
    - Neutral: #E1E1E1 (light grey)
    - Text: #323130 (dark grey)
    """
    # General App Styles
    APP_TITLE = """
        font-family: 'Segoe UI', sans-serif;
        font-size: 24px;
        font-weight: 600;
        color: #323130;
        margin-bottom: 8px;
    """
    
    APP_DESCRIPTION = """
        font-family: 'Segoe UI', sans-serif;
        font-size: 14px;
        color: #605E5C;
        margin-bottom: 16px;
    """
    
    SECTION_TITLE = """
        font-family: 'Segoe UI', sans-serif;
        font-size: 16px;
        font-weight: 600;
        color: #323130;
        margin-top: 16px;
        margin-bottom: 8px;
    """
    
    SEPARATOR = """
        background-color: #E1E1E1;
        max-height: 1px;
    """
    
    # Drop Area Styles
    DROP_AREA = """
        background-color: #F3F2F1;
        border: 2px dashed #C8C6C4;
        border-radius: 8px;
        padding: 20px;
    """
    
    DROP_AREA_DRAG_OVER = """
        background-color: #F0F8FF;
        border: 2px dashed #0078D4;
        border-radius: 8px;
        padding: 20px;
    """
    
    DROP_AREA_TITLE = """
        font-family: 'Segoe UI', sans-serif;
        font-size: 16px;
        font-weight: 600;
        color: #323130;
    """
    
    DROP_AREA_SUBTITLE = """
        font-family: 'Segoe UI', sans-serif;
        font-size: 14px;
        color: #605E5C;
    """
    
    # Form Settings Styles
    SETTINGS_GROUP = """
        border: 1px solid #E1E1E1;
        border-radius: 8px;
        padding: 16px;
        background-color: #FFFFFF;
    """
    
    # Button Styles
    PRIMARY_BUTTON = """
        font-family: 'Segoe UI', sans-serif;
        font-size: 14px;
        font-weight: 600;
        background-color: #0078D4;
        color: white;
        border: none;
        border-radius: 4px;
        padding: 8px 16px;
        min-width: 100px;
    """
    
    PRIMARY_BUTTON_HOVER = """
        background-color: #106EBE;
    """
    
    PRIMARY_BUTTON_PRESSED = """
        background-color: #005A9E;
    """
    
    SECONDARY_BUTTON = """
        font-family: 'Segoe UI', sans-serif;
        font-size: 14px;
        background-color: #F3F2F1;
        color: #323130;
        border: 1px solid #8A8886;
        border-radius: 4px;
        padding: 8px 16px;
        min-width: 100px;
    """
    
    SECONDARY_BUTTON_HOVER = """
        background-color: #EDEBE9;
    """
    
    SECONDARY_BUTTON_PRESSED = """
        background-color: #E1DFDD;
    """
    
    # Preview Styles
    PREVIEW_WIDGET = """
        background-color: #FCFCFC;
        border-radius: 8px;
        padding: 16px;
    """
    
    PREVIEW_FRAME = """
        background-color: #F5F5F5;
        border: 1px solid #E1E1E1;
        border-radius: 8px;
        padding: 16px;
    """
