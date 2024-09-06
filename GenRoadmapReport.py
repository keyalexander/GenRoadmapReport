# Import necessary libraries
import os
import re
import pandas as pd
import platform
import subprocess
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.shared import OxmlElement
from docx.oxml.ns import qn
import docx.opc.constants
from docx.enum.text import WD_BREAK
from docx.oxml import OxmlElement
from docx.enum.section import WD_ORIENTATION
import win32com.client
import time

def read_excel_file(file_path):
    """
    Reads the Excel file containing Jira issues.
    
    Input:
    - file_path: str, path to the Excel file
    
    Output:
    - data: list of dictionaries, each representing a row in the Excel file
    """
    try:
        # Read the Excel file
        df = pd.read_excel(file_path)
        
        # Convert DataFrame to list of dictionaries
        data = df.to_dict('records')
        
        # Ensure all values are strings
        for row in data:
            for key, value in row.items():
                row[key] = str(value) if pd.notna(value) else ""
        
        return data
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return []

def process_data(data):
    """
    Processes the data from Excel, organizing it into a hierarchical structure.
    
    Input:
    - data: list of dictionaries, each representing a row in the Excel file
    
    Output:
    - structured_data: dict, hierarchical structure of Theme -> Goal -> Status -> Initiative -> Lead
    """
    structured_data = {}
    current_theme = None
    current_goal = None
    current_status = None
    current_initiative = None
    
    for row in data:
        issue_type = row['Issue Type']
        key = row['Key']
        summary = row['Summary']
        hebrew_summary = row['Hebrew Summary']
        status = row['Status']
        description = row['Description']
        start_date = row['Start date']
        due_date = row['Due date']
        
        if issue_type == 'Theme':
            current_theme = key
            structured_data[key] = {
                'summary': summary,
                'hebrew_summary': hebrew_summary,
                'goals': {}
            }
        elif issue_type == 'Goal':
            current_goal = key
            if current_theme:
                structured_data[current_theme]['goals'][key] = {
                    'summary': summary,
                    'hebrew_summary': hebrew_summary,
                    'description': description,
                    'statuses': {}
                }
        elif issue_type == 'Initiative':
            if current_theme and current_goal:
                current_status = status
                current_initiative = key
                if status not in structured_data[current_theme]['goals'][current_goal]['statuses']:
                    structured_data[current_theme]['goals'][current_goal]['statuses'][status] = {}
                structured_data[current_theme]['goals'][current_goal]['statuses'][status][key] = {
                    'summary': summary,
                    'hebrew_summary': hebrew_summary,
                    'description': description,
                    'start_date': start_date,
                    'due_date': due_date,
                    'leads': {}
                }
        elif issue_type == 'Lead':
            if current_theme and current_goal and current_status and current_initiative:
                structured_data[current_theme]['goals'][current_goal]['statuses'][current_status][current_initiative]['leads'][key] = {
                    'summary': summary,
                    'hebrew_summary': hebrew_summary,
                    'description': description
                }
        elif issue_type == '':
            if summary == "Not an issue":
                # If we encounter "Not an issue", we've reached the end of relevant data
                break
            else:
                # This row represents a status aggregator, but we'll ignore it
                # as we're already using the actual statuses from the initiatives
                pass
    
    return structured_data

def create_word_document(structured_data, output_file_path, include_todo=False):
    doc = Document()
    
    # Set page orientation to landscape
    section = doc.sections[0]
    new_width, new_height = section.page_height, section.page_width
    section.orientation = WD_ORIENTATION.LANDSCAPE
    section.page_width = new_width
    section.page_height = new_height
    
    # Update existing heading styles
    styles = doc.styles
    
    # Heading 1 style
    style_h1 = styles['Heading 1']
    style_h1.font.size = Pt(22)
    style_h1.paragraph_format.left_indent = Cm(0)
    
    # Heading 2 style
    style_h2 = styles['Heading 2']
    style_h2.font.size = Pt(20)
    style_h2.paragraph_format.left_indent = Cm(1)
    
    # Heading 3 style
    style_h3 = styles['Heading 3']
    style_h3.font.size = Pt(18)
    style_h3.paragraph_format.left_indent = Cm(0)
    
    # Add title
    doc.add_heading("Roadmap Status Report", level=0)
    
    # Add table of contents
    doc.add_paragraph("Table of Contents", style='TOC Heading')
     # Add the table of contents
    paragraph = doc.add_paragraph()
    run = paragraph.add_run()
    fldChar = OxmlElement('w:fldChar')
    fldChar.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = 'TOC \\o "1-3" \\h \\z \\u'
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    fldChar3 = OxmlElement('w:t')
    fldChar3.text = "Right-click to update field."
    fldChar4 = OxmlElement('w:fldChar')
    fldChar4.set(qn('w:fldCharType'), 'end')

    r_element = run._r
    r_element.append(fldChar)
    r_element.append(instrText)
    r_element.append(fldChar2)
    r_element.append(fldChar3)
    r_element.append(fldChar4)
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    # Define the order of statuses
    status_order = ['Done', 'In progress', 'Next']
    if include_todo:
        status_order.append('To Do')

    for theme_key, theme_data in structured_data.items():
        theme_printed = False
        
        for goal_key, goal_data in theme_data['goals'].items():
            goal_printed = False
            
            for status in status_order:
                if status in goal_data['statuses'] and goal_data['statuses'][status]:
                    initiatives = goal_data['statuses'][status]
                    
                    if not theme_printed:
                        heading = doc.add_heading(level=1)
                        heading.add_run(f"{theme_data['summary']} (")
                        add_hyperlink(heading.add_run(), f"https://omnisys.atlassian.net/browse/{theme_key}", theme_key)
                        heading.add_run(")")
                        
                        # Add Hebrew summary for theme
                        hebrew_summary = doc.add_paragraph()
                        hebrew_summary.add_run(handle_hebrew_text(theme_data['hebrew_summary']))
                        hebrew_summary.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                        
                        theme_printed = True
                    
                    if not goal_printed:
                        heading = doc.add_heading(level=2)
                        heading.add_run(f"{goal_data['summary']} (")
                        add_hyperlink(heading.add_run(), f"https://omnisys.atlassian.net/browse/{goal_key}", goal_key)
                        heading.add_run(")")
                        
                        # Add Hebrew summary for goal
                        hebrew_summary = doc.add_paragraph()
                        hebrew_summary.add_run(handle_hebrew_text(goal_data['hebrew_summary']))
                        hebrew_summary.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                        
                        goal_printed = True
                    
                    # Create status heading and table
                    status_heading = doc.add_heading(f"Status: {status}", level=3)                  
                    doc.add_paragraph()
                    table = doc.add_table(rows=1, cols=2)
                    table.style = 'Table Grid'
                    table.autofit = False
                    table.columns[0].width = Cm(5)
                    table.columns[1].width = Cm(10)
                    
                    hdr_cells = table.rows[0].cells
                    hdr_cells[0].text = 'Title'
                    hdr_cells[1].text = 'Description'
                    
                    for initiative_key, initiative_data in initiatives.items():
                        row_cells = table.add_row().cells
                        row_cells[0].paragraphs[0].add_run(f"{initiative_data['summary']} (")
                        add_hyperlink(row_cells[0].paragraphs[0].add_run(), f"https://omnisys.atlassian.net/browse/{initiative_key}", initiative_key)
                        row_cells[0].paragraphs[0].add_run(")")
                        
                        # Add Hebrew summary for initiative
                        hebrew_summary = row_cells[0].add_paragraph()
                        hebrew_summary.add_run(handle_hebrew_text(initiative_data['hebrew_summary']))
                        hebrew_summary.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                        
                        description = initiative_data['description']
                        row_cells[1].text = description

                        if initiative_data['leads']:
                            row_cells[1].add_paragraph("\n\n\n")  # Add 3 lines of separation
                            p = row_cells[1].add_paragraph("Linked initiatives:")
                            for lead_key, lead_data in initiative_data['leads'].items():
                                p = row_cells[1].add_paragraph("- ")
                                p.add_run(f"{lead_data['summary']} (")
                                add_hyperlink(p.add_run(), f"https://omnisys.atlassian.net/browse/{lead_key}", lead_key)
                                p.add_run(")")
                    
                    doc.add_paragraph()

    try:
        doc.save(output_file_path)
    except PermissionError:
        print(f"Error: Unable to save '{output_file_path}'. Please close the file if it's open and try again.")
        word = win32com.client.Dispatch('Word.Application')
        word.Documents.Open(output_file_path)
        word.ActiveDocument.Close()
        doc.save(output_file_path)
        word.Quit()
        del word
        return True
    except Exception as e:
        print(f"Error: {e}")
        return False
    
    return True

def add_hyperlink(run, url, text):
    """
    A function that places a hyperlink within a paragraph object.
    """
    r = run
    r.font.underline = True
    
    # This gets access to the document.xml.rels file and gets a new relation id value
    r_id = r.part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    
    # Create the w:hyperlink tag and add needed values
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id, )
    
    # Create a w:r element and a new w:rPr element
    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    
    # Join all the xml elements together add add the required text to the w:r element
    new_run.append(rPr)
    new_run.text = text
    hyperlink.append(new_run)
    
    # Create a new Run object and add the hyperlink into it
    r._r.append(hyperlink)
    
    return hyperlink

def handle_hebrew_text(text):
    """
    Handles Hebrew text to ensure proper display in the Word document.
    
    Input:
    - text: str, potentially containing Hebrew characters
    
    Output:
    - processed_text: str, text properly formatted for display
    """
    # Reverse the text for proper right-to-left display
    return text[::-1]

def main():
    """
    Main function to orchestrate the entire process.
    """
    # Find the Excel file matching the pattern
    excel_files = [f for f in os.listdir() if re.match(r"Roadmap.*\.xlsx", f)]
    
    if not excel_files:
        print("No matching Excel file found.")
        return

    # Use the first matching file
    file_path = excel_files[0]
    
    # Call read_excel_file function
    data = read_excel_file(file_path)
    
    if not data:
        print("No data read from the Excel file.")
        return

    # Print statement to confirm successful file reading
    print(f"Successfully read {len(data)} rows from {file_path}")

    # Print a few sample rows
    print("\nSample rows from the file:")
    for i, row in enumerate(data[:5], 1):
        print(f"Row {i}: {row}")
    print("...")  # Indicate there might be more rows

    # Add a toggle for including 'To Do' status
    include_todo = True  # Set this to True if you want to include 'To Do' status

    # Process the data
    structured_data = process_data(data)
    
    # Print a sample of the structured data to verify
    print("\nSample of structured data:")
    for theme_key, theme_data in list(structured_data.items())[:1]:
        print(f"Theme: {theme_key}")
        print(f"  Summary: {theme_data['summary']}")
        print(f"  Hebrew Summary: {theme_data['hebrew_summary'][::-1]}")
        print("  Goals:")
        for goal_key, goal_data in list(theme_data['goals'].items())[:1]:
            print(f"    Goal: {goal_key}")
            print(f"      Summary: {goal_data['summary']}")
            print(f"      Hebrew Summary: {goal_data['hebrew_summary'][::-1]}")
            print("      Statuses:")
            for status, initiatives in list(goal_data['statuses'].items())[:1]:
                print(f"        Status: {status}")
                for initiative_key, initiative_data in list(initiatives.items())[:1]:
                    print(f"          Initiative: {initiative_key}")
                    print(f"            Summary: {initiative_data['summary']}")
                    print(f"            Hebrew Summary: {initiative_data['hebrew_summary'][::-1]}")
                    print(f"            Description: {initiative_data['description']}")
                    print("            Leads:")
                    for lead_key, lead_data in list(initiative_data['leads'].items())[:1]:
                        print(f"              Lead: {lead_key}")
                        print(f"                Summary: {lead_data['summary']}")
                        print(f"                Hebrew Summary: {lead_data['hebrew_summary'][::-1]}")
    print("...")  # Indicate there might be more data
    # Create Word document
    output_file_name = "Roadmap Status Report.docx"
    output_file_path = os.path.abspath(output_file_name)
    if create_word_document(structured_data, output_file_path, include_todo):
        print(f"\nWord document created: {output_file_path}")

        print(f"File path: {output_file_path}")
        print(f"File exists: {os.path.exists(output_file_path)}")

        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = True
            doc = word.Documents.Open(output_file_path)
            print("Word document opened successfully")
        except Exception as e:
            print(f"Error opening Word document: {e}")

        time.sleep(5)  # Keep the script running for 5 seconds
    else:
        print("\nFailed to create Word document. Please check the error message above.")

if __name__ == "__main__":
    main()