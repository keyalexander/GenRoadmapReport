# Import necessary libraries
import os
import re
import pandas as pd
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.shared import OxmlElement
from docx.oxml.ns import qn
import docx.opc.constants

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
            # This row represents a status aggregator, but we'll ignore it
            # as we're already using the actual statuses from the initiatives
            pass
    
    return structured_data

def create_word_document(structured_data, output_file_path):
    doc = Document()
    
    for theme_key, theme_data in structured_data.items():
        theme_title = f"Theme: {theme_data['summary']} ({theme_key})"
        heading = doc.add_heading(theme_title, level=1)
        add_hyperlink(heading.runs[0], f"https://omnisys.atlassian.net/browse/{theme_key}", theme_title)
        
        for goal_key, goal_data in theme_data['goals'].items():
            doc.add_heading(f"Goal: {goal_data['summary']}", level=2)
            
            # Define the order of statuses
            status_order = ['Done', 'In progress', 'Next', 'To Do']
            
            for status in status_order:
                if status in goal_data['statuses']:
                    initiatives = goal_data['statuses'][status]
                    doc.add_paragraph(f"Status: {status}", style='Heading 3')
                    
                    # Create a table for initiatives of this status
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
                        initiative_title = f"{initiative_data['summary']} ({initiative_key})"
                        add_hyperlink(row_cells[0].paragraphs[0].add_run(), f"https://omnisys.atlassian.net/browse/{initiative_key}", initiative_title)
                        
                        description = initiative_data['description']
                        row_cells[1].text = description

                        if initiative_data['leads']:
                            row_cells[1].add_paragraph("\n\n\n")  # Add 3 lines of separation
                            p = row_cells[1].add_paragraph("Linked initiatives:")
                            for lead_key, lead_data in initiative_data['leads'].items():
                                lead_title = f"{lead_data['summary']} ({lead_key})"
                                p = row_cells[1].add_paragraph("- ")
                                add_hyperlink(p.add_run(lead_title), f"https://omnisys.atlassian.net/browse/{lead_key}", lead_title)
                    
                    doc.add_paragraph()  # Add some space after each table
    
    doc.save(output_file_path)

def add_hyperlink(run, url, text):
    """
    A function that places a hyperlink within a paragraph object.
    """
    r = run
    r.font.color.rgb = RGBColor(0, 0, 255)
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
    pass

def main():
    """
    Main function to orchestrate the entire process.
    """
    # Find the Excel file matching the pattern
    excel_files = [f for f in os.listdir() if re.match(r"Roadmap.*\.xls", f)]
    
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
                    print("            Leads:")
                    for lead_key, lead_data in list(initiative_data['leads'].items())[:1]:
                        print(f"              Lead: {lead_key}")
                        print(f"                Summary: {lead_data['summary']}")
                        print(f"                Hebrew Summary: {lead_data['hebrew_summary'][::-1]}")
    print("...")  # Indicate there might be more data
    # Create Word document
    output_file_path = "Roadmap Status Report.docx"
    create_word_document(structured_data, output_file_path)
    print(f"\nWord document created: {output_file_path}")

if __name__ == "__main__":
    main()
