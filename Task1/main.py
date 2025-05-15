import pandas as pd
 
def excel_to_mermaid(excel_filepath, sheet_name='Sheet1'):
    """
    Reads data from an Excel file and generates Mermaid syntax for a UML class diagram.
    Args:
        excel_filepath (str): The path to the Excel file.
        sheet_name (str): The name of the sheet to read (default is 'Sheet1').
    Returns:
        str: The Mermaid syntax for the UML class diagram, or None if an error occurs.
    """
    try:
        df = pd.read_excel(excel_filepath, sheet_name=sheet_name)
        # Assuming your Excel sheet has columns like:
        # Class, Attributes, Methods, Relationships
        # Remove triple backticks from mermaid syntax when used programmatically
        mermaid_syntax = """classDiagram\n    direction LR\n"""
        for index, row in df.iterrows():
            class_name = row.get('Class', '')
            attributes = row.get('Attributes', '')
            methods = row.get('Methods', '')
            relationships = row.get('Relationships', '')
            # Skip rows with no class name
            if not class_name or pd.isna(class_name):
                continue
            # Start with the class declaration
            mermaid_syntax += f"    class {class_name} {{\n"
            # Handle attributes
            if attributes and not pd.isna(attributes):
                for attr in str(attributes).split('\n'):
                    if attr.strip():  # Ensure attribute is not empty
                        mermaid_syntax += f"        {attr}\n"
            # Handle methods
            if methods and not pd.isna(methods):
                for method in str(methods).split('\n'):
                    if method.strip():  # Ensure method is not empty
                        # Only add () if not already present
                        if '()' in method:
                            mermaid_syntax += f"        {method}\n"
                        else:
                            mermaid_syntax += f"        {method}()\n"
            mermaid_syntax += "    }\n"  # Close the class
            # Only add relationships if they exist
            if relationships and not pd.isna(relationships):
                for relation in str(relationships).split('\n'):
                    if relation.strip():  # Ensure relationship is not empty
                        mermaid_syntax += f"    {relation}\n"
        return mermaid_syntax
    except FileNotFoundError:
        print(f"Error: File not found at '{excel_filepath}'")
        return None
    except KeyError as e:
        print(f"Error: Column '{e}' not found in the Excel sheet.")
        return None
    except Exception as e:
        print(f"An error occurred: {e}")
        return None
 
def generate_html(mermaid_syntax):
    """
    Generates an HTML file that uses mermaid.js to render the UML diagram.
    Args:
        mermaid_syntax (str): The Mermaid syntax for the diagram.
    Returns:
        str: The content of the HTML file.
    """
    # Add triple backticks and mermaid keyword for HTML display
    clean_syntax = mermaid_syntax.replace('```', '').replace('mermaid', '').strip()
    html_content = f"""<!DOCTYPE html>
<html>
<head>
<title>UML Class Diagram</title>
<script src="https://cdn.jsdelivr.net/npm/mermaid@10/dist/mermaid.min.js"></script>
<script>
        mermaid.initialize({{ startOnLoad: true }});
</script>
<style>
        body {{
            font-family: Arial, sans-serif;
            margin: 20px;
        }}
        h1 {{
            color: #333;
        }}
        .diagram-container {{
            margin-top: 20px;
            padding: 10px;
            border: 1px solid #ddd;
            border-radius: 5px;
        }}
</style>
</head>
<body>
<h1>UML Class Diagram</h1>
<div class="diagram-container">
<div class="mermaid">
{clean_syntax}
</div>
</div>
</body>
</html>"""
    return html_content
 
def save_to_file(content, filename):
    """
    Saves content to a file.
    Args:
        content (str): The content to save.
        filename (str): The name of the file to save to.
    Returns:
        bool: True if successful, False otherwise.
    """
    try:
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(content)
        return True
    except Exception as e:
        print(f"Error saving file: {e}")
        return False
 
if __name__ == "__main__":
    excel_file = 'Task1/Book.xlsx'  # Replace with your Excel file path
    mermaid_code = excel_to_mermaid(excel_file)
    if mermaid_code:
        # Generate HTML and save to file
        html_output = generate_html(mermaid_code)
        if save_to_file(html_output, 'Task1/uml_diagram.html'):
            print("HTML file 'uml_diagram.html' generated successfully. Open it in your browser.")
        # Optionally, save the raw mermaid code as well
        if save_to_file(mermaid_code, 'Task1/mermaid_code.txt'):
            print("Raw Mermaid code saved to 'mermaid_code.txt'")