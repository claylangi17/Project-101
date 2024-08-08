import google.generativeai as genai
import PIL.Image
import pandas as pd
import os
import re

# Configure the API (do not share this key!)
genai.configure(api_key='AIzaSyCZ8dmOfcKbMYUv_gXMskmB3LJ3DH8UAUw')

# Set up the model
model = genai.GenerativeModel('gemini-1.5-flash')

def compare_columns(existing_columns, new_columns):
    """Compares two lists of column headers for semantic similarity."""
    prompt = f"""
    Compare these two lists of column headers and determine if they are semantically similar:
    Existing columns: {existing_columns}
    New columns: {new_columns}
    
    Respond with only 'True' if they are similar, or 'False' if they are different.
    """
    response = model.generate_content(prompt)
    return response.text.strip().lower() == 'true'

def save_to_excel(df, base_filename):
    """Saves the DataFrame to an Excel file, appending if similar columns exist."""
    index = 1
    while True:
        filename = f"{base_filename}_{index}.xlsx"
        if not os.path.exists(filename):
            df.to_excel(filename, index=False)
            print(f"Data saved to new file: {filename}")
            break
        else:
            existing_df = pd.read_excel(filename)
            if compare_columns(existing_df.columns.tolist(), df.columns.tolist()):
                with pd.ExcelWriter(filename, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
                    df.to_excel(writer, sheet_name='Sheet1', index=False, header=False, startrow=writer.sheets['Sheet1'].max_row)
                print(f"Data appended to existing file: {filename}")
                break
        index += 1

def extract_data_from_image(image_path):
    """Extracts data from the image using the Gemini model."""
    with PIL.Image.open(image_path) as image:
        response = model.generate_content([
            """
            You are a highly accurate data extraction expert. Your task is to analyze the provided form image and extract all present information into a structured table format. 

            Follow these guidelines:

            **Understanding the Form:**
            * Carefully analyze the image, identifying all fields, sections, and their relationships.
            * Pay close attention to the form's structure, including headings, labels, lines, and boxes. 

            **Data Extraction and Formatting:**
            * Extract all data fields, even if they are empty. Use "-" to represent empty cells.
            * Accurately transcribe all text, preserving the original format as much as possible. 
              * This includes preserving date formats, currency symbols, and special characters.
            * For numerical data, include individual line items, calculate subtotals if present, and include the overall total.
            * For checkboxes, use "X" if checked, and "-" if not checked.
            * If there are any handwritten portions, transcribe them as accurately as possible. Use "[?]" to indicate uncertain characters. 

            **Table Structure:**
            * Organize the extracted data into a clear table format using pipe characters "|" as separators.
            * Each row in the table should represent a distinct data record or line item from the form.
            * The first row should contain the column headers, representing the extracted fields.

            **Example:**

            | Name | Address | Phone Number |
            |---|---|---|
            | John Doe | 123 Main St | 555-123-4567 |
            | Jane Smith | 456 Oak Ave | 555-987-6543 |

            **Important:**
            * Ensure consistency in column headers and data types throughout the table.
            * Maintain the original order of information as closely as possible.
            * Present the extracted data in a clean and well-structured format suitable for direct import into a spreadsheet.
            """,
            image
        ])

    return response.text

# --- Main Execution ---
if __name__ == "__main__":
    image_path = 'stream-verification-form.png' 
    extracted_text = extract_data_from_image(image_path)
    print("Extracted Text:\n", extracted_text)

    # --- Data Processing (Enhancements here) ---
    # 1. Split into lines and remove empty lines
    lines = [line.strip() for line in extracted_text.split('\n') if line.strip()]

    # 2. Find the header row (assuming it's the first non-empty line)
    header_row = lines[0]
    columns = [col.strip() for col in header_row.split('|')]

    # 3. Extract data rows (use regex to handle potential variations in separator)
    data_rows = []
    for line in lines[1:]:
        row_data = re.split(r'\s*\|\s*', line)  # More flexible separator handling
        data_rows.append(row_data)

    # --- DataFrame Creation and Saving ---
    df = pd.DataFrame(data_rows, columns=columns)
    base_filename = 'extracted_data'
    save_to_excel(df, base_filename)

    print("Data extraction and saving completed.") 