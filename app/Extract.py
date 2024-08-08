import google.generativeai as genai
import PIL.Image
import pandas as pd
import os

# Configure the API (do not share this key!)
genai.configure(api_key='AIzaSyCZ8dmOfcKbMYUv_gXMskmB3LJ3DH8UAUw')

# Set up the model with the new name
model = genai.GenerativeModel('gemini-1.5-flash')

def compare_columns(existing_columns, new_columns):
    prompt = f"""
    Compare these two lists of column headers and determine if they are semantically similar:
    Existing columns: {existing_columns}
    New columns: {new_columns}
    
    Respond with only 'True' if they are similar, or 'False' if they are different.
    """
    response = model.generate_content(prompt)
    return response.text.strip().lower() == 'true'

def save_to_excel(df, base_filename):
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

# Load the image
image = PIL.Image.open('order-form.png')

response = model.generate_content([
    """
    Act as a data analyst tasked with accurately transcribing information from a traditional form image into a structured Excel format. Carefully analyze the provided image and extract all relevant information. Follow these guidelines:

    1. Identify all distinct sections of the form (e.g., header information, personal details, itemized lists, totals).
    2. Create a comprehensive list of column headers based on the form's content. Include all fields present in the form, even if some are empty.
    3. For any lists or repeated sections in the form:
       - Create separate entries that include all associated information.
       - Ensure each item or entry is on its own row, even if they belong to the same group or section.
    4. Pay attention to and preserve:
       - Date formats
       - Currency symbols and number formats
       - Check boxes or multiple choice selections
       - Any hierarchical relationships in the data
    5. For numerical data:
       - Include individual line items
       - Calculate and include subtotals if present
       - Include the overall total if applicable
    6. If there are multiple pages or sections, clearly indicate transitions between them.
    7. For any handwritten parts, transcribe them as accurately as possible, indicating uncertainty with [?] if necessary.

    Present the extracted data in a table format suitable for direct import into Excel:
    - Use a pipe character (|) to separate columns
    - Put each row on a new line
    - Start with a row of column headers
    - Include all data, using '-' for empty fields
    - Ensure that every row has the same number of pipe separators
    - Preserve the original order of information as much as possible

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

# Extract the text from the response
extracted_text = response.text

# Print extracted text for debugging
print("Extracted Text:")
print(extracted_text)

# Parse the extracted text into a DataFrame
table_data = extracted_text.split('\n')

# Remove the separator line if it exists
table_data = [row for row in table_data if '---' not in row]

# Extract column headers and data rows
columns = [col.strip() for col in table_data[0].split('|') if col.strip()]
data_rows = [[cell.strip() for cell in row.split('|') if cell.strip()] for row in table_data[1:] if row.strip()]

# Debug printing
print("Columns:")
print(columns)
print("Data Rows:")
for row in data_rows:
    print(row)

# Find the maximum number of columns
max_columns = max(len(columns), max(len(row) for row in data_rows))

# Pad columns and rows to ensure consistent column count
columns = columns + [''] * (max_columns - len(columns))
data_rows = [row + [''] * (max_columns - len(row)) for row in data_rows]

# Create DataFrame
df = pd.DataFrame(data_rows, columns=columns)

# Remove any empty columns
df = df.dropna(axis=1, how='all')

# Define the base filename
base_filename = 'Result/extracted_data'

# Save the data to Excel
save_to_excel(df, base_filename)

print("Data extraction and saving completed successfully.")