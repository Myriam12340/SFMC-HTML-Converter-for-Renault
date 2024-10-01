import tkinter as tk
from tkinter import filedialog, ttk
import pandas as pd
import os
import warnings
import re

# Ignore warnings related to openpyxl
warnings.filterwarnings("ignore", category=UserWarning, module='openpyxl')


# Function to read sheet names from the selected Excel file
def get_sheet_names_from_excel(file_path):
    try:
        # Load the Excel file and print available sheet names
        xls = pd.ExcelFile(file_path)
        print("Found sheets:", xls.sheet_names)
        return xls.sheet_names
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return []

# Function to read replacement mappings from an Excel file
def read_replacements_from_excel(file_path, country):
    try:
        # Load the Excel file and check if the country sheet exists
        xls = pd.ExcelFile(file_path)
        if country not in xls.sheet_names:
            print(f"Sheet for country '{country}' not found in Excel file.")
            return {}

        # Load the sheet into a DataFrame and convert it to a dictionary
        df = pd.read_excel(file_path, sheet_name=country)
        return pd.Series(df['new'].values, index=df['old']).to_dict()
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return {}

# Function to read specific tags or styles from an Excel file
def read_image_style_from_excel(file_path, country, model, tag_type, purpose=None):
    try:
        # Load the Excel file and check if the country sheet exists
        xls = pd.ExcelFile(file_path)
        if country not in xls.sheet_names:
            print(f"Sheet for country '{country}' not found in Excel file.")
            return ""

        # Load the sheet into a DataFrame and filter based on model, tag type, and purpose (if provided)
        df = pd.read_excel(file_path, sheet_name=country)
        if purpose:
            filtered_df = df[(df['model'].astype(str) == model) & (df['type'].astype(str) == tag_type) & (df['purpose'].astype(str) == purpose)]
        else:
            filtered_df = df[(df['model'].astype(str) == model) & (df['type'].astype(str) == tag_type)]

        if filtered_df.empty:
            print(f"No tag {tag_type} found for model '{model}' in country '{country}'.")
            return ""

        # Return the first matching code
        return filtered_df['code'].values[0]
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return ""

# Function to replace old strings with new strings in content
def replace_strings(content, replacements):
    # Replace each old string with its corresponding new string
    for old, new in replacements.items():
        content = content.replace(old, new)
    return content

# Function to count occurrences of substrings in the content
def count_occurrences(content, substrings):
    occurrences = {}
    for substring in substrings:
        occurrences[substring] = content.count(substring)
    return occurrences

# Function to replace the opt-out link in the content

def replace_opt_out_link(content, new_href):
    # Regex pattern to match <a> tags with the specified aliases, excluding those with <%%view_email_url%%>
    pattern = r'(<a\s+[^>]*?href=")(?!<%%view_email_url%%>)([^"]*?)(".*?alias="Opt-out link".*?>)'

    # Replace only the href value for the matched <a> tags
    content = re.sub(pattern, r'\1' + new_href + r'\3', content)

    return content
def replace_opt_in_link(content, new_href):
    pattern = r'(<a\s+[^>]*?href=")([^"]*?)(".*?alias="Opt-in link".*?>)'
    content = re.sub(pattern, r'\1' + new_href + r'\3', content)
    return content



# Function to process the file based on user input
# Function to process the file based on user input
def process_file():
    # Retrieve the file paths stored in the labels
    file_path = getattr(file_label, 'file_path', None)
    replacements_path = getattr(replacements_label, 'file_path', None)
    image_style_path = getattr(image_style_label, 'file_path', None)

    # Check if all required files are selected
    if not file_path or not replacements_path or not image_style_path:
        result_label.config(text="Please select the necessary files.", foreground="red")
        return

    # Read the replacement mappings based on the selected country
    selected_country = country_combo.get()
    replacements = read_replacements_from_excel(replacements_path, selected_country)

    if not replacements:
        result_label.config(text="Failed to read replacements from Excel file.", foreground="red")
        return

    # Open the HTML file and read its content
    with open(file_path, 'r', encoding='utf-8') as file:
        content = file.read()
    content = remove_header(content)

    # Replace strings in the content based on the mappings
    content = replace_strings(content, replacements)

    # Read additional tags or styles based on user selections
    selected_model = model_combo.get()
    selected_purpose = purpose_combo.get() if country_combo.get() == "PRT" else None

    prepend_content = read_image_style_from_excel(image_style_path, selected_country, selected_model, 'purpose', selected_purpose)
    partyIdCrypted = read_image_style_from_excel(image_style_path, selected_country, 'generale', 'partyIdCrypted')
    prefrencecenter = read_image_style_from_excel(image_style_path, selected_country, selected_model, 'prefrence')
    append_text = read_image_style_from_excel(image_style_path, selected_country, 'generale', 'append_text')
    img_tag = read_image_style_from_excel(image_style_path, selected_country, selected_model, 'image')
    style_tag = read_image_style_from_excel(image_style_path, selected_country, selected_model, 'style')

    # Prepend and append content
    content = prepend_content + "\n" + partyIdCrypted + "\n" + append_text+ "\n" + content 
    
    # Use a regex to find and replace the media print style if it exists
    media_print_pattern = r'<style[^>]*>\s*@media\s+print\s*\{.*?\}\s*[^<]*</style>'
    if re.search(media_print_pattern, content):
        content = re.sub(media_print_pattern, style_tag, content, flags=re.DOTALL)
    else:
        head_end_pos = content.find("</head>")
        if head_end_pos != -1:
            content = content[:head_end_pos] + style_tag + content[head_end_pos:]

    # Insert image tag in the body
    body_tag_pos = content.find("<body")
    if body_tag_pos != -1:
        # Localise la fin de la balise <body>
        body_tag_end = content.find(">", body_tag_pos) + 1
        # Récupérer le contenu après la balise <body>
        before_body_content = content[body_tag_end:]

        # Expression régulière pour capturer <div id='_two50'></div> suivi d'une balise <img /> entièrement
        pattern = r"""<div id=['"]_two50['"]></div>\s*<img[^>]*>\s*&c=%%jobid%%[^>]*>"""

        # Remplacement dans le contenu après <body>
        before_body_content = re.sub(pattern, img_tag, before_body_content)

        # Mise à jour du contenu avec la nouvelle partie après <body>
        content = content[:body_tag_end] + before_body_content
    else:
        # Si pas de balise <body>, ajoute img_tag à la fin du contenu
        content += img_tag

    if prefrencecenter:
        content = replace_opt_out_link(content, prefrencecenter)
        content = replace_opt_in_link(content, prefrencecenter)

    # Save the modified content to a new file
    new_file_path = file_path.replace(".html", "_updated.html")
    with open(new_file_path, 'w', encoding='utf-8') as file:
        file.write(content)

    # Count occurrences of specific URLs and keywords in the content
    url_to_count = ['https://t.enews.myrenault.fr', 'images/']
    occurrences = count_occurrences(content, url_to_count)

    used_keywords = ['targetData', 'recipient.']
    keyword_counts = {keyword: count_occurrences(content, [keyword]) for keyword in used_keywords}
    keyword_line_numbers = {keyword: [] for keyword in used_keywords}

    content_lines = content.splitlines()
    for line_number, line in enumerate(content_lines, 1):
        for keyword in used_keywords:
            if keyword in line:
                keyword_line_numbers[keyword].append(line_number)

    # Display the result in the result label
    result_text = f"Updated file saved as: {new_file_path}\nOccurrences of images to update: {occurrences}\n"
    for keyword, count in keyword_counts.items():
        line_info = ", ".join(map(str, keyword_line_numbers[keyword]))
        result_text += f"{keyword}: {count[keyword]} occurrences (Lines: {line_info})\n"

    result_label.config(text=result_text, foreground="cornflower blue")



# Fonction pour supprimer header
def remove_header(content):
    # Expression régulière pour trouver le commentaire et la balise qui suit
    pattern = r'<!--\s*hide description\s*-->\s*<([^>]+)>(.*?)</\1>'
    content = re.sub(pattern, '', content, flags=re.DOTALL)
    return content

# Function to open file dialog and select HTML file
def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("HTML files", "*.html")])
    if file_path:
        # Display the selected file name in the label with green color
        file_label.config(text=f"Selected HTML file: {os.path.basename(file_path)}", foreground="green")
        setattr(file_label, 'file_path', file_path)
        enable_convert_button()

# Function to open file dialog and select replacement Excel file
def select_replacements_file():
    replacements_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if replacements_path:
        # Display the selected file name in the label with green color
        replacements_label.config(text=f"Selected dictionary file: {os.path.basename(replacements_path)}", foreground="green")
        setattr(replacements_label, 'file_path', replacements_path)

        # Populate the country dropdown with available sheet names
        sheet_names = get_sheet_names_from_excel(replacements_path)
        if sheet_names:
            country_combo['values'] = sheet_names
            country_combo.set('')  # Do not pre-select any sheet

      
        # Enable the convert button once everything is selected
        enable_convert_button()


# Function to open file dialog and select image/style Excel file
def select_image_style_file():
    image_style_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if image_style_path:
        # Display the selected file name in the label with green color
        image_style_label.config(text=f"Selected ADD file: {os.path.basename(image_style_path)}", foreground="green")
        setattr(image_style_label, 'file_path', image_style_path)

        # Check available sheet names
        sheet_names = pd.ExcelFile(image_style_path).sheet_names
        print("Available sheets:", sheet_names)

        # Read the 'list' sheet
        try:
            df = pd.read_excel(image_style_path, sheet_name='list')
        except ValueError as e:
            print("Error reading sheet 'list':", e)
            return

        # Print DataFrame columns for debugging
        print("DataFrame columns:", df.columns.tolist())

        # Strip spaces from column names to ensure consistency
        df.columns = df.columns.str.strip()

        # Check if the 'Purpose' column exists in the DataFrame
        if 'Purpose' in df.columns:
            purpose_list = df['Purpose'].dropna().unique().tolist()  # Use 'Purpose' with uppercase P
            purpose_combo['values'] = purpose_list
            purpose_combo.set('')  # Clear selection if needed
        else:
            print("Column 'Purpose' not found in the DataFrame.")

        enable_convert_button()





# Function to enable the convert button if all required fields are selected
def enable_convert_button():
    # Enable the convert button only if all required files and country are selected
    if (getattr(file_label, 'file_path', None) and
        getattr(replacements_label, 'file_path', None) and
        getattr(image_style_label, 'file_path', None) and
        country_combo.get()):
        convert_button.config(state='normal')
    else:
        convert_button.config(state='disabled')

# Function to show or hide the purpose selection based on the selected country
def update_purpose_visibility(*args):
    enable_convert_button()
    if country_combo.get() == "PRT":
        purpose_frame.grid(row=2, column=0, padx=10, pady=10, columnspan=2)
    else:
        purpose_frame.grid_forget()

# Set up the Tkinter GUI
root = tk.Tk()
root.title("HTML Content Replacer")

# Frame for file selection
file_frame = ttk.Frame(root, padding="15")
file_frame.grid(row=0, column=0, padx=10, pady=10)

# Label and button for selecting the HTML file
file_label = ttk.Label(file_frame, text="No file selected", foreground="red")
file_label.grid(row=0, column=0, padx=5, pady=5, columnspan=2)
select_file_button = ttk.Button(file_frame, text="Select HTML file", command=select_file)
select_file_button.grid(row=1, column=0, padx=5, pady=5)

# Label and button for selecting the replacements Excel file
replacements_label = ttk.Label(file_frame, text="No Excel file selected", foreground="red")
replacements_label.grid(row=2, column=0, padx=5, pady=5, columnspan=2)
select_replacements_button = ttk.Button(file_frame, text="Select dictionary Excel file", command=select_replacements_file)
select_replacements_button.grid(row=3, column=0, padx=5, pady=5)

# Label and button for selecting the image/style Excel file
image_style_label = ttk.Label(file_frame, text="No ADD Excel file selected", foreground="red")
image_style_label.grid(row=4, column=0, padx=5, pady=5, columnspan=2)
select_image_style_button = ttk.Button(file_frame, text="Select ADD parameters Excel file", command=select_image_style_file)
select_image_style_button.grid(row=5, column=0, padx=5, pady=5)

# Frame for additional options (country, model, purpose)
options_frame = ttk.Frame(root, padding="10")
options_frame.grid(row=1, column=0, padx=10, pady=10)

# Country selection
country_label = ttk.Label(options_frame, text="Select country:")
country_label.grid(row=0, column=0, padx=5, pady=5)
country_combo = ttk.Combobox(options_frame)
country_combo.grid(row=0, column=1, padx=5, pady=5)
country_combo.bind("<<ComboboxSelected>>", update_purpose_visibility)

# Model selection
model_label = ttk.Label(options_frame, text="Select brand:")
model_label.grid(row=1, column=0, padx=5, pady=5)
model_combo = ttk.Combobox(options_frame, values=["Dacia", "Renault"])
model_combo.grid(row=1, column=1, padx=5, pady=5)

# Purpose selection (conditionally shown based on country)
purpose_frame = ttk.Frame(root, padding="10")
purpose_label = ttk.Label(purpose_frame, text="Select purpose:")
purpose_label.grid(row=0, column=0, padx=5, pady=5)
purpose_combo = ttk.Combobox(purpose_frame)
purpose_combo.grid(row=0, column=1, padx=5, pady=5)

# Convert button
convert_button = ttk.Button(root, text="Convert", command=process_file, state='disabled')
convert_button.grid(row=6, column=0, padx=10, pady=10)

# Result label
result_label = ttk.Label(root, text="", foreground="green")
result_label.grid(row=7, column=0, padx=10, pady=10)

# Start the Tkinter main loop
root.mainloop()
