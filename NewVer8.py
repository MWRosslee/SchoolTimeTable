import pandas as pd
import os
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages


# Function to create a half-hour interval timetable template
def create_timetable_template():
    times = pd.date_range("07:30", "14:30", freq="30min").strftime('%H:%M')
    timetable = pd.DataFrame(index=times, columns=["Mon", "Tue", "Wed", "Thu", "Fri"])
    return timetable


# Function to create a timetable for each entity type (Grade, Teacher, Class Teacher)
def create_timetables(df, entity_column):
    timetables = {}
    assist_timetables = {}  # New dictionary for assist subjects
    for entity in df[entity_column].unique():
        entity_df = df[df[entity_column] == entity]
        assist_filter = entity_df['Subject'].str.contains('Assist', case=False, na=False)
        timetables[entity] = entity_df[~assist_filter]
        assist_timetables[entity] = entity_df[assist_filter]
    return timetables, assist_timetables

# Function to fill the timetable for each entity
def fill_timetable(df, entity, entity_type):
    timetable = create_timetable_template()
    for _, row in df.iterrows():
        for day in ["Mon", "Tue", "Wed", "Thu", "Fri"]:
            start_time = row[f'{day}Start']
            end_time = row[f'{day}End']
            if start_time != '00:00' and end_time != '00:00':
                start = pd.to_datetime(start_time)
                end = pd.to_datetime(end_time)
                while start < end:
                    time_str = start.strftime('%H:%M')
                    cell_value = ''
                    if entity_type == 'Grade':
                        cell_value = f"{row['Subject']} ({row['Teacher']})"
                    elif entity_type == 'Teacher':
                        cell_value = f"{row['Subject']} ({row['Grade']})"
                    else:  # Class Teacher
                        cell_value = f"{row['Subject']} ({row['Grade']})"
                    # Combine data for the same time slot, separated by a semicolon
                    if pd.notna(timetable.at[time_str, day]):
                        timetable.at[time_str, day] += f"; {cell_value}"
                    else:
                        timetable.at[time_str, day] = cell_value
                    start += pd.Timedelta(minutes=30)
    # Replace NaN values with an empty string or a specific placeholder
    timetable.fillna('', inplace=True)

    return timetable


# Function to generate a unique filename by appending numbers if the file already exists
def generate_unique_filename(filepath):
    base, extension = os.path.splitext(filepath)
    counter = 1
    while os.path.exists(filepath):
        filepath = f"{base}_{counter}{extension}"
        counter += 1
    return filepath


# Function to display and save timetable as a PNG image
def wrap_text(text, max_width=20):
    """Wrap text for a given maximum width."""
    if len(text) <= max_width:
        return text
    else:
        # Find a space near the middle to split the text
        split_index = text.rfind(' ', 0, max_width)
        if split_index == -1:  # No space found, force split
            return text[:max_width] + '\n' + wrap_text(text[max_width:], max_width)
        return text[:split_index] + '\n' + wrap_text(text[split_index+1:], max_width)


def save_timetable_as_png(timetable, title, filename):
    fig, ax = plt.subplots(figsize=(12, 10))
    ax.axis('tight')
    ax.axis('off')

    # Create table and wrap text in cells
    cell_text = timetable.apply(lambda x: x.map(lambda y: wrap_text(y) if isinstance(y, str) else y))
    table = ax.table(cellText=cell_text.values, colLabels=timetable.columns, rowLabels=timetable.index, cellLoc='center', loc='center')

    # Adjust cell height if needed
    cell_height = 0.06  # Example height, adjust as needed
    for pos, cell in table.get_celld().items():
        cell.set_height(cell_height)

    plt.title(title)
    plt.savefig(filename, bbox_inches='tight', pad_inches=0.1)
    plt.close()


# Function to save a timetable as a PDF page
def save_timetable_to_pdf(timetable, title, pdf):
    fig, ax = plt.subplots(figsize=(11, 8.5))  # Landscape format
    ax.axis('tight')
    ax.axis('off')
    ax.table(cellText=timetable.values, colLabels=timetable.columns, rowLabels=timetable.index, cellLoc='center',
             loc='center')
    plt.title(title)
    pdf.savefig(fig, bbox_inches='tight')
    plt.close()


# Mapping of first letters to full category names
category_mapping = {
    'T': 'Teacher',
    'G': 'Grade',
    'C': 'Class Teacher',
    'S': 'Subject',
    'L': 'Classroom',
    'F': 'Full Timetable'
}

# Read CSV file
csv_file_path = r'2024 timetable.csv'  # Replace with the actual CSV file path
print("Reading CSV file...")
df = pd.read_csv(csv_file_path)

# User input for category and output format selection
category_letter = input(
    "Select category (T: Teacher, G: Grade, C: Class Teacher, S: Subject, L: Classroom, F: Full Timetable): ").strip().upper()
output_format_letter = input("Select output format (P: PNG, E: Excel, PDF): ").strip().upper()

# Map the first letter to the full category and output format name
category = category_mapping.get(category_letter, None)
output_format = {'P': 'PNG', 'E': 'Excel', 'PDF': 'PDF'}.get(output_format_letter, None)

if category and output_format:
    print(f"Selected category: {category}, Output format: {output_format}")

    if category in ["Teacher", "Grade", "Class Teacher", "Subject", "Classroom"]:
        entity_timetables, assist_entity_timetables = create_timetables(df, category)

        # Loop to handle main timetables
        for entity, entity_df in entity_timetables.items():
            timetable = fill_timetable(entity_df, entity, category)
            if output_format == 'PNG':
                png_folder_path = f'C:\\data\\to\\save\\{category.lower().replace(" ", "_")}\\png'  # Replace with actual directory
                if not os.path.exists(png_folder_path):
                    os.makedirs(png_folder_path)
                png_filename = f'{png_folder_path}\\{category}_{entity}.png'
                png_filename = generate_unique_filename(png_filename)
                save_timetable_as_png(timetable, f"{category} - {entity} Timetable", png_filename)
            elif output_format == 'Excel':
                excel_folder_path = f'C:\\data\\to\\save\\{category.lower().replace(" ", "_")}\\excel'  # Replace with actual directory
                if not os.path.exists(excel_folder_path):
                    os.makedirs(excel_folder_path)
                excel_filename = f'{excel_folder_path}\\{category}_{entity}.xlsx'
                excel_filename = generate_unique_filename(excel_filename)
                timetable.to_excel(excel_filename)

        # Loop to handle assist timetables
        for entity, entity_df in assist_entity_timetables.items():
            if not entity_df.empty:  # Check if there is data for assist subjects
                assist_timetable = fill_timetable(entity_df, entity, category)
                # Handle saving for assist subjects (PNG, Excel, PDF)
                if output_format == 'PNG':
                    png_folder_path = f'C:\\data\\to\\save\\{category.lower().replace(" ", "_")}\\png'  # Replace with actual directory
                    if not os.path.exists(png_folder_path):
                        os.makedirs(png_folder_path)
                    png_filename = f'{png_folder_path}\\{category}_{entity}.png'
                    png_filename = generate_unique_filename(png_filename)
                    save_timetable_as_png(assist_timetable, f"{category} - {entity} Timetable", png_filename)
                elif output_format == 'Excel':
                    excel_folder_path = f'C:\\data\\to\\save\\{category.lower().replace(" ", "_")}\\excel'  # Replace with actual directory
                    if not os.path.exists(excel_folder_path):
                        os.makedirs(excel_folder_path)
                    excel_filename = f'{excel_folder_path}\\{category}_{entity}.xlsx'
                    excel_filename = generate_unique_filename(excel_filename)
                    assist_timetable.to_excel(excel_filename)
            # Add similar code for PDF if needed

    elif category == "Full Timetable":
        full_category = input("Select full timetable category (Grade or Teacher): ").strip()
        print(f"Selected full timetable category: {full_category}")
        if full_category in ["Grade", "Teacher"]:
            entity_timetables = create_timetables(df, full_category)
            output_pdf_path = rf'C:\data\to\save\full_timetable_{full_category.lower()}.pdf'  # Replace with actual directory
            output_pdf_path = generate_unique_filename(output_pdf_path)
            print(f"Generating PDF: {output_pdf_path}")
            with PdfPages(output_pdf_path) as pdf:
                for entity, entity_df in entity_timetables.items():
                    timetable = fill_timetable(entity_df, entity, full_category)
                    save_timetable_to_pdf(timetable, f"{full_category} - {entity}", pdf)
            print("PDF generation complete.")
        else:
            print("Invalid full timetable category selection.")
else:
    print("Invalid category or output format selection.")

print("Script execution finished.")
