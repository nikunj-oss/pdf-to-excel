import pandas as pd
import re
import os
from PyPDF2 import PdfReader
from tkinter import Tk, filedialog


def extract_data_from_pdf(pdf_path):
    reader = PdfReader(pdf_path)
    text = ""
    for page in reader.pages:
        page_text = page.extract_text()
        if page_text:
            text += page_text + "\n"
    return text


def process_data(text):
    lines = text.splitlines()
    data = []

    pattern = re.compile(r'(\d{5})\s+([A-Z0-9/@&\-\s]+?)\s+((?:\d+\.\d{5}\s+)+)')

    for line in lines:
        match = pattern.search(line)
        if match:
            index = match.group(1)
            location = match.group(2).strip()
            distances = match.group(3).strip().split()
            data.append([index, location] + distances)

    if data:
        df = pd.DataFrame(data)
        df.columns = ['Index', 'Location'] + [f'Distance_{i+1}' for i in range(len(df.columns) - 2)]
        return df
    else:
        return pd.DataFrame(columns=['Index', 'Location'])


def convert_to_long_format(df):
    if df.empty:
        return pd.DataFrame()

    delivery_ids = df['Index'].tolist()
    delivery_names = df['Location'].tolist()

    long_data = []
    for _, row in df.iterrows():
        rec_id = row['Index']
        rec_name = row['Location']
        distances = row[2:]

        for i, dist in enumerate(distances):
            try:
                dist_value = float(dist)
            except ValueError:
                continue
            long_data.append({
                "Rec Loc": rec_name,
                "Rec LocID": rec_id,
                "Del Loc": delivery_names[i],
                "Del LocID": delivery_ids[i],
                "Distance": dist_value
            })

    return pd.DataFrame(long_data)


def select_file(prompt):
    root = Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title=prompt)
    root.update()
    return file_path


def select_output_path():
    root = Tk()
    root.withdraw()
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    root.update()
    return file_path


def main():
    print("Select the PDF file to convert:")
    pdf_path = select_file("Select PDF file")

    if not pdf_path or not os.path.exists(pdf_path):
        print("Invalid PDF path.")
        return

    print("Select where to save the output Excel file:")
    output_excel_path = select_output_path()
    if not output_excel_path:
        print("No output path selected.")
        return

    text = extract_data_from_pdf(pdf_path)

    if not text.strip():
        print("No text extracted from the PDF.")
        return

    df_matrix = process_data(text)

    if df_matrix.empty:
        print("No matrix data found in the PDF.")
        return

    df_long = convert_to_long_format(df_matrix)
    df_long.to_excel(output_excel_path, index=False)
    print(f"Excel file saved to: {output_excel_path}")


if __name__ == "__main__":
    main()
