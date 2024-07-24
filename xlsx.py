import os
import pandas as pd
import pickle

def preprocess_xlsx_files(directory_path):
    processed_data = []

    for file_name in os.listdir(directory_path):
        file_path = os.path.join(directory_path, file_name)
        
        if file_path.endswith('.xlsx'):
            try:
                xls = pd.ExcelFile(file_path, engine='openpyxl')
                sentences = []
                sheet_names = []
                row_numbers = []

                for sheet_name in xls.sheet_names:
                    df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
                    for row_index, row_data in df.iterrows():
                        for cell in row_data:
                            if pd.notnull(cell) and isinstance(cell, str):
                                for sentence in cell.split('.'):
                                    sentence = sentence.strip()
                                    if sentence:
                                        sentences.append(sentence)
                                        sheet_names.append(sheet_name)
                                        row_numbers.append(row_index + 1)  # Excel rows are 1-indexed

                processed_data.append({
                    'file_name': file_name,
                    'file_path': file_path,
                    'sentences': sentences,
                    'sheet_names': sheet_names,
                    'row_numbers': row_numbers
                })
            except Exception as e:
                print(f"Error processing {file_path} as .xlsx: {e}")
        else:
            print(f"Skipping unsupported file type: {file_path}")

    with open(r'C:\Users\fakhr\Videos\tfidf_search\New_tfidf\processed_xlsx_data.pkl', 'wb') as f:
        pickle.dump(processed_data, f)

if __name__ == "__main__":
    directory_path = r'C:\Users\fakhr\Videos\tfidf_search\output'
    preprocess_xlsx_files(directory_path)
