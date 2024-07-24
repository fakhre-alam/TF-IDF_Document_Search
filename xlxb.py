# import os
# import pickle
# from pyxlsb import open_workbook as open_xlsb

# def preprocess_xlsb_files(directory_path):
    # processed_data = []

    # for file_name in os.listdir(directory_path):
        # file_path = os.path.join(directory_path, file_name)
        # if file_path.endswith('.xlsb'):
            # with open_xlsb(file_path) as wb:
                # sentences = []
                # sheet_names = []
                # row_numbers = []
                # for sheet_name in wb.sheets:
                    # with wb.get_sheet(sheet_name) as sheet:
                        # for row_index, row in enumerate(sheet.rows()):
                            # row_data = [c.v for c in row]
                            # text_columns = [str(c) for c in row_data if c is not None]
                            # for text_column in text_columns:
                                # for sentence in text_column.split('.'):
                                    # sentence = sentence.strip()
                                    # if sentence:
                                        # sentences.append(sentence)
                                        # sheet_names.append(sheet_name)
                                        # row_numbers.append(row_index + 1)  # xlsb rows are 1-indexed
                # processed_data.append({
                    # 'file_name': file_name,
                    # 'file_path': file_path,
                    # 'sentences': sentences,
                    # 'sheet_names': sheet_names,
                    # 'row_numbers': row_numbers
                # })

    # with open(r'C:\Users\fakhr\Videos\tfidf_search\New_tfidf\processed_xlsb_data.pkl', 'wb') as f:
        # pickle.dump(processed_data, f)

# if __name__ == "__main__":
    # directory_path = r'C:\Users\fakhr\Videos\tfidf_search\output'
    # preprocess_xlsb_files(directory_path)


import os
import pickle
from pyxlsb import open_workbook as open_xlsb

def preprocess_xlsb_files(directory_path):
    processed_data = []

    for file_name in os.listdir(directory_path):
        file_path = os.path.join(directory_path, file_name)
        if file_path.endswith('.xlsb'):
            try:
                with open_xlsb(file_path) as wb:
                    sentences = []
                    sheet_names = []
                    row_numbers = []
                    for sheet_name in wb.sheets:
                        with wb.get_sheet(sheet_name) as sheet:
                            for row_index, row in enumerate(sheet.rows()):
                                row_data = [c.v for c in row]
                                text_columns = [str(c) for c in row_data if c is not None]
                                for text_column in text_columns:
                                    for sentence in text_column.split('.'):
                                        sentence = sentence.strip()
                                        if sentence:
                                            sentences.append(sentence)
                                            sheet_names.append(sheet_name)
                                            row_numbers.append(row_index + 1)  # xlsb rows are 1-indexed
                    if sentences:
                        processed_data.append({
                            'file_name': file_name,
                            'file_path': file_path,
                            'sentences': sentences,
                            'sheet_names': sheet_names,
                            'row_numbers': row_numbers
                        })
            except Exception as e:
                print(f"Error processing file {file_path}: {e}")

    with open(r'C:\Users\fakhr\Videos\tfidf_search\New_tfidf\processed_xlsb_data.pkl', 'wb') as f:
        pickle.dump(processed_data, f)

if __name__ == "__main__":
    directory_path = r'C:\Users\fakhr\Videos\tfidf_search\output'
    preprocess_xlsb_files(directory_path)
