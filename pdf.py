import os
import pickle
from PyPDF2 import PdfReader

def preprocess_pdf_files(directory_path):
    processed_data = []

    for file_name in os.listdir(directory_path):
        file_path = os.path.join(directory_path, file_name)
        if file_path.endswith('.pdf'):
            with open(file_path, 'rb') as f:
                pdf = PdfReader(f)
                sentences = []
                page_numbers = []
                for page_num in range(len(pdf.pages)):
                    page = pdf.pages[page_num]
                    text = page.extract_text()
                    for sentence in text.split('.'):
                        sentence = sentence.strip()
                        if sentence:
                            sentences.append(sentence)
                            page_numbers.append(page_num + 1)  # PDF pages are 1-indexed
                processed_data.append({
                    'file_name': file_name,
                    'file_path': file_path,
                    'sentences': sentences,
                    'page_numbers': page_numbers
                })

    with open(r'C:\Users\fakhr\Videos\tfidf_search\New_tfidf\processed_pdf_data.pkl', 'wb') as f:
        pickle.dump(processed_data, f)

if __name__ == "__main__":
    directory_path = r'C:\Users\fakhr\Videos\tfidf_search\output'
    preprocess_pdf_files(directory_path)
