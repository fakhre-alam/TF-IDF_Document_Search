import streamlit as st
import pandas as pd
import pickle
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import numpy as np
import pyxlsb

st.set_page_config(layout="wide")
st.title("Enhanced Document Question Answering App")

question = st.text_input("Enter your question", "")
threshold = st.sidebar.slider("Matching Threshold", 0.0, 1.0, 0.1, 0.01)

# Load processed data
with open(r'C:\Users\fakhr\Videos\tfidf_search\New_tfidf\processed_pdf_data.pkl', 'rb') as f:
    processed_pdf_data = pickle.load(f)

with open(r'C:\Users\fakhr\Videos\tfidf_search\New_tfidf\processed_xlsx_data.pkl', 'rb') as f:
    processed_xlsx_data = pickle.load(f)

with open(r'C:\Users\fakhr\Videos\tfidf_search\New_tfidf\processed_xlsb_data.pkl', 'rb') as f:
    processed_xlsb_data = pickle.load(f)

def process_question(question, threshold, processed_data):
    results = []
    for data in processed_data:
        sentences = data['sentences']
        vectorizer = TfidfVectorizer(stop_words='english')
        tfidf_matrix = vectorizer.fit_transform(sentences)
        question_tfidf = vectorizer.transform([question])
        similarities = cosine_similarity(question_tfidf, tfidf_matrix)
        max_sim_index = np.argmax(similarities)
        max_similarity_score = similarities[0, max_sim_index]
        if max_similarity_score >= threshold:
            most_similar_sentence = sentences[max_sim_index].strip()
            matching_percentage = max_similarity_score * 100
            if 'page_numbers' in data:  # PDF data
                page_num = data['page_numbers'][max_sim_index]
                file_name = data['file_name']
                results.append({
                    'Results': most_similar_sentence,
                    'Matching Percentage': matching_percentage,
                    'Document Page': page_num,
                    'Document Name': file_name,
                    'File Path': data['file_path'],
                    'Page Number': page_num
                })
            elif 'sheet_names' in data:  # XLSX and XLSB data
                sheet_name = data['sheet_names'][max_sim_index]
                row_number = data['row_numbers'][max_sim_index]
                file_name = data['file_name']
                results.append({
                    'Results': most_similar_sentence,
                    'Matching Percentage': matching_percentage,
                    'Sheet Name': sheet_name,
                    'Row Number': row_number,
                    'File Name': file_name,
                    'File Path': data['file_path']
                })
    return results

def extract_matching_rows(result):
    file_path = result['File Path']
    if file_path.endswith('.xlsx'):
        try:
            df_xlsx = pd.read_excel(file_path, sheet_name=result['Sheet Name'], engine='openpyxl')
            matching_row_df = df_xlsx.iloc[[result['Row Number'] - 1]]  # Row numbers are 1-indexed
            return matching_row_df
        except Exception as e:
            st.error(f"Error reading Excel file: {e}")
            return None
    elif file_path.endswith('.xlsb'):
        try:
            with pyxlsb.open_workbook(file_path) as wb:
                with wb.get_sheet(result['Sheet Name']) as sheet:
                    rows = [row for row in sheet.rows()]
                    headers = [r.v for r in rows[0]]
                    matching_row = [r.v for r in rows[result['Row Number']]]
                    matching_row_df = pd.DataFrame([matching_row], columns=headers)
                    return matching_row_df
        except Exception as e:
            st.error(f"Error reading XLSB file: {e}")
            return None
    elif file_path.endswith('.pdf'):
        # PDF extraction logic to be implemented based on your specific method
        st.warning("PDF extraction logic not implemented.")
        return None
    else:
        st.warning("Unsupported file format")
        return None

def make_column_names_unique(df):
    cols = pd.Series(df.columns)
    for dup in cols[cols.duplicated()].unique():
        cols[cols[cols == dup].index.values.tolist()] = [str(dup) + '_' + str(i) if i != 0 else str(dup) for i in range(sum(cols == dup))]
    df.columns = cols
    return df

if "results" not in st.session_state:
    st.session_state.results = []

if st.button("Get Results"):
    if question:
        with st.spinner('Processing...'):
            results = process_question(question, threshold, processed_xlsx_data + processed_pdf_data + processed_xlsb_data)
            if results:
                st.session_state.results = results
                st.write("Results:")
                df = pd.DataFrame(results)
                
                # Conditionally handle the 'Row Number' and 'Sheet Name' columns
                if 'Row Number' in df.columns:
                    df["Row Number"] = df["Row Number"].fillna(0).astype(int)
                if 'Matching Percentage' in df.columns:
                    df["Matching Percentage"] = df["Matching Percentage"].fillna(0).astype(int)
                    df["Matching Percentage"] = df["Matching Percentage"].astype(str) + '%'
                
                df = df.sort_values(by="Matching Percentage", ascending=False, key=lambda col: col.str.rstrip('%').astype(int))
                st.write(df)
            else:
                st.warning("No results found.")

if st.button("Get Matching Rows"):
    matching_rows = []
    for result in st.session_state.results:
        matching_row_df = extract_matching_rows(result)
        if matching_row_df is not None:
            matching_row_df = make_column_names_unique(matching_row_df)
            matching_rows.append({
                'File Name': result['File Name'],
                'Sheet Name': result.get('Sheet Name', 'N/A'),
                'Matching Row Data': matching_row_df
            })
    if matching_rows:
        st.write("Matching Rows:")
        for match in matching_rows:
            st.write(f"Matching Row from {match['File Name']} - Sheet '{match['Sheet Name']}':")
            st.dataframe(match['Matching Row Data'])
            st.download_button(
                label="Download Matching Row as CSV",
                data=match['Matching Row Data'].to_csv(index=False).encode('utf-8'),
                file_name=f"{match['File Name']}_{match['Sheet Name']}_matching_row.csv",
                mime="text/csv"
            )
    else:
        st.warning("No matching rows found.")
