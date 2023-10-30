import os
import pandas as pd
from azure.ai.formrecognizer import FormRecognizerClient
from azure.core.credentials import AzureKeyCredential
from azure.ai.formrecognizer import DocumentAnalysisClient

def recognize_tables_from_form_and_save_to_excel(endpoint, api_key, document_url, output_file):
    # Initialize the Form Recognizer client
    document_analysis_client = DocumentAnalysisClient(endpoint=endpoint, credential=AzureKeyCredential(api_key))

    # Start the form recognition process
    poller = document_analysis_client.begin_analyze_document_from_url("prebuilt-layout", document_url)
    result = poller.result()

    tables = result.tables
    if tables:
        # Create a new Excel writer
        with pd.ExcelWriter(output_file) as writer:
            for table_idx, table in enumerate(tables):
                # Organize cells into rows and columns
                organized_data = [[None for _ in range(table.column_count)] for _ in range(table.row_count)]
                for cell in table.cells:
                    for row_idx in range(cell.row_index, cell.row_index + cell.row_span):
                        for col_idx in range(cell.column_index, cell.column_index + cell.column_span):
                            organized_data[row_idx][col_idx] = cell.content

                
                # Convert data into DataFrame
                df = pd.DataFrame(organized_data)
                
                # Write DataFrame to Excel sheet
                df.to_excel(writer, sheet_name=f"Table_{table_idx + 1}", index=False)
    else:
        print("No tables found in the document!")

if __name__ == "__main__":
    # It's not safe to hard-code sensitive information like API keys in the script.
    # Ideally, these should be fetched securely from environment variables or a secure configuration.
    YOUR_ENDPOINT = ''      # Form Recognizer endpoint
    YOUR_API_KEY = ''       # Form Recognizer API key

    # Document URL
    DOCUMENT_URL = 'https://www.gov.br/saude/pt-br/composicao/sectics/daf/rename/20210367-rename-2022_final.pdf'

    # Determine the directory of the current script
    script_dir = os.path.dirname(os.path.abspath(__file__))

    # Construct the absolute path for the output file
    OUTPUT_FILE = os.path.join(script_dir, "output.xlsx")

    recognize_tables_from_form_and_save_to_excel(YOUR_ENDPOINT, YOUR_API_KEY, DOCUMENT_URL, OUTPUT_FILE)
