
import os
from flask import Flask, render_template, request, redirect, url_for, send_file
from doctest1 import *
app = Flask(__name__)

@app.route('/')
def list_files():
    folder_path = r'word_file'  # Replace with the actual folder path
    # Get a list of all files in the folder
    files = os.listdir(folder_path)

    return render_template('index.html', files=files)

@app.route('/delete', methods=['POST'])
def delete_file():
    folder_path = r'word_file'  # Replace with the actual folder path

    # Get the filename from the request
    filename = request.form['filename']

    # Construct the full file path
    file_path = os.path.join(folder_path, filename)

    # Delete the file
    os.remove(file_path)

    # Redirect back to the file list
    return redirect(url_for('list_files'))

@app.route('/fix', methods=['POST'])
def fix_file():
    folder_path = r'word_file'  # Replace with the actual folder path
    # Get the filename from the request
    filename = request.form['filename']

    # Construct the full file path
    file_path = os.path.join(folder_path, filename)
    # Define the API endpoint for code generation
    api_url = "https://3c92-103-253-89-37.ngrok-free.app/generate_code?max_length=512"

    #for i, file in enumerate(WordReplacer.docx_list(filedir), start=1):
        #print(f"{i} Processing file: {file}")

        # Load the Word document
        #word_replacer = WordReplacer(filedir2)
    word_replacer = WordReplacer(file_path)
        # Extract all paragraphs from the document
    paragraphs = [paragraph.text for paragraph in word_replacer.docx.paragraphs]
    print(paragraphs[1])
    table_texts = []
    for table in word_replacer.docx.tables:
            for row in table.rows:
                row_text = [cell.text for cell in row.cells]
                for text in row_text:
                    table_texts.append(text)
    
        # Create a list of prompts
    prompts_list = [f"Correct English grammar in the following text keep curly brackets keep it in one paragraph: {paragraph}\nHere is the corrected version: " for paragraph in paragraphs]
        # table still testing
    prompts_list_table = [f"Correct only grammar in the following text if needed do not define or add information keep it in one paragraph: {table_text}.\nHere is the corrected version: " for table_text in table_texts]
        
    all_prompts_list = prompts_list + prompts_list_table
        
        # Define API parameters
    api_params = {'prompts': all_prompts_list}
        
        # Send a GET request to the API
    response = requests.get(api_url, params=api_params)
        
        # Check the status code and response content
    if response.status_code == 200:
            corrected_paragraphs = response.json()
            
            all_text = paragraphs + table_texts

            # Replace original paragraphs with corrected paragraphs
            for i, (original, corrected) in enumerate(zip(all_text, corrected_paragraphs), start=1):
                word_replacer.replace_in_paragraph(original, corrected)
                print(f"Paragraph {i}: Replaced successfully!")
                
            # Save the document with replaced paragraphs
            output_filepath = os.path.join(folder_path, "document_updated.docx")
            #output_filepath = f"document_updated.docx"
            word_replacer.save(output_filepath)
            print(f"Saved updated document to: {output_filepath}\n")
    else:
            print("Failed to retrieve corrections. Status code:", response.status_code)
    

    

    # Redirect back to the file list
    return redirect(url_for('list_files'))

@app.route('/download', methods=['POST'])
def download_file():
    folder_path = r'word_file'  # Replace with the actual folder path

    # Get the filename from the request
    filename = request.form['filename']

    # Construct the full file path
    file_path = os.path.join(folder_path, filename)

    # Send the file as a response for download
    return send_file(file_path, as_attachment=True)

if __name__ == '__main__':
 app.run(host='0.0.0.0')
