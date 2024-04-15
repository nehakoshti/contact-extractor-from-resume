from flask import Flask, request, render_template,redirect
from flask import Flask, render_template, send_file
# from io import BytesIO

import pdfplumber
import re
import xlwt
import os
import spacy

import Levenshtein

app = Flask(__name__)


def process_email(email):
    # Remove the domain part (e.g., "@gmail.com")
    email_without_domain = email.split('@')[0]
    # Remove numbers from the email
    email_without_numbers = re.sub(r'\d', '', email_without_domain)
    # Replace "." with space
    email_processed = email_without_numbers.replace('.', ' ')
    return email_processed


def similarity_percentage(string1, string2):
    # ===========================if cases upper n lower - balance result===================================
    string1 = string1.lower()
    string2 = string2.lower()
    # Calculate Levenshtein distance
    distance = Levenshtein.distance(string1, string2)    
    # Calculate maximum length
    max_length = max(len(string1), len(string2))    
    # Calculate similarity percentage
    similarity1 = (max_length - distance) / max_length * 100    
    # ============================if sequence changed - balance result====================================
    # Convert strings to lowercase and split into words
    words1 = set(string1.lower().split())
    words2 = set(string2.lower().split())   
    common_words = words1.intersection(words2)
    total_words = len(words1.union(words2))
    similarity2 = len(common_words) / total_words * 100  

     
    return max(similarity1,similarity2)


from fuzzywuzzy import process
def find_similar_words(sentence, target_string, threshold=80):
    words = sentence.split()
    similar_words = []
    for word in words:
        similarity_score = process.extractOne(word, [target_string])[1]
        if similarity_score >= threshold:
            similar_words.append(word)
    return similar_words

def extract_info_from_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        emails = set()
        phone_numbers = set()
        names = set()
        for page in pdf.pages:
            text = page.extract_text()

            # Regular expressions to find email, phone number, and name patterns
            email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
            phone_number_pattern = r'\b(?:\+\d{1,3}\s)?\(?\d{3}\)?[-.\s]?\d{1,3}[-.\s]?\d{1,4}[-.\s]?\d{1,4}\b'            
            # name_pattern = r'\b(?:FF|FT)-?\s*\d+\s+(.?)\s+\b(?:\d+\s(?:pcs|pack|kg|L))\b'
            emails.update(re.findall(email_pattern, text))
            phone_numbers.update(re.findall(phone_number_pattern, text))
            filtered_list=[]
            PhoneNumber=[]
            for num in phone_numbers:
                if "." not in num:
                    update_num= num.replace(' ', '')
                    if len(update_num)==10:
                        PhoneNumber.append(update_num)
           
            # =================================================
            nlp = spacy.load("en_core_web_sm")
            # Function to extract names from resume PDF
            names = []
            doc = nlp(text)
            for ent in doc.ents:
                if ent.label_ == "PERSON":
                    names.append(ent.text)
            # print( names)
            # =================================================         
            
            # print("---------------------------")
            # print(emails, PhoneNumber, names)            
            # print()      
            first_line = text.strip().split('\n')[0]
            break      
    print("======================================")
    print(list(emails), PhoneNumber, names[0])
    print("============================")

    # similarity btw email and names
    emails=list(emails)
    emails=emails[0]
    processed_email = process_email(emails)
    print("Processed email:", processed_email)
    percentage = similarity_percentage(str(names[0]), str(processed_email))
    print( "Percentage of similarity:", names[0], processed_email,"::", percentage )
    print()
    print("============================")

    print("1st line:", first_line)
    print("Processed email:", processed_email)
    similar_words = find_similar_words(str(text), str(processed_email))
    print("Similar words:", similar_words)
    similar_words = [s for s in similar_words if "@" not in s]
    print(similar_words)
    word_string = ' '.join(similar_words[:2])
    print("word_string : ",word_string)
    percentage2 = similarity_percentage(word_string, processed_email)
    print( "Percentage of similarity:", word_string, processed_email,"::", percentage2 )
    print()
    print("============================")
    if percentage < percentage2:
        if percentage2 >=20:
            main_name = word_string
        else:
            main_name=processed_email
    else:
        if percentage >=20:
            main_name=names[0]
        else:
            main_name=processed_email
        

    print(emails, PhoneNumber, main_name  )
    return emails, PhoneNumber, main_name    



@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        files = request.files.getlist('files')
        if files:
            # Create a workbook
            workbook = xlwt.Workbook()
            sheet = workbook.add_sheet('Resume Info')
            # Write headers
            sheet.write(0, 0, 'Name')
            sheet.write(0, 1, 'Email')
            sheet.write(0, 2, 'Phone Number')
            # sheet.write(0, 3, 'Other Details')
            row = 1
            for file_idx, file in enumerate(files):
                if file.filename.endswith('.pdf'):
                    # Save the uploaded file temporarily
                    temp_path = f"temp_resume_{file_idx}.pdf"
                    file.save(temp_path)
                    # Extract information from the PDF
                    extracted_emails, extracted_phone_numbers, extracted_names = extract_info_from_pdf(temp_path)

                    print(extracted_emails, extracted_phone_numbers, extracted_names)                                        
                    print("Extracted emails:",extracted_emails)                  
                    print("\nExtracted phone numbers:")
                    for phone_number in extracted_phone_numbers:
                        print(phone_number)
                    print("\nExtracted name:")
                    print(extracted_names)
                    print()
                    print("---------------------------------------")

                    # Write data                    
                    sheet.write(row, 0, extracted_names)
                    sheet.write(row, 1, extracted_emails)
                    p=2
                    for phone_number in extracted_phone_numbers:
                        sheet.write(row, p, phone_number)
                        p+=1
                    # sheet.write(row, 2, extracted_phone_numbers)
                    row+=1

                    # Delete the temporary file
                    os.remove(temp_path)
            # Save the workbook
            workbook.save('resume_info.xls')

             # Redirect to the download page
            return redirect('/download')                            
            return 'Resume information extracted and saved successfully!'
        
    return render_template('index.html')

@app.route('/download')
def download_page():
    return render_template('download.html')

@app.route('/download_file')
def download_file():
    # Specify the path to the XLS file
    xls_file_path = 'resume_info.xls'
    
    # Send the file as an attachment with a specific filename
    return send_file(xls_file_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
