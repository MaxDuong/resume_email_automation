import numpy as np
import pandas as pd
import json
import yagmail
from word_to_pdf_converter import replace_company_name, replace_company_name_original, doc_to_pdf_converter

def email_authentication():
    # Opening JSON file
    f = open('authen.json',)

    # returns JSON object as a dictionary
    authen = json.load(f)
    return authen

def website_clean(website_url):
    website_array = ['indeed', 'glassdoor']
    for website in website_array:
        if website in website_url:
            return website.title()


def send_resume_to_recruiter(job_title, company_name, email, website):
    
    # Authentication
    authen = email_authentication()
    #yag = yagmail.SMTP({'maxduong97@gmail.com': 'Max Duong'}, 'Morningschool1!')
    yag = yagmail.SMTP({authen['email']: authen['user_name_email']}, authen['password'])
    
    body = """\
            <p> Dear {} Team,

                I'm very interested in applying for the {} position that is listed on {}.
                I've attached my resume, cover letter, and reference. If there's any additional information you need, please let me know.

                Thank you very much for your consideration.
                Sincerely,
                Max Duong
                <hr>
                Data Analyst
                Email: maxduong97@gmail.com  
                Phone: +1 604-318-4092
                Website: maxduong.github.io/Max-Duong
                Linkedln: linkedin.com/in/max-duong
            </p>    
            """.format(company_name, job_title, website)
    
    receiver = email
    subject = job_title+ ' Position'
    contents = body
    filename = ['../Resume-Max.pdf', '../CoverLetter-Max.pdf', '../Reference-Max.pdf']

    yag.send(
        to=receiver,
        subject=subject,
        contents=contents, 
        attachments=filename,
    )

def run_over_df_to_send_email(row):
    print(row)
    job_title = row['Job Applied For'].title()
    company_name = row['Company Name'].title()
    email = row['E-mail Address']
    website = row['Web Site']
    website = website_clean(website)
    

    # Change company's name
    replace_company_name(company_name)
    # Convert document to pdf file
    doc_to_pdf_converter()
    # Change the company's name to the original name
    replace_company_name_original(company_name)

    # Sending resume via email
    send_resume_to_recruiter(job_title, company_name, email, website)


def main():

    df = pd.read_excel('./contact_list.xlsx')
    df = df[:1]

    df.apply(run_over_df_to_send_email, axis =1)

if __name__ == "__main__":
    main()