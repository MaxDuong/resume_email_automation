import numpy as np
import pandas as pd
import json
import yagmail
from word_to_pdf_converter import job_title_cleaning, replace_company_name, replace_company_name_original, doc_to_pdf_converter

def email_authentication():
    # Opening JSON file
    f = open('authen.json',)

    # returns JSON object as a dictionary
    authen = json.load(f)
    return authen

def website_clean(website_url):
    website_array = ['indeed', 'glassdoor', 'workbc', 'linkedin']
    for website in website_array:
        if website in website_url:
            return website.title()

def job_title_an_a_clean(job_title):
    # a data analyst 
    if job_title[1] == ' ':
        job_title = job_title[2:]
    #an office data analyst
    else:
        job_title = job_title[3:]
    return job_title

def send_resume_to_recruiter(job_title, company_name, email, website):
    
    job_title = job_title_an_a_clean(job_title)
    
    # Authentication
    authen = email_authentication()
    #yag = yagmail.SMTP({'maxduong97@gmail.com': 'Max Duong'}, 'password')
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
    subject = job_title+ ' Application'
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
    job_title = row['Job Applied For']
    job_title = job_title_cleaning(job_title)
    company_name = row['Company Name'].title()
    email = row['E-mail Address']
    website = row['Web Site']
    website = website_clean(website)
    

    # Change company's name
    replace_company_name(company_name, job_title)
    # Convert document to pdf file
    doc_to_pdf_converter()
    # Change the company's name to the original name
    replace_company_name_original(company_name, job_title)

    # Sending resume via email
    send_resume_to_recruiter(job_title, company_name, email, website)


def main():

    df = pd.read_excel('./contact_list.xlsx')

    df.apply(run_over_df_to_send_email, axis =1)

if __name__ == "__main__":
    main()