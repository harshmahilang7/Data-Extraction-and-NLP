# -*- coding: utf-8 -*-
# @Author: Dastan Alam
# @Date:   2024-07-01 05:03:17 PM   17:07
# @Last Modified by:   Dastan Alam
# @Last Modified time: 2024-07-01 05:12:38 PM   17:07

# @Email:  HARSHMAHILANG7@GMAIL.COM
# @workspaceFolder:  c:\BACKCOFFER
# @workspaceFolderBasename:  BACKCOFFER
# @file:  c:\BACKCOFFER\20211030 Test Assignment\min.py
# @relativeFile:  20211030 Test Assignment\min.py
# @relativeFileDirname:  20211030 Test Assignment
# @fileBasename:  min.py
# @fileBasenameNoExtension:  min
# @fileDirname:  c:\BACKCOFFER\20211030 Test Assignment
# @fileExtname:  .py
# @cwd:  c:\BACKCOFFER


import requests
from bs4 import BeautifulSoup
import pandas as pd
import os

# Read input URLs from Input.xlsx
input_file = '20211030 Test Assignment\\Input.xlsx'
input_df = pd.read_excel(input_file)

# Function to extract article title and text
def extract_article_text(url):
    try:
        response = requests.get(url)
        soup = BeautifulSoup(response.content, 'html.parser')
        
        # Example logic to extract title and text; adjust based on actual HTML structure
        title = soup.find('h1').get_text().strip() if soup.find('h1') else ''
        paragraphs = soup.find_all('p')
        text = '\n'.join([para.get_text().strip() for para in paragraphs])
        
        return title, text
    except Exception as e:
        print(f"Error extracting {url}: {e}")
        return '', ''

# Ensure output directory exists
output_dir = 'extracted_articles'
os.makedirs(output_dir, exist_ok=True)

# Process each URL
for index, row in input_df.iterrows():
    url_id = row['URL_ID']
    url = row['URL']
    
    title, text = extract_article_text(url)
    
    if title and text:
        # Save extracted text to a file with URL_ID as its file name
        with open(os.path.join(output_dir, f'{url_id}.txt'), 'w', encoding='utf-8') as file:
            file.write(f"{title}\n\n{text}")

print("Article extraction completed.")
