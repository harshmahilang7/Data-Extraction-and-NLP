# -*- coding: utf-8 -*-
# @Author: Dastan Alam
# @Date:   2024-07-01 05:26:41 PM   17:07
# @Last Modified by:   Dastan Alam
# @Last Modified time: 2024-07-01 05:31:06 PM   17:07

# @Email:  HARSHMAHILANG7@GMAIL.COM
# @workspaceFolder:  c:\BACKCOFFER
# @workspaceFolderBasename:  BACKCOFFER
# @file:  c:\BACKCOFFER\20211030 Test Assignment\Data_Analysis.py
# @relativeFile:  20211030 Test Assignment\Data_Analysis.py
# @relativeFileDirname:  20211030 Test Assignment
# @fileBasename:  Data_Analysis.py
# @fileBasenameNoExtension:  Data_Analysis
# @fileDirname:  c:\BACKCOFFER\20211030 Test Assignment
# @fileExtname:  .py
# @cwd:  c:\BACKCOFFER


import requests
from bs4 import BeautifulSoup
import pandas as pd
import os
from textblob import TextBlob
import re
import nltk
from nltk.corpus import stopwords, cmudict
from openpyxl import load_workbook

# Ensure the necessary NLTK data is downloaded
nltk.download('punkt')
nltk.download('averaged_perceptron_tagger')
nltk.download('stopwords')
nltk.download('cmudict')  # Download cmudict

# Read input URLs from Input.xlsx
input_file = '20211030 Test Assignment\\Input.xlsx'
output_file = 'Output Data Structure.xlsx'
input_df = pd.read_excel(input_file)

# Define stopwords and cmudict for syllable counting
stop_words = set(stopwords.words('english'))
d = cmudict.dict()

# Function to count syllables
def syllable_count(word):
    if word.lower() in d:
        return max([len(list(y for y in x if y[-1].isdigit())) for x in d[word.lower()]])
    else:
        return len(re.findall(r'[aeiouy]+', word.lower()))

# Read positive and negative words from files
with open('20211030 Test Assignment\\MasterDictionary\\positive-words.txt', 'r') as file:
    positive_words = set(file.read().split())

with open('20211030 Test Assignment\\MasterDictionary\\negative-words.txt', 'r') as file:
    negative_words = set(file.read().split())

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

# Function to perform text analysis
def analyze_text(text):
    # Clean the text
    words = nltk.word_tokenize(text)
    cleaned_words = [word for word in words if word.lower() not in stop_words and word.isalpha()]
    
    # Positive, Negative, Polarity, Subjectivity Scores
    positive_score = sum(1 for word in cleaned_words if word.lower() in positive_words)
    negative_score = sum(1 for word in cleaned_words if word.lower() in negative_words)
    polarity_score = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)
    subjectivity_score = (positive_score + negative_score) / (len(cleaned_words) + 0.000001)
    
    # Sentence Analysis
    sentences = nltk.sent_tokenize(text)
    num_sentences = len(sentences)
    num_words = len(cleaned_words)
    avg_sentence_length = num_words / num_sentences if num_sentences > 0 else 0
    
    # Complex words (words with 3+ syllables)
    complex_words = [word for word in cleaned_words if syllable_count(word) >= 3]
    complex_word_count = len(complex_words)
    percentage_complex_words = (complex_word_count / num_words) * 100 if num_words > 0 else 0
    
    # Fog Index
    fog_index = 0.4 * (avg_sentence_length + percentage_complex_words)
    
    # Average number of words per sentence
    avg_words_per_sentence = avg_sentence_length
    
    # Word Count
    word_count = num_words
    
    # Syllables per word
    syllables_per_word = sum(syllable_count(word) for word in cleaned_words) / num_words if num_words > 0 else 0
    
    # Personal Pronouns
    personal_pronouns = len([word for word in cleaned_words if word.lower() in ['i', 'we', 'my', 'ours', 'us']])
    
    # Average word length
    avg_word_length = sum(len(word) for word in cleaned_words) / num_words if num_words > 0 else 0
    
    return [
        positive_score, negative_score, polarity_score, subjectivity_score, avg_sentence_length,
        percentage_complex_words, fog_index, avg_words_per_sentence, complex_word_count, word_count,
        syllables_per_word, personal_pronouns, avg_word_length
    ]

# Ensure output directory exists
output_dir = 'extracted_articles'
os.makedirs(output_dir, exist_ok=True)

# Prepare for output data
output_data = []

# Process each URL
for index, row in input_df.iterrows():
    url_id = row['URL_ID']
    url = row['URL']
    
    title, text = extract_article_text(url)
    
    if title and text:
        # Save extracted text to a file with URL_ID as its file name
        with open(os.path.join(output_dir, f'{url_id}.txt'), 'w', encoding='utf-8') as file:
            file.write(f"{title}\n\n{text}")
        
        # Perform text analysis
        analysis_results = analyze_text(text)
        
        # Prepare output row
        output_row = [url_id, url] + analysis_results
        output_data.append(output_row)

# Define output DataFrame
output_columns = [
    'URL_ID', 'URL', 'POSITIVE SCORE', 'NEGATIVE SCORE', 'POLARITY SCORE', 'SUBJECTIVITY SCORE',
    'AVG SENTENCE LENGTH', 'PERCENTAGE OF COMPLEX WORDS', 'FOG INDEX', 'AVG NUMBER OF WORDS PER SENTENCE',
    'COMPLEX WORD COUNT', 'WORD COUNT', 'SYLLABLE PER WORD', 'PERSONAL PRONOUNS', 'AVG WORD LENGTH'
]
output_df = pd.DataFrame(output_data, columns=output_columns)

# Save output DataFrame to Excel
with pd.ExcelWriter(output_file, engine='openpyxl', mode='a') as writer:
    book = load_workbook(output_file)
    writer.book = book
    output_df.to_excel(writer, index=False, sheet_name='Output')

print("Article extraction and analysis completed.")