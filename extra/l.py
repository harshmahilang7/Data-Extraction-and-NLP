# -*- coding: utf-8 -*-
# @Author: Dastan Alam
# @Date:   2024-07-01 07:28:19 PM   19:07
# @Last Modified by:   Dastan Alam
# @Last Modified time: 2024-07-02 01:05:27 PM   13:07

# @Email:  HARSHMAHILANG7@GMAIL.COM
# @workspaceFolder:  c:\BACKCOFFER\20211030 Test Assignment
# @workspaceFolderBasename:  20211030 Test Assignment
# @file:  c:\BACKCOFFER\20211030 Test Assignment\extracted_articles\M.PY
# @relativeFile:  extracted_articles\M.PY
# @relativeFileDirname:  extracted_articles
# @fileBasename:  M.PY
# @fileBasenameNoExtension:  M
# @fileDirname:  c:\BACKCOFFER\20211030 Test Assignment\extracted_articles
# @fileExtname:  .PY
# @cwd:  c:\BACKCOFFER\20211030 Test Assignment

import requests
from bs4 import BeautifulSoup
import pandas as pd
import os
import re
import nltk
from nltk.corpus import stopwords, cmudict
from tqdm import tqdm
from concurrent.futures import ThreadPoolExecutor

# Ensure the necessary NLTK data is downloaded
nltk.download('punkt')
nltk.download('averaged_perceptron_tagger')
nltk.download('stopwords')
nltk.download('cmudict')  # Download cmudict

# Function to count syllables
def syllable_count(word):
    d = cmudict.dict()
    if word.lower() in d:
        return max([len(list(y for y in x if y[-1].isdigit())) for x in d[word.lower()]])
    else:
        return len(re.findall(r'[aeiouy]+', word.lower()))

# Function to perform text analysis
def analyze_text(text, stop_words, positive_words, negative_words):
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

# Function to extract article title and text
def extract_article_text(url, session):
    try:
        response = session.get(url)
        soup = BeautifulSoup(response.content, 'html.parser')
        
        # Example logic to extract title and text; adjust based on actual HTML structure
        title = soup.find('h1').get_text().strip() if soup.find('h1') else ''
        paragraphs = soup.find_all('p')
        text = '\n'.join([para.get_text().strip() for para in paragraphs])
        
        return title, text
    except Exception as e:
        print(f"Error extracting {url}: {e}")
        return '', ''

# Main function to process URLs
def process_urls(input_file, output_file):
    # Read input URLs from Excel
    input_df = pd.read_excel(input_file)
    
    # Load additional stopwords from provided files
    stopwords_files = [
        'StopWords\\StopWords_Currencies.txt',
        'StopWords\\StopWords_DatesandNumbers.txt',
        'StopWords\\StopWords_Generic.txt',
        'StopWords\\StopWords_GenericLong.txt',
        'StopWords\\StopWords_Geographic.txt',
        'StopWords\\StopWords_Names.txt'
    ]
    
    additional_stopwords = set()
    for file in stopwords_files:
        with open(file, 'r') as f:
            additional_stopwords.update(f.read().split())
    
    # Define stopwords and cmudict for syllable counting
    stop_words = set(stopwords.words('english')).union(additional_stopwords)
    
    # Read positive and negative words from files
    with open('MasterDictionary\\positive-words.txt', 'r') as file:
        positive_words = set(file.read().split())
    
    with open('MasterDictionary\\negative-words.txt', 'r') as file:
        negative_words = set(file.read().split())
    
    # Ensure output directory exists
    output_dir = 'extracted_articles'
    os.makedirs(output_dir, exist_ok=True)
    
    # Prepare for output data
    output_data = []
    
    # Create a requests session
    session = requests.Session()
    
    # Function to process a single URL
    def process_single_url(row):
        url_id = row['URL_ID']
        url = row['URL']
        
        title, text = extract_article_text(url, session)
        
        if title and text:
            # Save extracted text to a file with URL_ID as filename
            output_path = os.path.join(output_dir, f'{url_id}.txt')
            with open(output_path, 'w', encoding='utf-8') as file:
                file.write(f'{title}\n{text}')
            
            # Analyze the text
            analysis_results = analyze_text(text, stop_words, positive_words, negative_words)
            
            # Append results to output data
            return [url_id] + analysis_results
        return None

    # Process URLs in parallel using ThreadPoolExecutor
    with ThreadPoolExecutor(max_workers=10) as executor:
        results = list(tqdm(executor.map(process_single_url, input_df.itertuples()), total=len(input_df), desc="Processing URLs"))

    # Filter out None results
    output_data = [result for result in results if result is not None]
    
    # Define columns for output DataFrame
    columns = [
        'URL_ID', 'Positive Score', 'Negative Score', 'Polarity Score', 'Subjectivity Score',
        'Avg Sentence Length', 'Percentage Complex Words', 'Fog Index', 'Avg Words Per Sentence',
        'Complex Word Count', 'Word Count', 'Syllables Per Word', 'Personal Pronouns', 'Avg Word Length'
    ]
    
    # Create output DataFrame
    output_df = pd.DataFrame(output_data, columns=columns)
    
    # Save output DataFrame to Excel
    output_df.to_excel(output_file, index=False)

# Execute the script with input and output file paths
process_urls('Input.xlsx', 'Output Data Structure.xlsx')
