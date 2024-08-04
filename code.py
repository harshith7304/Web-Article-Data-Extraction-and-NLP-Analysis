#!/usr/bin/env python
# coding: utf-8

# In[32]:


get_ipython().system('pip install textblob')
get_ipython().system('pip install nltk')
get_ipython().system('pip install chardet')

import pandas as pd
import requests
from bs4 import BeautifulSoup
import os
import chardet
from textblob import TextBlob
import nltk
import re

nltk.download('punkt')
# Read the input Excel file
input_path = 'C:/Users/Dell/Downloads/Input.xlsx'

input_df = pd.read_excel(input_path)


# In[33]:


input_path = 'C:/Users/Dell/Downloads/Input.xlsx'

input_df = pd.read_excel(input_path)


# In[34]:


def extract_article_text(url):
    response = requests.get(url)
    soup = BeautifulSoup(response.content, 'html.parser')
    
    # Extract title and article text
    title = soup.find('h1').get_text() if soup.find('h1') else 'No Title'
    paragraphs = soup.find_all('p')
    article_text = ' '.join([para.get_text() for para in paragraphs])
    
    return title, article_text

# Create a directory to save extracted articles
os.makedirs('articles', exist_ok=True)

# Extract data for each URL
for index, row in input_df.iterrows():
    url = row['URL']
    url_id = row['URL_ID']
    
    try:
        title, article_text = extract_article_text(url)
        
        # Save to text file
        with open(f'articles/{url_id}.txt', 'w', encoding='utf-8') as file:
            file.write(title + '\n' + article_text)
        print(f'Successfully extracted article {url_id}')
    except Exception as e:
        print(f'Failed to extract article {url_id}: {e}')

print("Data extraction complete.")


# In[40]:


# Function to detect encoding and read file
def read_file_with_encoding(filepath):
    with open(filepath, 'rb') as file:
        raw_data = file.read()
        result = chardet.detect(raw_data)
        encoding = result['encoding']
    with open(filepath, 'r', encoding=encoding, errors='ignore') as file:
        return file.read().splitlines()

# Load stop words
stopwords_dir = 'C:/Users/Dell/Downloads/StopWords-20240530T125937Z-001'
stopwords_files = [
    'StopWords_Auditor.txt', 'StopWords_Currencies.txt', 'StopWords_DatesandNumbers.txt',
    'StopWords_Generic.txt', 'StopWords_GenericLong.txt', 'StopWords_Geographic.txt',
    'StopWords_Names.txt'
]
stopwords = set()
for filename in stopwords_files:
    filepath = os.path.join(stopwords_dir, filename)
    if os.path.isfile(filepath):
        stopwords.update(read_file_with_encoding(filepath))
    else:
        print(f"File {filename} not found in directory {stopwords_dir}")

# Load positive and negative words
positive_words = set()
positive_words_path = 'C:/Users/Dell/Downloads/MasterDictionary-20240530T125936Z-001/MasterDictionary/positive-words.txt'
if os.path.isfile(positive_words_path):
    positive_words.update(read_file_with_encoding(positive_words_path))
else:
    print(f"File positive-words not found in directory")

negative_words = set()
negative_words_path = 'C:/Users/Dell/Downloads/MasterDictionary-20240530T125936Z-001/MasterDictionary/negative-words.txt'
if os.path.isfile(negative_words_path):
    negative_words.update(read_file_with_encoding(negative_words_path))
else:
    print(f"File negative-words not found in directory")

# Function to compute text analysis variables
def analyze_text(text):
    analysis = {}
    words = nltk.word_tokenize(text.lower())
    words = [word for word in words if word.isalpha() and word not in stopwords]
    
    # Positive and negative word count
    positive_count = sum(1 for word in words if word in positive_words)
    negative_count = sum(1 for word in words if word in negative_words)
    
    analysis['POSITIVE SCORE'] = positive_count
    analysis['NEGATIVE SCORE'] = negative_count
    
    # Polarity and subjectivity
    blob = TextBlob(' '.join(words))
    analysis['POLARITY SCORE'] = blob.sentiment.polarity
    analysis['SUBJECTIVITY SCORE'] = blob.sentiment.subjectivity
    
    # Average sentence length
    sentences = nltk.sent_tokenize(text)
    analysis['AVG SENTENCE LENGTH'] = sum(len(sentence.split()) for sentence in sentences) / len(sentences)
    
    # Percentage of complex words
    complex_words = [word for word in words if len(word) > 2 and (len(re.findall('[aeiouy]+', word)) > 2)]
    analysis['PERCENTAGE OF COMPLEX WORDS'] = len(complex_words) / len(words) if words else 0
    
    # Fog Index
    analysis['FOG INDEX'] = 0.4 * (analysis['AVG SENTENCE LENGTH'] + analysis['PERCENTAGE OF COMPLEX WORDS'] * 100)
    
    # Average number of words per sentence
    analysis['AVG NUMBER OF WORDS PER SENTENCE'] = analysis['AVG SENTENCE LENGTH']
    
    # Complex word count
    analysis['COMPLEX WORD COUNT'] = len(complex_words)
    
    # Word count
    analysis['WORD COUNT'] = len(words)
    
    # Syllable per word
    syllable_per_word = sum(len(re.findall('[aeiouy]+', word)) for word in words) / len(words) if words else 0
    analysis['SYLLABLE PER WORD'] = syllable_per_word
    
    # Personal pronouns
    personal_pronouns = re.findall(r'\b(I|we|my|ours|us)\b', text, re.I)
    analysis['PERSONAL PRONOUNS'] = len(personal_pronouns)
    
    # Average word length
    avg_word_length = sum(len(word) for word in words) / len(words) if words else 0
    analysis['AVG WORD LENGTH'] = avg_word_length
    
    return analysis

# Read input and initialize output dataframe
input_path = 'C:/Users/Dell/Downloads/Input.xlsx'
input_df = pd.read_excel(input_path)

# Perform analysis on each extracted text
results = []
for index, row in input_df.iterrows():
    url = row['URL']
    url_id = row['URL_ID']
    
    try:
        title, text = extract_article_text(url)
        
        # Analyze text and collect results
        analysis_results = analyze_text(text)
        analysis_results['URL_ID'] = url_id
        results.append(analysis_results)
            
        print(f'Successfully analyzed article {url_id}')
    except Exception as e:
        print(f'Failed to analyze article {url_id}: {e}')

# Convert results to DataFrame
output_df = pd.DataFrame(results)

# Merge with the input data to keep the original columns
final_df = pd.merge(input_df, output_df, on='URL_ID', how='left')

# Save output dataframe to Excel
final_df.to_excel('C:/Users/Dell/Downloads/Output Data Structure.xlsx', index=False)
print("Text analysis complete.")


# In[ ]:





# In[ ]:





# In[ ]:




