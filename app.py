# -*- coding: utf-8 -*-

import nltk
nltk.download('wordnet')
nltk.download('punkt')
import requests
from bs4 import BeautifulSoup
import pandas as pd
import nltk
from nltk.tokenize import word_tokenize
from nltk.stem import WordNetLemmatizer
import re
from collections import Counter
from textblob import TextBlob

# Load stop words
stop_words_files = ['StopWords/StopWords_Currencies.txt',
                    'StopWords/StopWords_DatesandNumbers.txt',
                    'StopWords/StopWords_Names.txt',
                    'StopWords/StopWords_Generic.txt',
                    'StopWords/StopWords_GenericLong.txt',
                    'StopWords/StopWords_Auditor.txt',
                    'StopWords/StopWords_Geographic.txt']

stop_words = set()
for file in stop_words_files:
    with open(file, 'r', encoding='latin-1') as f:
        stop_words.update(f.read().splitlines())


# Load positive and negative words
positive_words = set(open('MasterDictionary/positive-words.txt', 'r').read().splitlines())
negative_words = set(open('MasterDictionary/negative-words.txt', 'r', encoding='latin-1').read().splitlines())


# Function to clean text
def clean_text(text):
    # Tokenize
    tokens = word_tokenize(text.lower())
    # Remove stopwords and punctuations
    cleaned_tokens = [WordNetLemmatizer().lemmatize(word) for word in tokens if word.isalnum() and word not in stop_words]
    return cleaned_tokens

# Function to count syllables
def count_syllables(word):
    word = word.lower()
    count = 0
    vowels = "aeiouy"
    if word[0] in vowels:
        count += 1
    for index in range(1, len(word)):
        if word[index] in vowels and word[index - 1] not in vowels:
            count += 1
    if word.endswith("e"):
        count -= 1
    if word.endswith("le") and len(word) > 2 and word[-3] not in vowels:
        count += 1
    if count == 0:
        count += 1
    return count

# Function to calculate readability metrics
def calculate_readability(text):
    sentences = nltk.sent_tokenize(text)
    words = clean_text(text)
    num_words = len(words)
    num_sentences = len(sentences)
    avg_sentence_length = num_words / num_sentences
    complex_words = [word for word in words if count_syllables(word) > 2]
    num_complex_words = len(complex_words)
    percentage_complex_words = (num_complex_words / num_words) * 100
    fog_index = 0.4 * (avg_sentence_length + percentage_complex_words)
    avg_words_per_sentence = num_words / num_sentences
    return avg_sentence_length, percentage_complex_words, fog_index, avg_words_per_sentence

# Function to calculate sentiment scores
def calculate_sentiment(text):
    cleaned_text = ' '.join(clean_text(text))
    blob = TextBlob(cleaned_text)
    polarity_score = blob.sentiment.polarity
    subjectivity_score = blob.sentiment.subjectivity
    positive_score = sum(1 for word in blob.words if word in positive_words)
    negative_score = sum(1 for word in blob.words if word in negative_words)
    return positive_score, negative_score, polarity_score, subjectivity_score

# Function to count personal pronouns
def count_personal_pronouns(text):
    personal_pronouns = re.findall(r'\b(?:I|we|my|ours|us)\b', text, flags=re.IGNORECASE)
    return len(personal_pronouns)

# Function to calculate average word length
def calculate_avg_word_length(text):
    words = clean_text(text)
    total_chars = sum(len(word) for word in words)
    num_words = len(words)
    return total_chars / num_words if num_words > 0 else 0

# Function to process each article
def process_article(url_id, article_text):
    avg_sentence_length, percentage_complex_words, fog_index, avg_words_per_sentence = calculate_readability(article_text)
    positive_score, negative_score, polarity_score, subjectivity_score = calculate_sentiment(article_text)
    complex_word_count = sum(1 for word in clean_text(article_text) if count_syllables(word) > 2)
    word_count = len(clean_text(article_text))
    syllable_per_word = sum(count_syllables(word) for word in clean_text(article_text)) / word_count
    personal_pronouns = count_personal_pronouns(article_text)
    avg_word_length = calculate_avg_word_length(article_text)

    return [url_id, avg_sentence_length, percentage_complex_words, fog_index, avg_words_per_sentence,
            positive_score, negative_score, polarity_score, subjectivity_score,
            complex_word_count, word_count, syllable_per_word, personal_pronouns, avg_word_length]


# Function to extract article text from URL
def extract_article_text(url):
    try:
        response = requests.get(url)
        response.raise_for_status()  # Raise an exception for 4XX and 5XX status codes
        if response.status_code == 200:
            html = response.text
            soup = BeautifulSoup(html, 'html.parser')
            article_content = soup.find('div', class_='td-post-content tagdiv-type')  # Adjust class name as per HTML structure
            if article_content:
                paragraphs = article_content.find_all('p')
                article_text = ' '.join([p.get_text(strip=True) for p in paragraphs])
                return article_text
    except Exception as e:
        print(f"Error fetching article from {url}: {e}")
    return None


# Read input Excel file
input_data = pd.read_excel('Input.xlsx')

# Process articles and generate output
output_data = []
for index, row in input_data.iterrows():
    url_id = row['URL_ID']
    url = row['URL']
    article_text = extract_article_text(url)
    if article_text:
        output_data.append(process_article(url_id, article_text))

# Writing output to Excel
output_df = pd.DataFrame(output_data, columns=['URL_ID', 'Avg_Sentence_Length', 'Percentage_of_Complex_Words', 'Fog_Index',
                                               'Avg_Number_of_Words_Per_Sentence', 'Positive_Score', 'Negative_Score',
                                               'Polarity_Score', 'Subjectivity_Score', 'Complex_Word_Count',
                                               'Word_Count', 'Syllable_Per_Word', 'Personal_Pronouns', 'Avg_Word_Length'])
output_df.to_excel('Output.xlsx', index=False)