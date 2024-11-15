import pandas as pd
from textblob import TextBlob

# Load the uploaded file
file_path = ''

df = pd.read_excel(file_path, sheet_name='F23 Attendance')

# Extract relevant columns
columns_of_interest = ['Interactive Question:', 'Do you have any questions for Daniel Ly?']

# Ensure we have the columns and fill missing values with empty strings
df[columns_of_interest] = df[columns_of_interest].fillna('')

# Define a list of terms to skip in sentiment analysis
ignore_list = ['n/a', 'na', 'none', 'no', 'not sure', '-', '', 'nil']

# Function to check if a text is considered irrelevant
def is_irrelevant(text):
    return text.strip().lower() in ignore_list

# Function for sentiment analysis
def analyze_sentiment(text):
    if is_irrelevant(text):
        return 'Not Applicable'
    analysis = TextBlob(text)
    if analysis.sentiment.polarity > 0:
        return 'Positive'
    elif analysis.sentiment.polarity < 0:
        return 'Negative'
    else:
        return 'Neutral'

# Perform sentiment analysis on each column independently
df['Sentiment_F'] = df['Interactive Question:'].apply(analyze_sentiment)
df['Sentiment_G'] = df['Do you have any questions for Daniel Ly?'].apply(analyze_sentiment)

# Filter out rows where both sentiments are 'Not Applicable'
df = df[~((df['Sentiment_F'] == 'Not Applicable') & (df['Sentiment_G'] == 'Not Applicable'))]

# Save the combined cleaned data to a new Excel file
df.to_excel('', index=False)