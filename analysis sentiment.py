import pandas as pd
from textblob import TextBlob

# Load the Excel file
file_path = 'Final product reviews 200.xlsx'
df = pd.read_excel(file_path, sheet_name='Sheet1')

# Function to calculate sentiment polarity for a given text series
def calculate_sentiment(text_series):
    # Combine the text series into a single corpus
    corpus = text_series.dropna().astype(str).tolist()
    
    # Skip column if all values are the same
    if len(set(corpus)) <= 1:
        return None
    
    # Calculate sentiment polarity for each text entry
    sentiment_scores = [TextBlob(text).sentiment.polarity for text in corpus]
    return sentiment_scores

# Calculate and display average sentiment polarity for each column with a header
for column in df.columns:
    if df[column].dtype == object:  # Ensures the column contains text
        print(f"Sentiment analysis for column '{column}':")
        sentiment_scores = calculate_sentiment(df[column])
        if sentiment_scores:
            average_sentiment = sum(sentiment_scores) / len(sentiment_scores)
            print(f"Average Sentiment Polarity: {average_sentiment}")
        else:
            print("Not enough data for sentiment analysis or column contains common values.")
        print("\n")
