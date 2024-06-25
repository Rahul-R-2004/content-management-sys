# Install the required libraries
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import random
import pandas as pd
from google.colab import files

# Step 1: Upload the Excel file
uploaded = files.upload()

# Load the Excel file into a DataFrame
excel_file = list(uploaded.keys())[0]
df = pd.read_excel(excel_file)

# Extract the list of keywords
keywords = df.iloc[:, 0].tolist()  # Assuming keywords are in the first column

# Step 2: Upload the credentials file
uploaded = files.upload()

# Step 3: Authenticate and connect to Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("/content/" + list(uploaded.keys())[0], scope)
client = gspread.authorize(creds)
sheet = client.open("ContentCreator")
worksheet = sheet.get_worksheet(0)

# Step 4: Generate captions and hashtags function
def generate_captions_and_hashtags(keyword):
    hooklines = [
        "Discover the secrets of", "Unlock the potential of", "Master the art of",
        "Explore the world of", "Transform your life with"
    ]
    descriptions = [
        "Learn how to use", "Tips and tricks for", "The ultimate guide to",
        "Everything you need to know about", "Get started with"
    ]
    hashtags = [
        "#lifehack", "#tips", "#guide", "#tutorial", "#howto"
    ]

    captions = []
    for _ in range(15):
        hookline = random.choice(hooklines)
        description = random.choice(descriptions)
        selected_hashtags = random.sample(hashtags, 3)
        caption = f"{hookline} {keyword}: {description} {keyword} with {' '.join(selected_hashtags)}"
        captions.append(caption)

    return captions

# Step 5: Generate captions and hashtags for all keywords and prepare batch update data
batch_updates = []
for i, keyword in enumerate(keywords):
    captions = generate_captions_and_hashtags(keyword)
    for j, caption in enumerate(captions):
        cell_range = gspread.utils.rowcol_to_a1(i+2, j+2)  # i+2 for row (skip header and start at row 2), j+2 for columns B to P
        batch_updates.append({
            'range': cell_range,
            'values': [[caption]]
        })

# Step 6: Perform batch update
worksheet.batch_update(batch_updates)

print("Captions generated and written to Google Sheet successfully.")
