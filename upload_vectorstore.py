import os
import openai
from dotenv import load_dotenv
import csv
from pathlib import Path

# Load environment variables
load_dotenv()
openai.api_key = os.getenv("OPENAI_API_KEY")
vector_store_id = os.getenv("VECTOR_STORE_ID")

# Directory containing the files
output_dir = 'nedlastede_filer'

# Read the CSV file to get filenames and URLs
filedatabase = 'downloaded_files.csv'
downloaded_file_urls = {}
if Path(filedatabase).exists():
    print("Reading the existing database of downloaded files:")
    with open(filedatabase, 'r', newline='', encoding='utf-8') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            filename = row['Filename']
            url = row['URL']
            downloaded_file_urls[filename] = url
    print(f"Read {len(downloaded_file_urls)} files from database.")
else:
    print(f"File {filedatabase} does not exist.")
    exit(1)

# Process each file
for filename, url in downloaded_file_urls.items():
    filepath = os.path.join(output_dir, os.path.basename(filename))
    if not os.path.exists(filepath):
        print(f"File {filepath} does not exist, skipping.")
        continue
    print(f"Uploading file: {filepath}")

    try:
        # Upload the file to OpenAI
        with open(filepath, "rb") as f:
            response = openai.File.create(
                file=f,
                purpose="vectors"
            )
        file_id = response["id"]
        print(f"Uploaded file {filepath} with file ID {file_id}")

        # Attach the file to the vector store with metadata
        vector_response = openai.VectorStoreFile.create(
            vectorstore_id=vector_store_id,
            file_id=file_id,
            metadata={
                "source_url": url
            }
        )
        print(f"Attached file {file_id} to vector store {vector_store_id}")

    except Exception as e:
        print(f"Error uploading file {filepath}: {e}")
