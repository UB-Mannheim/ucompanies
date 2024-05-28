import requests
import os
import pandas as pd
from urllib.parse import urlsplit


a = pd.read_stata('../data/Beneficiaries2022_FDZ.dta')

def download_pdfs(url_list, download_folder):
    # Ensure the download folder exists
    if not os.path.exists(download_folder):
        os.makedirs(download_folder)

    for url in url_list:
        try:
            # Send a GET request to the URL
            response = requests.get(url, stream=True)
            response.raise_for_status()  # Check for HTTP errors

            # Extract the original filename from the URL
            filename = os.path.basename(urlsplit(url).path)
            if not filename:
                print(f"Failed to extract filename from URL: {url}")
                continue

            # Determine the full path to save the file
            file_path = os.path.join(download_folder, filename + '.pdf')

            # Write the content to a file
            with open(file_path, 'wb') as pdf_file:
                for chunk in response.iter_content(chunk_size=8192):
                    pdf_file.write(chunk)

            print(f"Downloaded: {file_path}")

        except requests.exceptions.RequestException as e:
            print(f"Failed to download {url}: {e}")

if __name__ == "__main__":
    # List of URLs to download
    url_list = a.applform.to_list()

    # Folder to save the downloaded PDFs
    download_folder = "../data/raw_pdfs/"

    download_pdfs(url_list, download_folder)
