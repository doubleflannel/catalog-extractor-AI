import streamlit as st
import pandas as pd
from scraper_HCOE_v2 import run_scraper, save_results_to_excel
from datetime import datetime
import os
import re

# Title of the web app
st.title('Product Catalog Data Extractor')
st.write('Proof of Concept of an AI tool to extract specs from an online product catalog. Uses GPT-4o. Developed in 2 weeks by Ivan for Thomas & Jack (HCOE).')

# Input for ChatGPT API key
api_key = st.text_input("Enter your ChatGPT API key", type="password")

# File uploader allows user to add their own Excel
uploaded_file = st.file_uploader("Choose an Excel file that contains the URLs of the product catalog", type=['xlsx'])

if uploaded_file is not None:
    # To read file as bytes:
    bytes_data = uploaded_file.getvalue()
    # To convert to a DataFrame
    df = pd.read_excel(uploaded_file)
    st.write(df)  # Displaying the dataframe on the page

    # Dropdown to select the column containing URLs
    url_column = st.selectbox("Select the column containing URLs", df.columns)

    # Text area for custom prompt
    default_prompt = ("Make a table that specifies the main product on this page, not the suggested products. "
                      "The table should contain the Product name, McKesson #, Manufacturer #, Brand #, Manufacturer, "
                      "Country of Origin, Active Ingredients, Application, Container Type, Dimensions, Form, Sterility, "
                      "Type, UNSPSC Code, Volume. If Active Ingredients are available in the features tab, use the "
                      "definition to replace the Active Ingredients field from the Product Specifications")
    custom_prompt = st.text_area("Enter your custom prompt", default_prompt)

    # Button to start scraping
    if st.button('Scrape Data'):
        # Extract URLs from the selected column
        urls = df[url_column].dropna().astype(str).tolist()

        # Validate URLs
        valid_urls = []
        invalid_urls = []
        url_pattern = re.compile(r'^(http|https)://')

        for url in urls:
            if url_pattern.match(url):
                valid_urls.append(url)
            else:
                invalid_urls.append(url)

        # Inform the user about invalid URLs
        if invalid_urls:
            st.warning(f"The following URLs are invalid and will be skipped: {invalid_urls}")

        # Initialize progress bar
        progress_bar = st.progress(0)
        status_text = st.empty()

        # Progress callback function
        def progress_callback(current, total):
            progress_bar.progress(current / total)
            status_text.text(f"Processing {current}/{total} URLs")

        # Run scraper with progress callback
        results = run_scraper(valid_urls, api_key=api_key, custom_prompt=custom_prompt, progress_callback=progress_callback)
        
        # Display results as a table
        results_df = pd.DataFrame(results)
        st.write(results_df)
        
        # Save results to Excel
        output_filename = f'ProductSpec_{datetime.now().strftime("%Y%m%d%H%M")}.xlsx'
        save_results_to_excel(results, output_filename)
        
        # Provide download link for the output file
        with open(output_filename, "rb") as file:
            btn = st.download_button(
                label="Download Excel",
                data=file,
                file_name=output_filename,
                mime="application/vnd.ms-excel"
            )

# Running the app: streamlit run app_scraper_HCOE_v2.py

