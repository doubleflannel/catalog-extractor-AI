import xlsxwriter
import json
import os
import openpyxl
import pandas as pd
import requests
from scrapegraphai.graphs import SmartScraperGraph
from scrapegraphai.utils import prettify_exec_info
from datetime import datetime

# Default configuration for SmartScraperGraph
def get_default_graph_config(api_key):
    return {
        "llm": {
            "api_key": api_key,
            "model": "gpt-4-turbo",
        },
        "verbose": True,
    }

# Function to read URLs from an Excel file
def read_urls_from_excel(file_path, sheet_name='Sheet1', column_name='URLs'):
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    return df[column_name].tolist()

# Function to fetch the content of a URL
def fetch_url_content(url):
    response = requests.get(url)
    response.raise_for_status()
    return response.text

# Function to initialize and run the scraper for each URL
def run_scraper(urls, api_key, custom_prompt=None, config=None, progress_callback=None):
    all_results = []
    if config is None:
        config = get_default_graph_config(api_key)
    total_urls = len(urls)
    for i, source in enumerate(urls):
        prompt = custom_prompt if custom_prompt else "Make a table that specifies the main product on this page, not the suggested products. The table should contain the Product name, McKesson #, Manufacturer #, Brand #, Manufacturer, Country of Origin, Active Ingredients, Application, Container Type, Dimensions, Form, Sterility, Type, UNSPSC Code, Volume. If Active Ingredients are available in the features tab, use the definition to replace the Active Ingredients field from the Product Specifications"
        
        # Fetch the content of the URL
        content = fetch_url_content(source)
        
        smart_scraper_graph = SmartScraperGraph(
            prompt=prompt,
            source=content,
            config=config
        )
        result = smart_scraper_graph.run()
        result['Product URL'] = source
        all_results.append(result)
        graph_exec_info = smart_scraper_graph.get_execution_info()
        print(prettify_exec_info(graph_exec_info))
        
        # Update progress
        if progress_callback:
            progress_callback(i + 1, total_urls)
            
    return all_results

# Function to save results to an Excel file
def save_results_to_excel(results, output_filename):
    if not results:
        print("No results to save.")
        return

    workbook = xlsxwriter.Workbook(output_filename)
    worksheet = workbook.add_worksheet()

    # Extract headers dynamically from the first result
    headers = list(results[0].keys())

    # Write headers to the first row
    for col, header in enumerate(headers):
        worksheet.write(0, col, header)

    # Write data rows
    for row, result_set in enumerate(results, start=1):
        for col, key in enumerate(headers):
            worksheet.write(row, col, result_set.get(key, ''))

    workbook.close()

# Main function to orchestrate the scraping and saving process
def main(excel_file_path, output_file_path, api_key, custom_prompt=None):
    urls = read_urls_from_excel(excel_file_path)
    results = run_scraper(urls, api_key, custom_prompt=custom_prompt)
    save_results_to_excel(results, output_file_path)

# Execution
if __name__ == "__main__":
    input_excel_file = 'C:/Users/vnkbr/ai-scrape-HCOE/souce-test1.xlsx'  # Replace with your input Excel file path
    output_excel_file = f'ProductSpec_{datetime.now().strftime("%Y%m%d%H%M")}.xlsx'
    main(input_excel_file, output_excel_file)
