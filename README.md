
# Custom PDF Processing and Document Generation Tool

## Description
This tool is designed to process PDF files, search for specific text, and generate a Word document based on the search results. For example, input Queensland Health and Safety Legislation, ask for a summary of Enterprise Bargining Agreements, and have a structured word document generated with headings organised based on the results of the inputted PDF documents. This function has largely been replaced by GPTs and other RAG systems, but I think this is still a cool starting off point if you want to generate documents with information from a wide range of sources. It utilizes several Python libraries for PDF processing, natural language processing, and document creation.

## Features
- PDF text extraction with `pdfplumber`
- Text processing and search with `nltk` and custom search algorithms
- Document generation with `python-docx`
- Integration with OpenAI's GPT-3 for advanced text analysis and generation
- GUI interface with `tkinter`

## Installation
To install the required libraries, run:
```bash
pip install pdfplumber pandas numpy pickle nltk requests tkinter python-docx openai
```

## Usage
Run the script with:
```bash
python script_name.py
```
Follow the GUI prompts to add PDF files, specify output folders, and input search phrases.

## Note
Make sure to replace `"YOUR API KEY"` with your actual OpenAI API key.
