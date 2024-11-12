# -*- coding: utf-8 -*-
"""
Created on Fri Oct 11 09:20:08 2024

@author: mulwa
"""

import fitz
from PIL import Image
import pytesseract
from transformers import pipeline, AutoTokenizer
import os

out_path =r"C:\Users\mulwa\Desktop\BMC3.2\BMC 4.1\SoftWareEngineer\Chapter 8 summary 1.docx"" "
out_path2 = r"C:\Users\mulwa\Desktop\BMC3.2\BMC 4.1\SoftWareEngineer\Chapter 8 summary 2.docx"""

# Initialize summarizer and tokenizer
summarizer = pipeline("summarization", model="sshleifer/distilbart-cnn-12-6", device=-1)
tokenizer = AutoTokenizer.from_pretrained("sshleifer/distilbart-cnn-12-6") 

# Function to extract text from PDF with OCR fallback
def extract_text_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    text = ""
    for page_num in range(doc.page_count):
        page = doc.load_page(page_num)
        page_text = page.get_text()
        if not page_text.strip():
            pix = page.get_pixmap()
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            page_text = pytesseract.image_to_string(img)
        text += page_text
    doc.close()
    return text

# Function to split large text initially by paragraphs or sentences
def initial_split(text, length_threshold=1024):
    paragraphs = text.split('\n\n')
    split_text = []
    for paragraph in paragraphs:
        if len(tokenizer(paragraph)["input_ids"]) > length_threshold:
            sentences = paragraph.split('. ')
            for sentence in sentences:
                split_text.append(sentence)
        else:
            split_text.append(paragraph)
    return split_text

# Recursive function to split text down to manageable sizes
def split_text_recursively(text, max_tokens=1024):
    tokens = tokenizer(text, return_tensors="pt", truncation=False)["input_ids"][0]
    if len(tokens) <= max_tokens:
        return [text]
    else:
        mid = len(text) // 2
        return split_text_recursively(text[:mid], max_tokens) + split_text_recursively(text[mid:], max_tokens)

# Summarize a chunk of text with a check for minimum length
def summarize_chunk(chunk, max_new_tokens=50, min_length=50):
    try:
        # Skip chunks that are too short to summarize meaningfully
        if len(tokenizer(chunk)["input_ids"]) < 50:
            print(f"Skipping short chunk of length: {len(chunk)}")
            return None

        summary = summarizer(chunk, max_new_tokens=max_new_tokens, min_length=min_length, do_sample=False)
        return summary[0]['summary_text']
    except Exception as e:
        print(f"Error during summarization: {e}")
        return None


# Process and summarize the PDF
def summarize_pdf(pdf_path):
    print(f"Processing {pdf_path}...")
    text = extract_text_from_pdf(pdf_path)
    if not text:
        print("No text found, unable to summarize.")
        return ""
    
    # Initial split into paragraphs or smaller chunks
    initial_chunks = initial_split(text)
    
    # Further split each initial chunk recursively
    final_chunks = []
    for chunk in initial_chunks:
        final_chunks.extend(split_text_recursively(chunk))
    
    summaries = []
    for chunk in final_chunks:
        if chunk.strip():  # Only process non-empty chunks
            print(f"Summarizing chunk of length: {len(chunk)}")
            summary = summarize_chunk(chunk)
            if summary:
                summaries.append(summary)
    
    return " ".join(summaries)

# List of PDFs to summarize
pdf_paths = [out_path, out_path2]

# Summarize each PDF
for pdf_path in pdf_paths:
    summary = summarize_pdf(pdf_path)
    print(f"\nSummary for {os.path.basename(pdf_path)}:\n{summary}\n")
