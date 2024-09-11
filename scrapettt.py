import requests
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import RGBColor
import re

# Base URL
base_url = "https://talktajiktoday.com/word-a-day/page/"

# List to store all posts
all_posts = []

# Loop through all pages
for page_num in range(1, 5):  # Adjust the range as needed
    # Construct the URL for the current page
    url = f"{base_url}{page_num}/"
    
    # Fetch the page content
    response = requests.get(url)
    
    # Check if the request was successful
    if response.status_code == 200:
        # Parse the HTML content
        soup = BeautifulSoup(response.content, 'html.parser')
        
        # Find all posts
        posts = soup.find_all('div', class_='entry-content')
        
        # Extract and store the content of each post
        for post in posts:
            # Extract the heading
            heading_tag = post.find('h2')
            heading = heading_tag.get_text(strip=True) if heading_tag else ""
            
            # Remove the heading from the post content
            if heading_tag:
                heading_tag.decompose()
            
            # Extract the rest of the content
            post_content = post.get_text(separator="\n", strip=True)
            
            # Store the heading and content
            all_posts.append((heading, post_content))
    else:
        print(f"Failed to retrieve page {page_num}")

# Create a new Word document
doc = Document()

# Regular expression to detect Cyrillic characters
cyrillic_pattern = re.compile('[\u0400-\u04FF]+')

# Function to add text with Cyrillic characters in bold
def add_text_with_cyrillic_bold(paragraph, text):
    parts = cyrillic_pattern.split(text)
    matches = cyrillic_pattern.findall(text)
    
    for i, part in enumerate(parts):
        paragraph.add_run(part)
        if i < len(matches):
            run = paragraph.add_run(matches[i])
            run.bold = True

# Add each post to the document
for heading, post in all_posts:
    if heading:
        doc.add_heading(heading, level=2)
    
    paragraph = doc.add_paragraph()
    if "Bonus:" in post:
        parts = post.split("Bonus:", 1)
        
        # Add text with Cyrillic characters in bold
        add_text_with_cyrillic_bold(paragraph, parts[0])

        # Add a blank line before the bonus content
        paragraph.add_run().add_break()
        
        bonus_run = paragraph.add_run("Bonus:" + parts[1])
        bonus_run.font.color.rgb = RGBColor(0, 100, 0)  # Dark green color
    else:
        # Add text with Cyrillic characters in bold
        add_text_with_cyrillic_bold(paragraph, post)

# Save the document
doc.save("scraped_posts7.docx")

print("Posts have been saved to scraped_posts.docx")