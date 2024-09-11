import requests
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import RGBColor

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

# Add each post to the document
for heading, post in all_posts:
    if heading:
        doc.add_heading(heading, level=2)
    
    paragraph = doc.add_paragraph()
    if "Bonus:" in post:
        parts = post.split("Bonus:", 1)
        paragraph.add_run(parts[0])
        bonus_run = paragraph.add_run("Bonus:" + parts[1])
        bonus_run.font.color.rgb = RGBColor(0, 100, 0)  # Dark green color
    else:
        paragraph.add_run(post)

# Save the document
doc.save("scraped_posts4.docx")

print("Posts have been saved to scraped_posts.docx")