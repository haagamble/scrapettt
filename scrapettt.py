import requests
from bs4 import BeautifulSoup
from docx import Document

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
            post_content = post.get_text(separator="\n", strip=True)
            all_posts.append(post_content)
    else:
        print(f"Failed to retrieve page {page_num}")

# Create a new Word document
doc = Document()

# Add each post to the document
for post in all_posts:
    doc.add_paragraph(post)

# Save the document
doc.save("scraped_posts.docx")

print("Posts have been saved to scraped_posts.docx")