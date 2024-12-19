import os
import requests
from bs4 import BeautifulSoup
import re

def sanitize_filename(filename):
    sanitized = re.sub(r'[<>:"/\\|?*]', '', filename)
    sanitized = sanitized.replace(' ', '_')
    
    max_length = 255
    return sanitized[:max_length]

def download_image(url, save_path):
    try:
        response = requests.get(url)
        response.raise_for_status()
        with open(save_path, 'wb') as file:
            file.write(response.content)
        print(f"Downloaded: {save_path}")
    except Exception as e:
        print(f"Error downloading {url}: {e}")

def scrape_product_page(url):
    try:
        response = requests.get(url)
        response.raise_for_status()
        
        soup = BeautifulSoup(response.text, 'html.parser')
        
        title_element = soup.select_one('.product-detail__title')
        if not title_element:
            print(f"No title found for {url}")
            return
        
        title = sanitize_filename(title_element.get_text(strip=True))
        
        os.makedirs(title, exist_ok=True)
        
        main_image = soup.select_one('.product-detail__images-container img')
        if main_image and main_image.get('src'):
            main_image_url = main_image.get('src')
            if not main_image_url.startswith('http'):
                main_image_url = f"https:{main_image_url}"
            download_image(main_image_url, os.path.join(title, 'main_image.jpg'))
        
        thumbnails = soup.select('.product-detail__thumbnails a img')
        for i, thumb in enumerate(thumbnails, 1):
            thumb_url = thumb.get('src')
            if not thumb_url:
                continue
            
            if not thumb_url.startswith('http'):
                thumb_url = f"https:{thumb_url}"
            
            download_image(thumb_url, os.path.join(title, f'thumbnail_{i}.jpg'))
        
        print(f"Processed: {title}")
    
    except Exception as e:
        print(f"Error processing {url}: {e}")

def main():
    urls = [
        'https://kbdfans.com/products/f1-8x-v2'
    ]

    for url in urls:
        scrape_product_page(url)

if __name__ == "__main__":
    main()