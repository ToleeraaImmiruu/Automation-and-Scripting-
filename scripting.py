from bs4 import BeautifulSoup
import requests

url = "https://www.amazon.com/s?k=playstation+4"
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
}
response = requests.get(url, headers=headers)
soup = BeautifulSoup(response.content, "html.parser")
# Find the FIRST product link
first_product = soup.find("a", class_="a-link-normal s-line-clamp-2 puis-line-clamp-3-for-col-4-and-8 s-link-style a-text-normal")

# From that link, go inside h2 → span → text
title = first_product.find("h2").text.strip()
price_div = soup.find("div", attrs={"data-cy": "secondary-offer-recipe"})
price = price_div.find("span", class_="a-color-base").text
print(price)
