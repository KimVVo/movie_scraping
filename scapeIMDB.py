from bs4 import BeautifulSoup
import requests, openpyxl
import re

excel = openpyxl.Workbook()
print(excel.sheetnames)
sheet = excel.active
sheet.title = 'Top 250 Rated Movies by IMDb'
print(excel.sheetnames)
sheet.append(['Movie Rank', 'Movie Name', 'Released Year', 'IMDb Rating'])

# Catch the error
try:  
    
    url = "https://www.imdb.com/chart/top/"
    headers = {
        'User-Agent': 
        'Mozilla/5.0 (iPad; CPU OS 12_2 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Mobile/15E148'
        }

    source = requests.get(url, headers=headers)
    source.raise_for_status() # raise erorr website doesn't exist

    # Convert the data source to soup Object
    soup = BeautifulSoup(source.text, 'html.parser')
    

    # Access all top 250 movies
    movies = soup.find('ul', class_="ipc-metadata-list ipc-metadata-list--dividers-between sc-3f13560f-0 sTTRj compact-list-view ipc-metadata-list--base").find_all('li')

   
        

    for movie in movies:
        
        # Find the movie rank
        rank = re.findall(r'[0-9]+', movie.find('div', class_="ipc-title ipc-title--base ipc-title--title ipc-title-link-no-icon ipc-title--on-textPrimary sc-4dcdad14-9 dZscOy cli-title").a.text)[0]

        # Find the movie name
        name = re.findall(r'[a-zA-Z]+', movie.find('div', class_="ipc-title ipc-title--base ipc-title--title ipc-title-link-no-icon ipc-title--on-textPrimary sc-4dcdad14-9 dZscOy cli-title").a.text)
        name = "".join(str(i) for i in name)
        # Find the movie year 
        year = movie.find('span', class_="sc-4dcdad14-8 cvucyi cli-title-metadata-item").text

        # Find the movie rating
        rating = movie.find('span', class_="ipc-rating-star ipc-rating-star--base ipc-rating-star--imdb ratingGroup--imdb-rating").text.split('\xa0')[0]

        print(rank, name, year, rating)
        sheet.append([rank, name, year, rating])        

           
        
                      

except Exception as e:
    print(e) # print error

# Save the file 
excel.save('IMDb Rating.xlsx')

