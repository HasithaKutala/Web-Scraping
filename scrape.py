from bs4 import BeautifulSoup
import requests, openpyxl

excel=openpyxl.Workbook()

print(excel.sheetnames)

sheet=excel.active

sheet.title="Top Rated Movies"

print(excel.sheetnames)

sheet.append(["Movie Rank","Movie Name", "Year of Release", "IMDB Rating"])




try:
    #Add website link
    source=requests.get("https://www.imdb.com/chart/toptv/")
    #If there is any mkstake in website, to detect we use raise_for_status()
    source.raise_for_status()
    
    #Beautifulsoup is to extract data from html documents. it will parse the html code of webpage 
    soup= BeautifulSoup(source.text,"html.parser")
    
    movies=soup.find('tbody',class_="lister-list").find_all("tr")
    
    #print(len(movies))
    
    for movie in movies:
        
        content= movie.find('td',class_="titleColumn") # it gives you full content present in td
        #if you want only movie name then it is in <a> so type .a.text
        moviename=movie.find('td',class_="titleColumn").a.text
        #we are entering into tr tag then td inside it then a tag to fetch movie name 
        rank=movie.find('td',class_="titleColumn").get_text(strip=True).split('.')[0]
        #to get ranking of the movie, we used strip to remove spaces and split(.) to split the text based on dot and then we need 0 index as it is rank
        year=movie.find('td',class_="titleColumn").span.text.strip('()')
        #to get rating, it is in other td tag 
        rating= movie.find('td',class_="ratingColumn imdbRating").strong.text
        
        
        print(rank,moviename, year,rating)
        
        sheet.append([rank,moviename, year,rating])
        
        #if you use break statement you will get one movie detail.
except Exception as e:
    print(e)


excel.save("IMDB Movie Ratings.xlsx")
      
