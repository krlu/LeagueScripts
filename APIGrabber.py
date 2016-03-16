import io
import os

from PIL import Image
import certifi
import requests
import urllib3

#################################################
########### Basic calls to riot API ############# 
#################################################
def requestItemData(APIKey):
    URL = "https://global.api.pvp.net/api/lol/static-data/na/v1.2/item?api_key=" + APIKey
    response = requests.get(URL)
    return response.json()

def requestChampionData(APIKey):
    URL = "https://global.api.pvp.net/api/lol/static-data/na/v1.2/champion?api_key=" + APIKey
    response = requests.get(URL)
    return response.json()

def requestItemStatsByID(Id, APIKey):
    URL = "https://global.api.pvp.net/api/lol/static-data/na/v1.2/item/" + Id + "?itemData=stats&api_key=" + APIKey
    response = requests.get(URL)
    return response.json() 
 
def getChampionImage(championKey):
    url = 'http://ddragon.leagueoflegends.com/cdn/5.24.2/img/champion/' + championKey + '.png'
    http = urllib3.PoolManager(cert_reqs='CERT_REQUIRED', ca_certs=certifi.where())
    img_file = http.urlopen('GET',url, preload_content=False)
    image_file = io.BytesIO(img_file.read())
    return image_file

def getItemStats(APIKey):
    itemStats = []
    itemData = requestItemData(APIKey)
    for key in itemData["data"]:
        print(requestItemStatsByID(key, APIKey)['name'])
        print(requestItemStatsByID(key, APIKey)['stats'])
        itemStats.append(requestItemStatsByID(key, APIKey))
    return itemStats
    
def getItemImage(itemID):
    url = 'http://ddragon.leagueoflegends.com/cdn/6.1.1/img/item/'+ itemID +'.png'
    http = urllib3.PoolManager(cert_reqs='CERT_REQUIRED', ca_certs=certifi.where())
    img_file = http.urlopen('GET',url, preload_content=False)

    if img_file.status == 404:
        print("404 page not found")
        return -1   
    else:
        image_file = io.BytesIO(img_file.read())
        return image_file

def downloadItemImages(APIKey, directory):
    json =  requestItemData(APIKey)
    for i in json['data']:
        key = json['data'][i]['id']
        name = json['data'][i]['name']
        image_file = getItemImage(str(key))
        if image_file != -1:
            im = Image.open(image_file)
            if not os.path.exists(directory + '/itemImages/'):
                os.makedirs(directory + '/itemImages/')
            im.save(directory + '/itemImages/' + name + '.png')   
##############################################
########### More specialized Methods########## 
##############################################
def downloadChampionImages(APIKey, directory):
    json2 = requestChampionData(APIKey)
    for i in json2['data']:
        key = json2['data'][i]['key']
        name = json2['data'][i]['name']
        image_file = getChampionImage(key)
        im = Image.open(image_file)
        if not os.path.exists(directory +'/championImages/'):
            os.makedirs(directory +'/championImages/')
        im.save(directory +'/championImages/' + name + '.png') 

def main():
    #APIKey = (str)(input('Copy and paste your API Key here: '))
    #directory = os.path.dirname(__file__)
    getItemStats("7c5aaed2-90a2-4aa8-895a-4948eddac1a9")

    
if __name__ == "__main__":
    main()

