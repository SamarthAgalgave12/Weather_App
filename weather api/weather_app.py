# importing the requests for web searching and win32com as wincom for speaking
import requests
import win32com.client as wincom
speak = wincom.Dispatch("SAPI.SpVoice")

while True:
    # taking name of city as input from the user and creating its url with f string
    city=input("Enter The Name Of The City : ")
    url=f'http://api.weatherapi.com/v1/current.json?key=0a54868b60be46ab866104643242808&q={city}'

    # making command stop code
    if city=='q':
        speak.Speak("See You Next Time Thank yOU fOR using our weather app")
        exit()
    # sending get request and storing it in get 
    get = requests.get(url)
    # create data to store information in get
    data=get.json()

    # resolving them
    if "error" in data:
        speak.Speak("City Not Found, Please Enter Correct City Name")
        continue
    else:
        location = data['location']['name']
        region = data['location']['region']
        country = data['location']['country']
        temperature_c = data['current']['temp_c']
        condition = data['current']['condition']['text']
        time=data['location']["localtime"]
        

    # storing them in weather info
    weather_info = (
            f"The Current Weather in {location}, {region}, {country} is "
            f"{temperature_c} Degrees Celsius With {condition}."
            f"This is Seen At Time And Date : {time}"
        )
    print(weather_info)
    speak.Speak(weather_info)
