import requests
import json
import win32com.client as wincom

API_key = ""
city = input("Enter city name: ")
print()

url = f"https://api.openweathermap.org/data/2.5/weather?q={city}&appid={API_key}&units=metric"

weather_data = requests.get(url)
weather_dict = json.loads(weather_data.text)

# Setting up TTS
engine = wincom.Dispatch("SAPI.SpVoice")

# Printing begins
degree_sign = u'\N{DEGREE SIGN}'

if weather_dict["cod"] == "404":
    print(weather_dict["message"])

elif weather_dict["cod"] == 200:
    print(f"Displaying weather information for {weather_dict["name"]}, {weather_dict["sys"]["country"]}")
    print(f"Temperature: {weather_dict["main"]["temp"]}{degree_sign}C")
    print(f"Feels like: {weather_dict["main"]["feels_like"]}{degree_sign}C")
    print(f"Humidity: {weather_dict["main"]["humidity"]}%")
    engine.Speak(f"Displaying weather information for {weather_dict["name"]}")

else:
    print("Something went wrong!")

print()
print("End of program")
