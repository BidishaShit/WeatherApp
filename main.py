import requests
import json
import win32com.client as wincom

city = input("Enter the name of the city\n")

url = f"https://api.weatherapi.com/v1/current.json?key=b13989793f184149a91141538230103&q={city}"

r = requests.get(url)

weather_dic = json.loads(r.text)
w = weather_dic["current"]["temp_c"]
w2 = weather_dic["current"]["temp_f"]
w3 = weather_dic["current"]["last_updated"]
w4 = weather_dic["location"]["region"]
w5 = weather_dic["current"]["wind_kph"]
w6 = weather_dic["current"]["humidity"]
w7 = weather_dic["current"]["pressure_in"]
w8 = weather_dic["current"][("feelslike_c")]
w9 = weather_dic["location"]["country"]

speak = wincom.Dispatch("SAPI.SpVoice")
#f"This is {w9}. Welcome to {w4}. "
text = (f"The current weather in {city} is {w} degrees and {w2} in fareinheit. It feels like {w8}. Last updated {w3}. Wind speed is {w5} km per hour. Humidity is {w6}. Pressure is {w7}")

speak.Speak(text)


