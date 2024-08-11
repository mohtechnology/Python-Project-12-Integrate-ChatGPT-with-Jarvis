import speech_recognition as sr
import win32com.client

speaker = win32com.client.Dispatch('SAPI.SpVoice')

def speak(text):
    print(text)
    speaker.speak(text)

def takecommand():
    r = sr.Recognizer()
    with sr.Microphone(device_index = 0) as source:
        print('Listening...')
        r.pause_threshold = 1
        audio = r.listen(source)
        print("Recognizing...")
        try:
            text = r.recognize_google(audio, language="en-in")
            print(f"user said:{text}")
            return text
        except Exception as e:
            print("Some Error occurred . Sorry")
    return ''

def open_website(website):
    import webbrowser

    sites_dict =  {"youtube": "https://www.youtube.com", 
            "wikipedia": "https://www.wikipedia.com",
            "google": "https://www.google.com",
            "facebook": "https://www.facebook.com"}

    if website in sites_dict:
        url = sites_dict[website]
        speak(f"Opening {website} Sir")
    else:
        url = f"https://www.google.com/search?q={website}"
        speak(f"Searching {website} Sir")

    webbrowser.open(url)

def open_app(app):
    import subprocess

    try:
        subprocess.run(app)
        speak(f"Opening {app} Sir")
    except:
        open_website(app)

def ai(query):
    import openai

    openai.api_key = "Enter-Your-API-Key"
    try:
        response = openai.Completion.create(
            engine = 'gpt-3.5-turbo',
            prompt = query
        )
        speak(response.choices[0].text.strip())
    except Exception as e:
        speak(e)

def jarvis():
    while 1:
        command = takecommand()
        if "hello" in command or 'good morning' in command:
            speak("Good Morning Boss How Can I Help You Today?")
        elif "your name" in command:
            speak("I Am Jarvis Your Virtual Assistant")
        elif 'time' in command:
            from datetime import datetime
            now = datetime.now()
            current_time = now.strftime("%I:%M:%S %p")
            speak(f"The Current Time Is : {current_time}")
        elif 'date' in command:
            from datetime import datetime
            now = datetime.today()
            current_date = now.strftime("%Y-%m-%d")
            speak(f"The Date Is : {current_date}")
        elif 'open' in command:
            website = command.split('open ')[-1].lower()
            open_app(website)
        elif "exit" in command or "bye" in command or 'good night' in command:
            speak("Good Night Boss")
            break
        elif command == "":
            speak("Say Something Sir")
        else:
            ai(command)

if __name__ == "__main__":
    jarvis()
