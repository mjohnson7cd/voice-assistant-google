import requests
import os
import pickle

# Gmail API utils
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from gnewsclient import gnewsclient
# for encoding/decoding messages in base64
from base64 import urlsafe_b64encode
# for dealing with attachement MIME types
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from mimetypes import guess_type as guess_mime_type
from gtts import gTTS
import speech_recognition as sr
import json
import playsound
import feedparser
from pandas import json_normalize
import requests

# Request all access (permission to read/send/receive emails, manage the inbox, and more)
SCOPES = ['https://mail.google.com/']
our_email = 'mjohnson7cd@gmail.com'

r = sr.Recognizer()
language = 'en'

contact_file = open('contacts.json', 'r')
contacts = json.load(contact_file)

topics_file = open('topics.json', 'r')
topics_json = json.load(topics_file)


# gmail authentication handshake
def gmail_authenticate():
    creds = None
    # the file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)
    # if there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # save the credentials for the next run
        with open("token.pickle", "wb") as token:
            pickle.dump(creds, token)
    return build('gmail', 'v1', credentials=creds)


# get the Gmail API service
service = gmail_authenticate()


# Adds the attachment with the given filename to the given message
def add_attachment(message, filename):
    content_type, encoding = guess_mime_type(filename)
    if content_type is None or encoding is not None:
        content_type = 'application/octet-stream'
    main_type, sub_type = content_type.split('/', 1)
    if main_type == 'text':
        fp = open(filename, 'rb')
        msg = MIMEText(fp.read().decode(), _subtype=sub_type)
        fp.close()
    elif main_type == 'image':
        fp = open(filename, 'rb')
        msg = MIMEImage(fp.read(), _subtype=sub_type)
        fp.close()
    elif main_type == 'audio':
        fp = open(filename, 'rb')
        msg = MIMEAudio(fp.read(), _subtype=sub_type)
        fp.close()
    else:
        fp = open(filename, 'rb')
        msg = MIMEBase(main_type, sub_type)
        msg.set_payload(fp.read())
        fp.close()
    filename = os.path.basename(filename)
    msg.add_header('Content-Disposition', 'attachment', filename=filename)
    message.attach(msg)


# Construct messages in correct format for gmail api
def build_message(destination, obj, body, attachment):
    if not attachment:  # no attachments given
        message = MIMEText(body)
        message['to'] = destination
        message['from'] = our_email
        message['subject'] = obj
    else:
        message = MIMEMultipart()
        message['to'] = destination
        message['from'] = our_email
        message['subject'] = obj
        message.attach(MIMEText(body))
        add_attachment(message, attachment)
    return {'raw': urlsafe_b64encode(message.as_bytes()).decode()}


# sends constructed message to gmail
def send_message(service, destination, obj, body, attachments):
    return service.users().messages().send(
        userId="me",
        body=build_message(destination, obj, body, attachments)
    ).execute()


# parses voice input for email elements and sends email to the correct user
def send_message_to(text_input, contact):
    attachment = ''
    body_start = text_input.find('body')
    subject_start = text_input.find('subject')
    attachment_start = text_input.find('attachment')

    if attachment_start:  # google speech to text inserts spaces between filenames and extensions.
        attachment = ((text_input[(attachment_start + 10):len(text_input)]).lower()).replace(' ', '')

    subject = text_input[subject_start + 7: body_start - 1]
    body = text_input[body_start + 4: attachment_start - 1]

    draft_list = ['Sending message to', contact, 'with a subject of', subject, "and a body of", body]
    draft = " ".join(draft_list)
    draft_audio = gTTS(text=draft, lang=language, slow=False)
    draft_audio.save("./draft.mp3")
    playsound.playsound("./draft.mp3")
    os.remove("./draft.mp3")

    if confirm_action():
        if attachment == '':
            send_message(service, contacts[contact], subject, body)
            playsound.playsound('./sent.mp3')
        else:
            send_message(service, contacts[contact], subject, body, attachment)
    else:
        playsound.playsound('./repeat.mp3')
        listen()

    contact_file.close()


#  maps voice input with key in contact.json (needs refactoring)
def find_contact(text_input):
    if "mark" in text_input.lower():
        send_message_to(text_input, "mark")
    if "cody" in text_input.lower():
        send_message_to(text_input, "cody")
    if "von" in text_input.lower():
        send_message_to(text_input, "van")
    if "quan" in text_input.lower():
        send_message_to(text_input, "quan")
    if "alex" in text_input.lower():
        send_message_to(text_input, "alex")
    if "brian" in text_input.lower():
        send_message_to(text_input, "brian")


# ask for user confirmation before preforming chosen action
def confirm_action():
    text_output = "Are you sure you want to preform this action?"

    audio_output = gTTS(text=text_output, lang=language, slow=False)
    audio_output.save('confirm.mp3')
    playsound.playsound("confirm.mp3")
    os.remove("confirm.mp3")

    with sr.Microphone() as response_in:
        print("Listening...")
        input_audio = r.record(response_in, duration=3)
        input_text = r.recognize_google(input_audio, language="en-US")

    if "yes" in input_text.lower():
        return True
    elif "no" in input_text.lower():
        return False


# receives and echo's weather data to user
def get_weather():
    url = 'https://api.weatherapi.com/v1/current.json?key=e30d198fdce14666986200637222505&q=35630'
    response = requests.get(url)
    weather = response.json()

    text_list = ['It is currently', str(weather["current"]["temp_f"]), 'degrees outside. It feels like',
                 str(weather["current"]["feelslike_f"]), 'degrees. The skies are',
                 weather["current"]["condition"]["text"],
                 "with", str(weather["current"]["precip_in"]), "inches of precipitation."]

    text = " ".join(text_list)
    weather_audio = gTTS(text=text, lang=language, slow=False)
    weather_audio.save("./weather.mp3")
    playsound.playsound("./weather.mp3")
    os.remove("./weather.mp3")


# gets rss news feed form google based on topic (top news if no topic is chosen)
def get_news(text_input):
    num_entries = 5
    if 'topic' not in text_input:
        url = 'https://news.google.com/rss'
        news_feed = feedparser.parse(url)
        df_news_feed = json_normalize(news_feed.entries)

        for i in range(num_entries):
            news = df_news_feed.title[i]
            news_audio = gTTS(text=news, lang=language, slow=False)
            news_audio.save("./news.mp3")
            playsound.playsound("./news.mp3")
            os.remove("./news.mp3")
    else:
        topics = ['world', 'local', 'technology', 'entertainment', 'sports', 'science']

        topic_start = text_input.find('topic')
        topic_input = text_input[topic_start + 5: len(text_input)]

        for topic in topics:
            if topic in topic_input:
                print("getting news...")
                url = topics_json[topic]
                news_feed = feedparser.parse(url)
                df_news_feed = json_normalize(news_feed.entries)
                for i in range(num_entries):
                    news = df_news_feed.title[i]
                    news_audio = gTTS(text=news, lang=language, slow=False)
                    news_audio.save("./topic_news.mp3")
                    playsound.playsound("./topic_news.mp3")
                    os.remove("./topic_news.mp3")

    topics_file.close()


# determines what action to preform based on contents to voice input
def find_action(text_input):
    if "email" or "message" in text_input.lower():
        find_contact(text_input)

    if 'weather' in text_input.lower():
        get_weather()

    if 'news' in text_input.lower():
        get_news(text_input)


# listens for words and sends phrases to google recognition api
def listen():
    with sr.Microphone() as source:
        print("Listening...")
        audio_data = r.listen(source)
        print("Working...")
        text = r.recognize_google(audio_data, language="en-US")
        if 'quit' in text:
            exit()
        find_action(text)


# a = Recorder()

listen()
