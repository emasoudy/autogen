import win32com.client
import openai
import re
from datetime import datetime, timedelta

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)  # "6" refers to the inbox. No subfolder.

messages = inbox.Items

last_seven_days_date = datetime.now() - timedelta(days=1)

# Restrict to items in the last 7 days
messages = messages.Restrict("[ReceivedTime] >= '{}'".format(
    last_seven_days_date.strftime('%m/%d/%Y')))

openai.api_key = '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'

fullemails = ""

for message in messages:
    if 'Khalid A. Alghligah' in message.SenderName:
        fullemails += "\n\n ---Next Email--- \n\nSubject: " + message.Subject + \
            "\n\nReceived: " + str(message.ReceivedTime) + \
            "\n\nBody: " + message.body


def shorten_text(text):
    """
    This function takes in a text and attempts to reduce its length by removing some less important parts
    """

    # If the email has a legal disclaimer, remove it

    try:
        text = text[:text.lower().index("disclaimer")]
    except ValueError:
        pass

    # Remove any other sections you deem unnecessary (e.g., headers, footers, etc.)

    # Remove new lines and extra spaces
    text = re.sub(r'\s+', ' ', text)

    return text


response = openai.ChatCompletion.create(
    model="gpt-3.5-turbo",
    messages=[
        {"role": "system", "content": "You are a helpful assistant.here is a list of emails I received from: " + message.SenderName},
        {"role": "user", "content": f"Please summarize action points or requests in bullet points for each subject: " +
            shorten_text(fullemails)}
    ],
    temperature=0.5,
    max_tokens=100
)
print(f'Action points from: {message.SenderName}')
print(response['choices'][0]['message']['content'])
