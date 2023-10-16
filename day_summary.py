import win32com.client
import openai
import re
from datetime import datetime, timedelta
import configparser
import time


# Your existing code
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)  # "6" refers to the inbox. No subfolder.

messages = inbox.Items

# Get date from user
user_date_str = input("Enter the date (MM/DD/YYYY): ")
user_date = datetime.strptime(user_date_str, '%m/%d/%Y')

# Restrict to items on the user-specified date
messages = messages.Restrict("[ReceivedTime] >= '{}' AND [ReceivedTime] <= '{}'".format(
    user_date.strftime('%m/%d/%Y'),
    (user_date + timedelta(days=1)).strftime('%m/%d/%Y')
))

# Get OpenAI API key
config = configparser.ConfigParser()
config.read('config.ini')

api_key = config['DEFAULT']['api_key']

openai.api_key = api_key

emails_by_sender = {}

for message in messages:
    sender_name = message.SenderName
    email_data = "Subject: " + message.Subject + " | Snippet: " + \
        message.body[:100]  # Shorten the body to a snippet

    # Collect emails by sender
    if sender_name not in emails_by_sender:
        emails_by_sender[sender_name] = email_data
    else:
        emails_by_sender[sender_name] += " | " + \
            email_data  # Use a delimiter to separate emails


def process_batch(sender, email_batch):
    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=[
            {"role": "system", "content": f"You are a helpful assistant. Considering that I am Essam Masoudy known as Abu Abdulaziz, Here are emails from {sender} I received today"},
            {"role": "user", "content": f"Please summarize these emails in short bullet points: {email_batch}"}
        ],
        temperature=0.7,
        max_tokens=100
    )
    return response['choices'][0]['message']['content']


summary_responses = {}

for sender, fullemails in emails_by_sender.items():
    email_batches = fullemails.split(' | ')  # Split emails back into a list
    batch_size = 3  # Adjust batch size as needed
    email_batches = [email_batches[i:i + batch_size]
                     for i in range(0, len(email_batches), batch_size)]

    batch_summaries = []
    for email_batch in email_batches:
        email_batch_str = ' | '.join(email_batch)
        batch_summary = process_batch(sender, email_batch_str)
        batch_summaries.append(batch_summary)

    summary_responses[sender] = ' '.join(batch_summaries)
    print(f'Action points from: {sender} \n')
    print(summary_responses[sender])
    print('\n\n --- Next Sender --- \n\n')
    time.sleep(3)  # Add a delay to avoid hitting the API rate limit)


# import win32com.client
# import openai
# import re
# from datetime import datetime, timedelta

# from test.test_code import OAI_CONFIG_LIST

# # Your existing code
# outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
# inbox = outlook.GetDefaultFolder(6)  # "6" refers to the inbox. No subfolder.

# messages = inbox.Items

# # Get date from user
# user_date_str = input("Enter the date (MM/DD/YYYY): ")
# user_date = datetime.strptime(user_date_str, '%m/%d/%Y')

# # Restrict to items on the user-specified date
# messages = messages.Restrict("[ReceivedTime] >= '{}' AND [ReceivedTime] <= '{}'".format(
#     user_date.strftime('%m/%d/%Y'),
#     (user_date + timedelta(days=1)).strftime('%m/%d/%Y')
# ))

# openai.api_key = OAI_CONFIG_LIST[0]['api_key']

# emails_by_sender = {}

# for message in messages:
#     sender_name = message.SenderName
#     email_data = "\n\n ---Next Email--- \n\nSubject: " + message.Subject + \
#         "\n\nReceived: " + str(message.ReceivedTime) + \
#         "\n\nBody: " + message.body

#     # Collect emails by sender
#     if sender_name not in emails_by_sender:
#         emails_by_sender[sender_name] = email_data
#     else:
#         emails_by_sender[sender_name] += email_data


# def shorten_text(text):
#     try:
#         text = text[:text.lower().index("disclaimer")]
#     except ValueError:
#         pass
#     text = re.sub(r'\s+', ' ', text)
#     return text


# summary_responses = {}

# for sender, fullemails in emails_by_sender.items():
#     shortened_emails = shorten_text(fullemails)
#     response = openai.ChatCompletion.create(
#         model="gpt-3.5-turbo",
#         messages=[
#             {"role": "system", "content": f"You are a helpful assistant. Here are emails from {sender} I received today"},
#             {"role": "user", "content": f"Please summarize these emails: {shortened_emails}"}
#         ],
#         temperature=0.5,
#         max_tokens=100
#     )
#     summary_responses[sender] = response['choices'][0]['message']['content']
#     print(f'Action points from: {sender}')
#     print(summary_responses[sender])

# # Output the summaries
# for sender, summary in summary_responses.items():
#     print(f'Action points from: {sender}')
#     print(summary)


# import win32com.client
# import openai
# import re
# from datetime import datetime, timedelta

# outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
# inbox = outlook.GetDefaultFolder(6)  # "6" refers to the inbox. No subfolder.

# messages = inbox.Items

# last_seven_days_date = datetime.now() - timedelta(days=1)

# # Restrict to items in the last 7 days
# messages = messages.Restrict("[ReceivedTime] >= '{}'".format(
#     last_seven_days_date.strftime('%m/%d/%Y')))

# openai.api_key = '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'

# fullemails = ""

# for message in messages:
#     fullemails += "\n\n ---Next Email--- \n\nSubject: " + message.Subject + \
#         "\n\nReceived: " + str(message.ReceivedTime) + \
#         "\n\nBody: " + message.body


# def shorten_text(text):
#     """
#     This function takes in a text and attempts to reduce its length by removing some less important parts
#     """

#     # If the email has a legal disclaimer, remove it

#     try:
#         text = text[:text.lower().index("disclaimer")]
#     except ValueError:
#         pass

#     # Remove any other sections you deem unnecessary (e.g., headers, footers, etc.)

#     # Remove new lines and extra spaces
#     text = re.sub(r'\s+', ' ', text)

#     return text


# response = openai.ChatCompletion.create(
#     model="gpt-3.5-turbo",
#     messages=[
#         {"role": "system", "content": "You are a helpful assistant.here is a list of emails I received today"},
#         {"role": "user", "content": f"Please summarize my emails in couple of sentences: " +
#             shorten_text(fullemails)}
#     ],
#     temperature=0.5,
#     max_tokens=100
# )
# print(f'Action points from: {message.SenderName}')
# print(response['choices'][0]['message']['content'])
