import win32com.client
import openai
import datetime

# connect to Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)

emails = inbox.Items
email = emails.GetLast()

# connect to the OpenAI API
openai.api_key = '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'

response = openai.ChatCompletion.create(
    model="gpt-4",
    messages=[
        {"role": "system", "content": "You are a helpful assistant. Please summarize last email I received? \n\nEmail subject: " +
            email.Subject + "\n\nSummary:"},
        {"role": "user", "content": f"Summarize this email body: " + email.Body + ""}
    ],
    temperature=0.5,
    max_tokens=1000
)
print(f'Email subject: {email.Subject}')
print(response['choices'][0]['message']['content'])

del emails
del email
