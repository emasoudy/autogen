from flask import Flask, render_template_string, request
import win32com.client
import openai
import re
import pythoncom
import format_gpt_output
from datetime import datetime, timedelta


def get_email_action_points(person):
    pythoncom.CoInitialize()
    data = ""
    outlook = win32com.client.Dispatch(
        "Outlook.Application").GetNamespace("MAPI")
    # "6" refers to the inbox. No subfolder.
    inbox = outlook.GetDefaultFolder(6)

    messages = inbox.Items

    last_seven_days_date = datetime.now() - timedelta(days=1)

    # Restrict to items in the last 7 days
    messages = messages.Restrict("[ReceivedTime] >= '{}'".format(
        last_seven_days_date.strftime('%m/%d/%Y')))

    openai.api_key = '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'

    fullemails = ""

    for message in messages:
        if person in message.SenderName:
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
    data += f"<p><b>GPT response:</b> {response['choices'][0]['message']['content']}</p>"
    data += "<hr>"

    return data


app = Flask(__name__)


@app.route('/', methods=['GET', 'POST'])
def home():
    if request.method == 'POST':
        person = request.form.get('person')
        response_data = get_email_action_points(person)
    else:
        response_data = ""

    html = f"""
        <h1>Action Points</h1>
        <form method="POST">
            Person: <input type="text" name="person">
            <input type="submit" value="Submit">
        </form>
        <div>
            {response_data}
        </div>
    """

    return render_template_string(html)


if __name__ == '__main__':
    app.run(debug=True)
