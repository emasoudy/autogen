from flask import Flask, render_template_string, request
from datetime import datetime, timedelta
import win32com.client
import openai


def get_email_action_points(email):
    outlook = win32com.client.Dispatch(
        "Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)

    messages = inbox.Items
    last_seven_days_date = datetime.now() - timedelta(days=7)
    messages = messages.Restrict("[ReceivedTime] >= '{}'".format(
        last_seven_days_date.strftime('%m/%d/%Y')))

    data = ""
    for message in messages:
        if email.lower() in message.SenderName.lower():

            response = openai.Completion.create(
                engine="davinci",
                prompt=message.body,
                temperature=0.5,
                max_tokens=100
            )
            data += f"<p><b>Subject:</b> {message.Subject}</p>"
            data += f"<p><b>Sender:</b> {message.SenderName}</p>"
            data += f"<p><b>Received Time:</b> {message.ReceivedTime}</p>"
            data += f"<p><b>GPT response:</b> {response.choices[0].text.strip()}</p>"
            data += "<hr>"

    return data


app = Flask(__name__)


@app.route('/', methods=['GET', 'POST'])
def home():
    if request.method == 'POST':
        email = request.form.get('email')
        response_data = get_email_action_points(email)
    else:
        response_data = ""

    html = f"""
        <h1>Action Points</h1>
        <form method="POST">
            Email: <input type="text" name="email">
            <input type="submit" value="Submit">
        </form>
        <div>
            {response_data}
        </div>
    """

    return render_template_string(html)


if __name__ == '__main__':
    app.run(debug=True)
