import json


def format_gpt_output(response):
    formatted_string = ""
    if "choices" in response:
        for choice in response["choices"]:
            formatted_string += "<p><b>Prompt:</b> " + \
                choice["finish_reason"] + "<br>" + \
                "<b>Response:</b> " + choice["text"] + "</p>"
    return formatted_string
