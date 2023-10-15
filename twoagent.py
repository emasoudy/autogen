from autogen import AssistantAgent, UserProxyAgent, config_list_from_json
# Load LLM inference endpoints from an env variable or a file
# See https://microsoft.github.io/autogen/docs/FAQ#set-your-api-endpoints
# and OAI_CONFIG_LIST_sample
config_list = config_list_from_json(env_or_file="OAI_CONFIG_LIST")
# You can also set config_list directly as a list, for example, config_list = [{'model': 'gpt-4', 'api_key': '<your OpenAI API key here>'},]
assistant = AssistantAgent("assistant", llm_config={
                           "config_list": config_list})
user_proxy = UserProxyAgent(
    "user_proxy", code_execution_config={"work_dir": "coding"})
user_proxy.initiate_chat(
    assistant, message="I need you to connect my eamils from outlook using com method to gpt4 API so i can ask gpt4 for any information i need. e.g. what are the action points received from specific person in last 7 days?")
# This initiates an automated chat between the two agents to solve the task
