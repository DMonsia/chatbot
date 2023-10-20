from time import time

import requests


def conversation_with_powerbi(
    prompt: str, query: str, username: str, password: str
) -> dict[str, str]:
    """Perform yellowsys api llm call"""

    auth = requests.post(
        url="https://labs.uat.chatapi.yellowsys.io/token",
        data={"username": username, "password": password},
    ).json()

    # Change conversation id for each call to avoid history
    _id = f"{time()}".replace(".", "")
    return requests.post(
        url="https://labs.uat.chatapi.yellowsys.io/chatbots/conversationWithpowerBI",
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {auth.get('access_token')}",
        },
        json={
            "temperature": 0,
            "max_tokens": 4096,
            "max_retries": 6,
            "conversation_prompt": prompt,
            "chatbot_name": "azure-openai-chatbot",
            "conversation_id": _id,
            "message": query,
            "metadata": {"user_id": _id},
        },
    ).json()
