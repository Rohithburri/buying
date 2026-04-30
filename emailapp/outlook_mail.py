import os
import requests
import msal


def get_access_token():
    TENANT_ID = os.environ.get("TENANT_ID")
    CLIENT_ID = os.environ.get("CLIENT_ID")
    CLIENT_SECRET = os.environ.get("CLIENT_SECRET")

    authority = f"https://login.microsoftonline.com/{TENANT_ID}"

    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=authority,
        client_credential=CLIENT_SECRET
    )

    token = app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )

    return token.get("access_token")




def send_outlook_mail(to_email, subject, message):
    access_token = get_access_token()

    SENDER_EMAIL = os.environ.get("SENDER_EMAIL")

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    email_msg = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": "Text",
                "content": message
            },
            "toRecipients": [
                {
                    "emailAddress": {
                        "address": to_email
                    }
                }
            ]
        }
    }

    url = f"https://graph.microsoft.com/v1.0/users/{SENDER_EMAIL}/sendMail"

    response = requests.post(url, headers=headers, json=email_msg)

    print("MAIL STATUS:", response.status_code)
    print("MAIL RESPONSE:", response.text)

    return response.status_code