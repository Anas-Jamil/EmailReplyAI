import win32com.client as win32
import openai
import datetime

openai.api_key = ""

today = datetime.date.today()
today_str = today.strftime("%Y-%m-%d")

def generate_gpt_response(prompt):
    try:
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",  # Use the appropriate model name
            messages=[
                {"role": "system", "content": "You are a helpful assistant."},
                {"role": "user", "content": prompt}
            ],
            max_tokens=200  # Adjust max tokens as needed
        )
        return response.choices[0].message['content'].strip()
    except Exception as e:
        print(f"Error generating GPT-3 response: {e}")
        return ""
    
def getMail():
    olApp = win32.Dispatch('Outlook.Application')
    olNS = olApp.GetNamespace('MAPI')

    inbox = olNS.GetDefaultFolder(6)  
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)
    most_recent_email = messages.GetFirst()

    return most_recent_email

def generate_reply(email):
    original_body = email.Body
    gpt_response = generate_gpt_response(f"Generate a reply to match the tone of the following email. Only include the new email u write. DO NOT INCLUDE SUBJETCT IN THE TEXT: {original_body}" + today_str)
    return gpt_response

def mail_reply(text):
    olApp = win32.Dispatch('Outlook.Application')
    olNS = olApp.GetNamespace('MAPI')
    
    inbox = olNS.GetDefaultFolder(6)  
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)
    most_recent_email = messages.GetFirst()
    
    reply = most_recent_email.Reply()
    reply.Body = text + "_______________________________________________________________________________________________________________________" + reply.Body
    
    reply.Display()


