# Download the helper library from https://www.twilio.com/docs/python/install
import os
from twilio.rest import Client
from flask import Flask, request, redirect
from twilio.twiml.messaging_response import MessagingResponse

'''
try:

except Exception as e:
    print(e)

'''

# Find your Account SID and Auth Token at twilio.com/console
# and set the environment variables. See http://twil.io/secure


############# 변수 설정
# 메일 내용


# SMS 발송
try:
    account_sid = 'AC477d0e9cf7e151cfa77ab02304819c2e'
    auth_token = '6a181bf2a174bec0095f31c12bec6871'
    client = Client(account_sid, auth_token)
    body_text = "어떻게"

    message = client.messages.create(to='+821076293345', from_='+17164568703', body=body_text)

except Exception as e:
    print(e)


'''
    app = Flask(__name__)
    receive_text = client.incoming_phone_numbers.create(phone_number='+17164568703', sms_method='POST', sms_url='https://demo.twilio.com/welcome/sms/reply/')
    print(receive_text)
    if __name__ == "__main__":
        app.run(debug=True)
'''

'''
# SMS 수신

try:
    app = Flask(__name__)

    @app.route("/sms", methods=['GET', 'POST'])
    def incoming_sms():
        """Send a dynamic reply to an incoming text message"""
        # Get the message the user sent our Twilio number
        body = request.values.get('Body', None)

        # Start our TwiML response
        resp = MessagingResponse()

        # Determine the right reply for this message
        if body == 'hello':
            resp.message("Hi!")
        elif body == 'bye':
            resp.message("Goodbye")

        return str(resp)

    if __name__ == "__main__":
        app.run(debug=True)

except Exception as e:
    print(e)
'''