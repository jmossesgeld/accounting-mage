from __future__ import print_function
import clicksend_client
from clicksend_client import SmsMessage
from clicksend_client.rest import ApiException


def send_message(message, number):
    # Configure HTTP basic authorization: BasicAuth
    configuration = clicksend_client.Configuration()
    configuration.username = 'jmossesgeld@gmail.com'
    configuration.password = '4BF6CA38-5176-0843-628E-D31244B6E2AD'

    # create an instance of the API class
    api_instance = clicksend_client.SMSApi(
        clicksend_client.ApiClient(configuration))

    # If you want to explictly set from, add the key _from to the message.
    sms_message = SmsMessage('GGEWebSys', message, "+639762537772")

    sms_messages = clicksend_client.SmsMessageCollection(messages=[
                                                         sms_message])

    try:
        # Send sms message(s)
        api_response = api_instance.sms_send_post(sms_messages)
        print(api_response)
    except ApiException as e:
        print("Exception when calling SMSApi->sms_send_post: %s\n" % e)



