import sys
import requests
import urllib

apikey = 'a3f100cd733c52dddf0504a4e494a08f'
# sendername = '[YOUR SENDERNAME]'


def send_message(message, number):
    print('Sending Message...')
    params = (
        ('apikey', apikey),
        # ('sendername', sendername),
        ('message', message),
        ('number', number)
    )
    path = 'https://semaphore.co/api/v4/messages?' + urllib.parse.urlencode(params)
    return requests.post(url=path, headers={'Allow-Control-Allow-Origin': '*'}).json()