import requests

def post(url, header, payload):
    response = requests.post(
        url = url,
        headers=header,
        data=payload
    )
    return response