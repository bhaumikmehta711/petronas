import requests

def post(url, header, payload):
    try:
        response = requests.post(
            url = url,
            headers=header,
            json=payload
        )
        return response
    except Exception as e:
        raise ValueError(e)