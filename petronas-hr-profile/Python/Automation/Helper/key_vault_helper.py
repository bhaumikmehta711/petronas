def get_secret(key_vault_client, secret_name):
    try:
        return key_vault_client.get_secret(secret_name).value
    except Exception as e:
        raise ValueError(e)