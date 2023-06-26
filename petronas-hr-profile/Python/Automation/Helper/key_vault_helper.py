def get_secret(key_vault_client, secret_name):
    return key_vault_client.get_secret(secret_name).value