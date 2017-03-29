# -*- coding: utf-8 -*-
"""
Created on Sat Mar 11 20:28:07 2017

@author: Kumar.Singh
"""

import onedrivesdk

redirect_uri = 'urn:ietf:wg:oauth:2.0:oob'
client_secret = '0EE9D0935B755202DB7E053A0D1A6080DF37AD6B'
client_id='b51307ea-b9c4-42f9-99c1-f8e495ac3c67'
api_base_url='https://api.onedrive.com/v1.0/'
scopes=['wl.signin', 'wl.offline_access', 'onedrive.readwrite']

http_provider = onedrivesdk.HttpProvider()
auth_provider = onedrivesdk.AuthProvider(    http_provider=http_provider,    client_id=client_id,    scopes=scopes)

client = onedrivesdk.OneDriveClient(api_base_url, auth_provider, http_provider)
auth_url = client.auth_provider.get_auth_url(redirect_uri)
# Ask for the code
print('Paste this URL into your browser, approve the app\'s access.')
print('Copy everything in the address bar after "code=", and paste it below.')
print(auth_url)
code = 'M8d843909-7ce6-3a78-5372-d5be80e3f1a5'

client.auth_provider.authenticate(code, redirect_uri, client_secret)

import onedrivesdk
from onedrivesdk.helpers import GetAuthCodeServer

redirect_uri = 'http://localhost:8080/signin-microsoft'
client_secret = '0EE9D0935B755202DB7E053A0D1A6080DF37AD6B'
scopes=['wl.signin', 'wl.offline_access', 'onedrive.readwrite']

client = onedrivesdk.get_default_client(    client_id='b51307ea-b9c4-42f9-99c1-f8e495ac3c67', scopes=scopes)

auth_url = client.auth_provider.get_auth_url(redirect_uri)

#this will block until we have the code
code = GetAuthCodeServer.get_auth_code(auth_url, redirect_uri)

client.auth_provider.authenticate(code, redirect_uri, client_secret)