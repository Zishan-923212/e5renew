import requests
import json
import time

# Register the azure app first and make sure the app has the following permissions:
# files: Files.Read.All、Files.ReadWrite.All、Sites.Read.All、Sites.ReadWrite.All
# user: User.Read.All、User.ReadWrite.All、Directory.Read.All、Directory.ReadWrite.All
# mail: Mail.Read、Mail.ReadWrite、MailboxSettings.Read、MailboxSettings.ReadWrite
# After registration, you must click on behalf of xxx to grant administrator consent, otherwise outlook api cannot be called
refresh_token = "1.AT0AhDaOyDA-OU-EOxAgXUWw-oanOIqkZkBIuQhcdethDZahAAo9AA.AgABAwEAAADW6jl31mB3T7ugrWTT8pFeAwDs_wUA9P_ZxuFf5_tgepihQLCz6oiFQHnRY2Il04FVQbnqUY2L5sFGfaiihoR-utWc6F5FMqI_MRcPHbeKiVkf8V3C_bymshvLpEn3lnahJ-7D8Vc8z-DAUEQjpG-ivLKwlTApu8Y-hGz48VKo814W8mVosCE-ngM_nH-rtxhjkjiDMsov-wTralDuEvQXaBfAKYcrEsPd07cXlodYL2dpymymimQ0HpYRHmdgXXsZaZK3raiv5FtOyjqouQnWQhx2dpD0iYXojpVG2gxoVKHV6O16SJWTqiGNB1Z7KljTuzIESG2Ap_RMm7dfSIjgE5x38aQGp6_iX-53A7YiCYsHXYP6ZxeN_dWzbSXAuLJuD41Jw6YZQjucu82F9TMJDUvpvgqxUJ7gmrQvK6ViYUZ_iSkISaMFBqc1YFygpcX17bclMElZ5U12XnPODJOHKwyVdKVz0Pz-Q3lpiSX5omXfwsHUp3bjRHaoMAtFCCp8S-BQ_vGaFhfaJE6j99_xsVMicl6k6PbBC_QKyRkhaRIesEjny6Bwh3SVzIWYXIfMBj5daqtSbFPtNRFZbD8NBDIPBrzn1TIgOLu_uLd-KoEhIt1FyvhVXuKfvQhyC-2oV68hYq5fEn9sdoW9LtMWQJlZyt7fYxyDo26UEwcqW3llETzK2zUh8Sqdt7V-B25GmGgZH5TOgsvaXX5ywx7fez6YX2DZOSAIH0se_r9utTLPUkwtVrjO2Oj0DYyzhaA0R_P607mfM7AYC5TLINpg9Icq79eCPaDBsKkrqKPvPoHDhfMuBdYERybml_iyZoudTWmbE4STRrH0heOXaQUl-w"
client_id = "8a38a786-66a4-4840-b908-5c75eb610d96"
client_secret = "zQC8Q~NTXqLmfYNISoZ5FrWmDQThwmFRPGL51dr-"




endpoints = [
    'https://graph.microsoft.com/v1.0/me/drive/root',
    'https://graph.microsoft.com/v1.0/me/drive',
    'https://graph.microsoft.com/v1.0/drive/root',
    'https://graph.microsoft.com/v1.0/users',
    'https://graph.microsoft.com/v1.0/me/messages',
    'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules',
    'https://graph.microsoft.com/v1.0/me/drive/root/children',
    'https://api.powerbi.com/v1.0/myorg/apps',
    'https://graph.microsoft.com/v1.0/me/mailFolders',
    'https://graph.microsoft.com/v1.0/me/outlook/masterCategories'
]

def get_access_token(refresh_token, client_id, client_secret):
    headers = {
        'Content-Type': 'application/x-www-form-urlencoded'
    }
    data = {
        'grant_type': 'refresh_token',
        'refresh_token': refresh_token,
        'client_id': client_id,
        'client_secret': client_secret,
        'redirect_uri': 'http://localhost:53682/'
    }
    html = requests.post('https://login.microsoftonline.com/common/oauth2/v2.0/token', data=data, headers=headers)
    jsontxt = json.loads(html.text)
    refresh_token = jsontxt['refresh_token']
    access_token = jsontxt['access_token']
    return access_token

def main():
    access_token = get_access_token(refresh_token, client_id, client_secret)
    session = requests.Session()
    session.headers.update({
        'Authorization': access_token,
        'Content-Type': 'application/json'
    })
    num = 0
    for endpoint in endpoints:
        try:
            response = session.get(endpoint)
            if response.status_code == 200:
                num += 1
                print(f'{num}th Call successful')
        except requests.exceptions.RequestException as e:
            print(e)
            pass
    localtime = time.asctime(time.localtime(time.time()))
    print('The end of this run is :', localtime)

for _ in range(3):
    main()
