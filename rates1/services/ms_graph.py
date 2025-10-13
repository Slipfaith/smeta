# services/ms_graph.py

import requests
import pandas as pd
import msal
from io import BytesIO


def authenticate_with_msal(client_id, tenant_id, scope):
    """
    Авторизация через MSAL (interactive/silent).
    Возвращает access_token или None.
    """
    authority = f'https://login.microsoftonline.com/{tenant_id}'
    app = msal.PublicClientApplication(client_id, authority=authority)
    try:
        accounts = app.get_accounts()
        if accounts:
            result = app.acquire_token_silent(scope, account=accounts[0])
        else:
            result = app.acquire_token_interactive(scopes=scope)
        if "access_token" in result:
            return result["access_token"]
        else:
            return None
    except Exception as e:
        print(f"Ошибка в authenticate_with_msal: {e}")
        return None


def download_excel_from_sharepoint(access_token, site_id, file_path):
    """
    Скачивание Excel (старый метод) по Root:/file_path:/
    Возвращаем pd.DataFrame или None.
    """
    if not access_token:
        return None
    try:
        headers = {'Authorization': f'Bearer {access_token}'}
        api_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_path}:/"
        response = requests.get(api_url, headers=headers)
        if response.status_code != 200:
            print(f"download_excel_from_sharepoint: ошибка HTTP {response.status_code}")
            return None

        download_url = response.json().get('@microsoft.graph.downloadUrl')
        if not download_url:
            print("download_excel_from_sharepoint: downloadUrl не найден.")
            return None

        file_response = requests.get(download_url)
        file_response.raise_for_status()
        df = pd.read_excel(BytesIO(file_response.content))
        return df
    except Exception as e:
        print(f"Ошибка при скачивании Excel (старый): {e}")
        return None


def download_excel_by_fileid(access_token, site_id, file_id, sheet_name=None, skiprows=0):
    """
    Скачиваем Excel по fileId.
    Пример:
      site_id = 'logrusit-my.sharepoint.com,xxx,xxx'
      file_id = '0155BTNWDWFO43EDZICVC33ZJ5S2XKDIK7'
    Параметры:
      - sheet_name: например, "TEP (Source RU)"
      - skiprows: 3 (пропустить строки)
    Возвращает pd.DataFrame или None.
    """
    if not access_token:
        return None
    try:
        headers = {'Authorization': f'Bearer {access_token}'}
        api_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{file_id}"
        response = requests.get(api_url, headers=headers)
        if response.status_code != 200:
            print(f"download_excel_by_fileid: ошибка HTTP {response.status_code}")
            return None

        download_url = response.json().get('@microsoft.graph.downloadUrl')
        if not download_url:
            print("download_excel_by_fileid: downloadUrl не найден.")
            return None

        file_response = requests.get(download_url)
        file_response.raise_for_status()

        if sheet_name:
            df = pd.read_excel(BytesIO(file_response.content), sheet_name=sheet_name, skiprows=skiprows, header=None)
        else:
            df = pd.read_excel(BytesIO(file_response.content), skiprows=skiprows, header=None)

        return df
    except Exception as e:
        print(f"Ошибка при скачивании Excel (новый): {e}")
        return None