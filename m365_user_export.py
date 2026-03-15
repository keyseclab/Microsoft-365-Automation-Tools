import requests
import pandas as pd

# Questo script utilizza Microsoft Graph API per esportare la lista utenti
# Nota: Richiede un'applicazione registrata su Azure AD con permessi User.Read.All

def get_m365_users(access_token):
    url = "https://graph.microsoft.com/v1.0/users"
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    
    response = requests.get(url, headers=headers)
    
    if response.status_status == 200:
        users_data = response.json().get('value', [])
        return users_data
    else:
        print(f"Errore nella richiesta: {response.status_code}")
        return None

def save_to_excel(users_list):
    if not users_list:
        print("Nessun dato da salvare.")
        return
    
    df = pd.DataFrame(users_list)
    # Selezioniamo solo le colonne principali per pulizia
    columns_to_keep = ['displayName', 'mail', 'userPrincipalName', 'jobTitle']
    df_filtered = df[df.columns.intersection(columns_to_keep)]
    
    df_filtered.to_excel("m365_users_report.xlsx", index=False)
    print("Report generato con successo: m365_users_report.xlsx")

if __name__ == "__main__":
    # Placeholder per il token (da ottenere tramite OAuth2)
    TEMP_TOKEN = "INSERISCI_IL_TUO_TOKEN_QUI"
    print("Avvio esportazione utenti Microsoft 365...")
    # users = get_m365_users(TEMP_TOKEN)
    # save_to_excel(users)