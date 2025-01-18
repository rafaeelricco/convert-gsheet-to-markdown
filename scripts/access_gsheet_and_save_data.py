import gspread
import json
import os
import pickle
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from dotenv import load_dotenv

# Carrega as variáveis de ambiente do arquivo .env
load_dotenv()

# Define os escopos necessários
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


def authenticate_google_sheets():
    """
    Autentica no Google Sheets usando as credenciais OAuth2 do arquivo JSON
    """
    creds = None
    token_path = os.path.join("json", "token.pickle")
    client_secrets_path = os.path.join(
        "json",
        "client_secret.json",
    )

    # Verifica se o arquivo de token já existe
    if os.path.exists(token_path):
        with open(token_path, "rb") as token:
            creds = pickle.load(token)

    # Se não houver credenciais válidas ou se estiverem expiradas, solicita a autenticação
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # Carrega as configurações do cliente OAuth2 do arquivo JSON
            with open(client_secrets_path, "r") as f:
                client_config = json.load(f)

            flow = InstalledAppFlow.from_client_config(client_config, SCOPES)
            print("Iniciando processo de autenticação...")
            print(
                "Adicione o seguinte endereço de redirecionamento no Console do Google Cloud:"
            )
            print("http://localhost:8080/")
            creds = flow.run_local_server(port=8080)

            # Salva as credenciais para uso futuro
            with open(token_path, "wb") as token:
                pickle.dump(creds, token)

    # Autoriza o cliente gspread com as credenciais
    client = gspread.authorize(creds)
    return client


def list_sheets_and_save_info(spreadsheet_id, output_file):
    """
    Acessa a planilha, lista as abas e salva as informações em um arquivo JSON.
    """
    # Autenticar
    client = authenticate_google_sheets()

    # Acessar a planilha
    try:
        spreadsheet = client.open_by_key(spreadsheet_id)
    except gspread.SpreadsheetNotFound:
        print(f"Erro: Planilha com ID '{spreadsheet_id}' não encontrada.")
        return
    except Exception as e:
        print(f"Erro ao acessar a planilha: {e}")
        return

    # Listar todas as abas
    worksheets = spreadsheet.worksheets()
    print("\nAbas disponíveis na planilha:")
    for i, sheet in enumerate(worksheets):
        print(f"{i + 1}. {sheet.title} (ID: {sheet.id})")

    # Selecionar uma aba
    while True:
        try:
            choice = int(input("\nEscolha o número da aba que deseja acessar: "))
            if 1 <= choice <= len(worksheets):
                selected_sheet = worksheets[choice - 1]
                break
            else:
                print("Escolha inválida. Tente novamente.")
        except ValueError:
            print("Entrada inválida. Digite um número.")

    # Salvar informações da aba selecionada em um arquivo JSON
    sheet_info = {
        "spreadsheet_id": spreadsheet_id,
        "sheet_id": selected_sheet.id,
        "sheet_title": selected_sheet.title,
    }

    with open(output_file, "w") as f:
        json.dump(sheet_info, f, indent=4)

    print(f"\nInformações da aba '{selected_sheet.title}' salvas em '{output_file}'.")


if __name__ == "__main__":
    try:
        # Configurações
        SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
        if not SPREADSHEET_ID:
            raise ValueError(
                "SPREADSHEET_ID não encontrado nas variáveis de ambiente. Verifique seu arquivo .env."
            )

        OUTPUT_FILE = os.path.join("json", "sheet_info.json")

        # Executar o script
        list_sheets_and_save_info(SPREADSHEET_ID, OUTPUT_FILE)
    except Exception as e:
        print(f"Erro: {str(e)}")
