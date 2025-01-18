import json
import os
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from dotenv import load_dotenv
import pickle

load_dotenv()

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]


def get_sheet_metadata():
    """Carrega os metadados da planilha do arquivo JSON"""
    json_path = os.path.join("json", "sheet_info.json")
    with open(json_path, "r", encoding="utf-8") as f:
        return json.load(f)


def get_sheet_data(service, spreadsheet_id, sheet_title):
    """Recupera todos os dados da planilha, incluindo fórmulas e validações"""
    try:
        # Obtém os dados básicos
        result = (
            service.spreadsheets()
            .get(
                spreadsheetId=spreadsheet_id, ranges=[sheet_title], includeGridData=True
            )
            .execute()
        )

        sheet_data = result["sheets"][0]["data"][0]
        rows = sheet_data.get("rowData", [])

        formatted_data = []

        for row in rows:
            if "values" not in row:
                continue

            row_data = []
            for cell in row["values"]:
                # Obtém o valor formatado ou o valor bruto
                display_value = cell.get("formattedValue", "")

                # Obtém a fórmula se existir
                formula = cell.get("userEnteredValue", {}).get("formulaValue", "")

                # Obtém as opções de dropdown se existirem
                data_validation = cell.get("dataValidation", {})
                dropdown_options = data_validation.get("condition", {}).get(
                    "values", []
                )

                # Monta o valor final da célula
                cell_value = display_value
                if formula:
                    cell_value += f" [formula: {formula}]"
                if dropdown_options:
                    options = [
                        opt.get("userEnteredValue", "") for opt in dropdown_options
                    ]
                    cell_value += f" [opções: {', '.join(options)}]"

                row_data.append(cell_value)

            formatted_data.append(row_data)

        return formatted_data

    except Exception as e:
        print(f"Erro ao acessar a planilha: {e}")
        return None


def format_to_markdown(data):
    """Formata os dados em uma tabela markdown"""
    if not data or len(data) == 0:
        return "Nenhum dado encontrado."

    # Primeira linha como cabeçalho
    headers = data[0]
    content = data[1:]

    # Cria o cabeçalho da tabela
    markdown = "| " + " | ".join(str(h) for h in headers) + " |\n"
    markdown += "|" + "|".join(["---"] * len(headers)) + "|\n"

    # Adiciona as linhas de dados
    for row in content:
        markdown += "| " + " | ".join(str(cell) for cell in row) + " |\n"

    return markdown


def main():
    """Função principal"""
    try:
        # Carrega os metadados da planilha
        sheet_info = get_sheet_metadata()
        spreadsheet_id = sheet_info["spreadsheet_id"]
        sheet_title = sheet_info["sheet_title"]

        # Autentica e cria o serviço
        creds = None
        token_path = os.path.join("json", "token.pickle")
        client_secrets_path = os.path.join(
            "json",
            "client_secret.json",
        )

        if os.path.exists(token_path):
            with open(token_path, "rb") as token:
                creds = pickle.load(token)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                with open(client_secrets_path, "r") as f:
                    client_config = json.load(f)

                flow = InstalledAppFlow.from_client_config(client_config, SCOPES)
                creds = flow.run_local_server(port=8080)

                with open(token_path, "wb") as token:
                    pickle.dump(creds, token)

        service = build("sheets", "v4", credentials=creds)

        # Obtém os dados da planilha
        data = get_sheet_data(service, spreadsheet_id, sheet_title)

        if data:
            # Converte para markdown
            markdown_table = format_to_markdown(data)

            # Salva em um arquivo na pasta output
            output_file = os.path.join(
                "output", f"{sheet_title.lower().replace(' ', '_')}_table.md"
            )
            with open(output_file, "w", encoding="utf-8") as f:
                f.write(markdown_table)

            print(f"Tabela markdown salva em {output_file}")
            print("\nVisualização da tabela:")
            print(markdown_table)

    except Exception as e:
        print(f"Erro durante a execução: {e}")


if __name__ == "__main__":
    main()
