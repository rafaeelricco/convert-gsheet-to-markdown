import json
import os
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from dotenv import load_dotenv
import pickle
import google.generativeai as genai
import sys
import time
import random
import threading

load_dotenv()

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]


class ProgressBar:
    def __init__(self, total_width=80):
        self.total_width = total_width
        self.current_step = ""
        self.progress = 0
        self._stop_fake_progress = False
        self._is_running_fake = False
        self._current_thread = None
        self._step_printed = False

    def update(self, step, progress=None):
        """Atualiza a barra de progresso com uma nova etapa e progresso"""
        if step != self.current_step:
            if self.current_step and self.progress < 100:
                # Força completar a etapa anterior
                self.update(self.current_step, 100)
            sys.stdout.write(f"\n==> {step}\n")
            self.current_step = step
            self.progress = 0
            self._step_printed = True

        if progress is not None:
            self.progress = min(100, max(0, progress))

            # Calcula o tamanho da barra de progresso
            filled_width = int(self.total_width * self.progress / 100)
            empty_width = self.total_width - filled_width

            # Cria a barra de progresso
            bar = "#" * filled_width + "-" * empty_width

            # Atualiza a barra na mesma linha
            sys.stdout.write(f"\r{bar} {self.progress:.1f}%")
            sys.stdout.flush()

            if self.progress == 100:
                sys.stdout.write("\n")
                sys.stdout.flush()
                self._stop_fake_progress = True

    def start_fake_progress(self, step, start_from=0, until=80):
        """Inicia um progresso falso em background"""
        if self._current_thread and self._current_thread.is_alive():
            self._stop_fake_progress = True
            self._current_thread.join()

        self._stop_fake_progress = False
        self._is_running_fake = True

        # Força a atualização inicial
        self.update(step, start_from)

        def fake_progress_worker():
            current_progress = start_from
            while not self._stop_fake_progress and current_progress < until:
                time.sleep(random.uniform(0.5, 2))
                if not self._stop_fake_progress:
                    current_progress += random.uniform(0.5, 2)
                    self.update(step, current_progress)
            self._is_running_fake = False

        self._current_thread = threading.Thread(target=fake_progress_worker)
        self._current_thread.daemon = True
        self._current_thread.start()

    def wait_for_fake_progress(self):
        """Aguarda a conclusão do progresso falso atual"""
        if self._current_thread and self._current_thread.is_alive():
            self._current_thread.join()


def get_sheet_metadata(progress):
    """Carrega os metadados da planilha do arquivo JSON"""
    progress.start_fake_progress("Carregando metadados da planilha")
    json_path = os.path.join("json", "sheet_info.json")
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    progress.update("Carregando metadados da planilha", 100)
    progress.wait_for_fake_progress()  # Aguarda conclusão
    return data


def get_sheet_data(service, spreadsheet_id, sheet_title, progress):
    """Recupera todos os dados da planilha, incluindo fórmulas e validações"""
    try:
        progress.start_fake_progress("Obtendo dados da planilha")
        result = (
            service.spreadsheets()
            .get(
                spreadsheetId=spreadsheet_id, ranges=[sheet_title], includeGridData=True
            )
            .execute()
        )
        progress.update("Obtendo dados da planilha", 85)

        sheet_data = result["sheets"][0]["data"][0]
        rows = sheet_data.get("rowData", [])
        formatted_data = []

        for row in rows:
            if "values" not in row:
                continue

            row_data = []
            for cell in row["values"]:
                display_value = cell.get("formattedValue", "")
                formula = cell.get("userEnteredValue", {}).get("formulaValue", "")
                data_validation = cell.get("dataValidation", {})
                dropdown_options = data_validation.get("condition", {}).get(
                    "values", []
                )

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

        progress.update("Obtendo dados da planilha", 100)
        progress.wait_for_fake_progress()  # Aguarda conclusão
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


def format_with_gemini(data, progress):
    """Usa a IA do Gemini para formatar os dados de forma mais legível e organizada."""
    try:
        progress.start_fake_progress(
            "Formatando dados com Gemini", start_from=10, until=90
        )

        # Configuração da API do Gemini
        genai.configure(api_key=os.getenv("GEMINI_API_KEY"))
        model = genai.GenerativeModel("gemini-pro")

        # Converte os dados para uma string formatada
        data_str = "\n".join([", ".join(row) for row in data])

        # Prompt para o Gemini
        prompt = f"""Aqui estão os dados de uma planilha:
        {data_str}

        Por favor, formate esses dados de forma mais legível e organizada, destacando informações importantes e removendo redundâncias. Retorne o resultado em formato Markdown.
        """

        # Envia a solicitação para o Gemini
        response = model.generate_content(prompt)
        formatted_data = response.text

        progress.update("Formatando dados com Gemini", 100)
        progress.wait_for_fake_progress()  # Aguarda conclusão
        return formatted_data

    except Exception as e:
        print(f"Erro ao formatar dados com Gemini: {e}")
        return None


def authenticate_google(progress):
    """Autentica o usuário no Google Sheets API."""
    try:
        progress.update("Autenticando no Google", 0)
        creds = None
        token_path = os.path.join("json", "token.pickle")
        client_secrets_path = os.path.join("json", "client_secret.json")

        # Verifica se o arquivo credentials.json existe
        if not os.path.exists(client_secrets_path):
            raise Exception("Arquivo client_secret.json não encontrado na pasta json/")

        # Verifica e carrega o token existente
        if os.path.exists(token_path):
            try:
                with open(token_path, "rb") as token:
                    creds = pickle.load(token)
                progress.update("Autenticando no Google", 30)
            except Exception as e:
                print(f"Erro ao ler token.pickle: {e}")
                os.remove(token_path)
                creds = None

        # Se não houver credenciais válidas ou se estiverem expiradas
        if not creds or not creds.valid:
            progress.update("Autenticando no Google", 50)
            if creds and creds.expired and creds.refresh_token:
                try:
                    creds.refresh(Request())
                except Exception as e:
                    print(f"Erro ao renovar token: {e}")
                    creds = None

            if not creds:
                # Carrega as configurações do cliente OAuth2 do arquivo JSON
                with open(client_secrets_path, "r") as f:
                    client_config = json.load(f)

                flow = InstalledAppFlow.from_client_config(client_config, SCOPES)
                print("\nIniciando processo de autenticação...")
                print(
                    "Adicione o seguinte endereço de redirecionamento no Console do Google Cloud:"
                )
                print("http://localhost:8080/")
                creds = flow.run_local_server(port=8080)

                # Salva as credenciais para uso futuro
                with open(token_path, "wb") as token:
                    pickle.dump(creds, token)

            progress.update("Autenticando no Google", 90)

        progress.update("Autenticando no Google", 100)
        return creds

    except Exception as e:
        progress.update("Erro na autenticação", 100)
        raise Exception(f"Erro na autenticação do Google: {str(e)}")


def main():
    """Função principal do script."""
    progress = ProgressBar()

    try:
        # Autenticação no Google
        creds = authenticate_google(progress)
        service = build("sheets", "v4", credentials=creds)

        # Carrega os metadados da planilha
        sheet_metadata = get_sheet_metadata(progress)
        spreadsheet_id = sheet_metadata["spreadsheet_id"]
        sheet_title = sheet_metadata["sheet_title"]

        # Obtém os dados da planilha
        sheet_data = get_sheet_data(service, spreadsheet_id, sheet_title, progress)

        if not sheet_data:
            print("Nenhum dado foi recuperado da planilha.")
            return

        # Formata os dados em Markdown
        markdown_data = format_to_markdown(sheet_data)
        print("\nDados formatados em Markdown:")
        print(markdown_data)

        # Formata os dados com o Gemini
        gemini_formatted_data = format_with_gemini(sheet_data, progress)
        if gemini_formatted_data:
            print("\nDados formatados com Gemini:")
            print(gemini_formatted_data)

    except Exception as e:
        print(f"\nErro durante a execução do script: {e}")
    finally:
        progress.wait_for_fake_progress()  # Garante que o progresso falso seja finalizado


if __name__ == "__main__":
    main()
