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
            sys.stdout.write(f"\r==> {step}\n")
            self.current_step = step
            self.progress = 0
            self._step_printed = True

        if progress is not None:
            self.progress = min(100, max(0, progress))

            # Calcula o tamanho da barra de progresso
            filled_width = int(self.total_width * self.progress / 100)
            empty_width = self.total_width - filled_width

            # Cria a barra de progresso
            bar = "#" * filled_width + " " * empty_width

            # Atualiza a barra na mesma linha
            sys.stdout.write(f"\r{bar} {self.progress:.1f}%")
            sys.stdout.flush()

            if self.progress == 100:
                sys.stdout.write("\n")
                sys.stdout.flush()
                self._stop_fake_progress = True
                self._step_printed = False

    def start_fake_progress(self, step, start_from=0, until=80):
        """Inicia um progresso falso em background"""
        # Aguarda a conclusão de qualquer thread anterior
        if self._current_thread and self._current_thread.is_alive():
            self._stop_fake_progress = True
            self._current_thread.join()

        self._stop_fake_progress = False
        self._is_running_fake = True

        def fake_progress_worker():
            current_progress = start_from
            if not self._step_printed:
                self.update(step, current_progress)
            while not self._stop_fake_progress and current_progress < until:
                time.sleep(
                    random.uniform(1, 5)
                )  # Intervalo aleatório entre 1 e 5 segundos
                if not self._stop_fake_progress:
                    current_progress += 1
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
    """Usa a IA do Gemini para formatar e organizar os dados"""
    try:
        progress.start_fake_progress(
            "Formatando dados com Gemini", start_from=0, until=70
        )
        genai.configure(api_key=os.getenv("GEMINI_API_KEY"))
        model = genai.GenerativeModel("gemini-1.5-pro")

        data_str = "\n".join([",".join(map(str, row)) for row in data])
        prompt = f"""
        Por favor, organize e formate os seguintes dados em uma tabela markdown bem estruturada.
        Os dados estão em formato CSV, onde a primeira linha representa os cabeçalhos.
        Mantenha todas as informações originais, incluindo fórmulas e opções de dropdown.
        
        Dados:
        {data_str}
        """

        progress.update("Formatando dados com Gemini", 80)
        response = model.generate_content(prompt)
        progress.update("Formatando dados com Gemini", 100)
        progress.wait_for_fake_progress()  # Aguarda conclusão
        return response.text

    except Exception as e:
        print(f"Erro ao usar Gemini AI: {e}")
        return format_to_markdown(data)


def main():
    """Função principal"""
    try:
        progress = ProgressBar()

        # Carrega os metadados da planilha
        sheet_info = get_sheet_metadata(progress)
        spreadsheet_id = sheet_info["spreadsheet_id"]
        sheet_title = sheet_info["sheet_title"]

        # Autentica e cria o serviço
        creds = None
        token_path = os.path.join("json", "token.pickle")
        client_secrets_path = os.path.join("json", "client_secret.json")

        if os.path.exists(token_path):
            with open(token_path, "rb") as token:
                creds = pickle.load(token)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                progress.start_fake_progress("Atualizando credenciais")
                creds.refresh(Request())
                progress.update("Atualizando credenciais", 100)
            else:
                progress.start_fake_progress("Realizando nova autenticação")
                with open(client_secrets_path, "r") as f:
                    client_config = json.load(f)

                flow = InstalledAppFlow.from_client_config(client_config, SCOPES)
                creds = flow.run_local_server(port=8080)

                with open(token_path, "wb") as token:
                    pickle.dump(creds, token)
                progress.update("Realizando nova autenticação", 100)

        service = build("sheets", "v4", credentials=creds)

        # Obtém os dados da planilha
        data = get_sheet_data(service, spreadsheet_id, sheet_title, progress)

        if data:
            # Usa o Gemini para formatar os dados
            markdown_table = format_with_gemini(data, progress)

            # Salva em um arquivo na pasta output
            progress.start_fake_progress("Salvando arquivo markdown")
            output_file = os.path.join(
                "output", f"{sheet_title.lower().replace(' ', '_')}_table.md"
            )
            with open(output_file, "w", encoding="utf-8") as f:
                f.write(markdown_table)
            progress.update("Salvando arquivo markdown", 100)

            print(f"\nTabela markdown salva em {output_file}")
            print("\nVisualização da tabela:")
            print(markdown_table)

    except Exception as e:
        print(f"\nErro durante a execução: {e}")


if __name__ == "__main__":
    main()
