import requests
import os
import re
from bs4 import BeautifulSoup
import time
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed
import logging
import certifi
from urllib.parse import urljoin
import shutil

# Configurar logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - Thread %(thread)d - %(message)s',
    handlers=[
        logging.FileHandler('relatorio_hapvida.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# URLs
LOGIN_URL = "https://webhap.hapvida.com.br/pls/webhap/webNewTrocaArquivo.loginValidacao"
BASE_URL = "https://webhap.hapvida.com.br/pls/webhap/"

# Número de instâncias padrão
NUM_INSTANCIAS = 8

# Dicionário para armazenar resultados
resultados = {}

class GerenciadorSessao:
    """Classe para gerenciar sessões do requests."""
    def __init__(self, num_instancias, base_download_dir):
        self.num_instancias = num_instancias
        self.base_download_dir = base_download_dir
        self._sessoes = []
        self._download_dirs = []
        self.inicializar_sessoes()

    def inicializar_sessoes(self):
        """Inicializa todas as sessões necessárias."""
        for i in range(self.num_instancias):
            sessao, download_dir = self.configurar_sessao(str(i + 1))
            self._sessoes.append({"sessao": sessao, "em_uso": False, "id": i + 1})
            self._download_dirs.append(download_dir)

    def configurar_sessao(self, sessao_id):
        """Configura uma instância de requests.Session."""
        download_dir = os.path.join(self.base_download_dir, f"driver_{sessao_id}")
        os.makedirs(download_dir, exist_ok=True)
        logger.info(f"Pasta de downloads criada: {download_dir}")
        sessao = requests.Session()
        sessao.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8",
            "Accept-Language": "pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7",
            "Accept-Encoding": "gzip, deflate, br",
            "Referer": "https://webhap.hapvida.com.br/pls/webhap/webNewTrocaArquivo",
            "Origin": "https://webhap.hapvida.com.br",
            "Connection": "keep-alive",
            "Upgrade-Insecure-Requests": "1"
        })
        logger.info(f"Sessão {sessao_id} configurada com sucesso")
        return sessao, download_dir

    def obter_sessao_disponivel(self):
        """Obtém uma sessão disponível do pool."""
        while True:
            for sessao_info in self._sessoes:
                if not sessao_info["em_uso"]:
                    sessao_info["em_uso"] = True
                    return (sessao_info["sessao"], self._download_dirs[sessao_info["id"] - 1], sessao_info["id"])
            time.sleep(0.5)

    def liberar_sessao(self, sessao_id):
        """Libera uma sessão após o uso."""
        for sessao_info in self._sessoes:
            if sessao_info["id"] == sessao_id:
                sessao_info["em_uso"] = False
                break

    def fechar_todos(self):
        """Fecha todas as sessões."""
        for sessao_info in self._sessoes:
            try:
                sessao_info["sessao"].close()
                logger.info(f"Sessão {sessao_info['id']} fechada com sucesso")
            except Exception as e:
                logger.error(f"Erro ao fechar sessão {sessao_info['id']}: {e}")

def processar_login(sessao, download_dir, cnpj, senha, empresa, sessao_id):
    """Processa o login e download de PDFs e CSVs para um CNPJ."""
    try:
        logger.info(f"Sessão {sessao_id} - Iniciando login para CNPJ: {cnpj}")
        payload = {
            "pCpf": cnpj,
            "pSenha": senha,
            "pTokenCaptcha": ""
        }
        sessao.headers.update({"Referer": "https://webhap.hapvida.com.br/pls/webhap/webNewTrocaArquivo"})
        response = sessao.post(LOGIN_URL, data=payload, allow_redirects=True, verify=certifi.where(), timeout=(5, 30))
        logger.info(f"Sessão {sessao_id} - Status do login: {response.status_code}, URL: {response.url}")
        logger.info(f"Sessão {sessao_id} - Cookies após login: {sessao.cookies.get_dict()}")

        if response.status_code != 200 or "Troca de Arquivos - Menu Principal" not in response.text:
            logger.error(f"Sessão {sessao_id} - Falha no login para CNPJ {cnpj}")
            return "❌ Falha no login"

        logger.info(f"Sessão {sessao_id} - Login bem-sucedido")
        soup = BeautifulSoup(response.text, 'html.parser')
        download_link = soup.find('a', string=re.compile(r'BAIXAR ARQUIVOS - DOWNLOAD \(Novo\)'))
        
        if not download_link:
            logger.error(f"Sessão {sessao_id} - Link de download não encontrado")
            return "⚠️ Link de download não encontrado"

        relative_url = download_link['href']
        download_url = urljoin(BASE_URL, relative_url)
        logger.info(f"Sessão {sessao_id} - Acessando URL de download: {download_url}")

        download_response = sessao.get(download_url, allow_redirects=True, verify=certifi.where(), timeout=(5, 30))
        logger.info(f"Sessão {sessao_id} - Status da página de download: {download_response.status_code}")

        if download_response.status_code != 200:
            logger.error(f"Sessão {sessao_id} - Erro ao acessar página de download")
            return "⚠️ Erro ao acessar página de download"

        soup = BeautifulSoup(download_response.text, 'html.parser')
        pdf_links = []
        csv_links = []

        for link in soup.find_all('a', href=True):
            href = link.get('href')
            if re.search(r'\.PDF', href, re.IGNORECASE):
                pdf_links.append((href, link.get_text(strip=True)))
            elif re.search(r'\.CSV', href, re.IGNORECASE):
                csv_links.append((href, link.get_text(strip=True)))

        status = ""

        # Processar PDF
        if pdf_links:
            pdf_links_sorted = []
            pattern = rf'EMPRESA_{re.escape(empresa)}_REMESSA_(\d+)\.PDF'
            for href, _ in pdf_links:
                match = re.search(pattern, href, re.IGNORECASE)
                if match:
                    numero_remessa = int(match.group(1))
                    pdf_links_sorted.append((href, numero_remessa))

            pdf_links_sorted = sorted(pdf_links_sorted, key=lambda x: x[1], reverse=True)
            if pdf_links_sorted:
                href, numero_remessa = pdf_links_sorted[0]
                pdf_url = urljoin(BASE_URL, href) if not href.startswith('http') else href
                filename = f"EMPRESA_{empresa}_REMESSA_{numero_remessa}.pdf"
                filepath = os.path.join(download_dir, filename)

                if os.path.exists(filepath):
                    base, ext = os.path.splitext(filename)
                    filename = f"{base}_2{ext}"
                    filepath = os.path.join(download_dir, filename)

                for attempt in range(2):
                    sessao.headers.update({"Referer": download_url})
                    pdf_response = sessao.get(pdf_url, stream=True, verify=certifi.where(), timeout=(5, 30))
                    logger.info(f"Sessão {sessao_id} - Status do download PDF (tentativa {attempt + 1}): {pdf_response.status_code}")

                    if pdf_response.status_code == 200 and 'application/pdf' in pdf_response.headers.get('Content-Type', ''):
                        with open(filepath, 'wb') as f:
                            for chunk in pdf_response.iter_content(chunk_size=8192):
                                if chunk:
                                    f.write(chunk)
                        if os.path.exists(filepath) and os.path.getsize(filepath) > 0:
                            logger.info(f"Sessão {sessao_id} - PDF salvo: {filepath}")
                            status += "✅ PDF salvo com sucesso; "
                            break
                        else:
                            logger.error(f"Sessão {sessao_id} - PDF vazio ou não criado")
                            status += "⚠️ Falha ao salvar PDF; "
                    else:
                        logger.warning(f"Sessão {sessao_id} - Falha no download do PDF (tentativa {attempt + 1})")
                else:
                    status += "⚠️ Erro ao baixar PDF após retries; "
            else:
                logger.warning(f"Sessão {sessao_id} - Nenhum PDF válido encontrado")
                status += "⚠️ Nenhum PDF encontrado; "
        else:
            logger.warning(f"Sessão {sessao_id} - Nenhum PDF encontrado")
            status += "⚠️ Nenhum PDF encontrado; "

        # Processar CSV
        if csv_links:
            csv_links_sorted = []
            pattern = rf'EMPRESA_{re.escape(empresa)}_REMESSA_(\d+)\.CSV'
            for href, _ in csv_links:
                match = re.search(pattern, href, re.IGNORECASE)
                if match:
                    numero_remessa = int(match.group(1))
                    csv_links_sorted.append((href, numero_remessa))

            csv_links_sorted = sorted(csv_links_sorted, key=lambda x: x[1], reverse=True)
            if csv_links_sorted:
                href, numero_remessa = csv_links_sorted[0]
                csv_url = urljoin(BASE_URL, href) if not href.startswith('http') else href
                filename = f"EMPRESA_{empresa}_REMESSA_{numero_remessa}.csv"
                filepath = os.path.join(download_dir, filename)

                if os.path.exists(filepath):
                    base, ext = os.path.splitext(filename)
                    filename = f"{base}_2{ext}"
                    filepath = os.path.join(download_dir, filename)

                for attempt in range(2):
                    sessao.headers.update({"Referer": download_url})
                    csv_response = sessao.get(csv_url, stream=True, verify=certifi.where(), timeout=(5, 30))
                    logger.info(f"Sessão {sessao_id} - Status do download CSV (tentativa {attempt + 1}): {csv_response.status_code}")

                    if csv_response.status_code == 200:
                        with open(filepath, 'wb') as f:
                            for chunk in csv_response.iter_content(chunk_size=8192):
                                if chunk:
                                    f.write(chunk)
                        if os.path.exists(filepath) and os.path.getsize(filepath) > 0:
                            logger.info(f"Sessão {sessao_id} - CSV salvo: {filepath}")
                            status += "✅ CSV salvo com sucesso"
                            break
                        else:
                            logger.error(f"Sessão {sessao_id} - CSV vazio ou não criado")
                            status += "⚠️ Falha ao salvar CSV"
                    else:
                        logger.warning(f"Sessão {sessao_id} - Falha no download do CSV (tentativa {attempt + 1})")
                else:
                    status += "⚠️ Erro ao baixar CSV após retries"
            else:
                logger.warning(f"Sessão {sessao_id} - Nenhum CSV válido encontrado")
                status += "⚠️ Nenhum CSV encontrado"
        else:
            logger.warning(f"Sessão {sessao_id} - Nenhum CSV encontrado")
            status += "⚠️ Nenhum CSV encontrado"

        return status or "✅ Processado com sucesso"

    except Exception as e:
        logger.error(f"Sessão {sessao_id} - Erro ao processar CNPJ {cnpj}: {e}")
        return f"❌ Erro inesperado: {str(e)}"

def processar_linha_com_pool(args):
    """Função wrapper para processar uma linha usando o pool de threads."""
    gerenciador_sessao, dados_linha = args
    index, row = dados_linha
    cnpj = str(row['cnpj']).strip()
    senha = str(row['senha']).strip()
    empresa = str(row['empresa']).strip()

    sessao, download_dir, sessao_id = gerenciador_sessao.obter_sessao_disponivel()
    try:
        status = processar_login(sessao, download_dir, cnpj, senha, empresa, sessao_id)
        resultados[index] = status
        return index, True
    except Exception as e:
        logger.error(f"Erro ao processar CNPJ {cnpj} com sessão {sessao_id}: {e}")
        resultados[index] = f"❌ Erro: {str(e)}"
        return index, False
    finally:
        gerenciador_sessao.liberar_sessao(sessao_id)

def atualizar_status_excel(excel_path):
    """Atualiza os status na planilha Excel."""
    try:
        df = pd.read_excel(excel_path)
        if 'relatorio' not in df.columns:
            df['relatorio'] = ''
        for index, status in resultados.items():
            df.at[index, 'relatorio'] = status
        df.to_excel(excel_path, index=False)
        logger.info("Status atualizados no Excel com sucesso")
    except Exception as e:
        logger.error(f"Erro ao atualizar status no Excel: {e}")

def mover_arquivos_para_raiz(save_dir):
    """Move todos os arquivos das subpastas driver_X para o diretório raiz."""
    try:
        for subdir in os.listdir(save_dir):
            subdir_path = os.path.join(save_dir, subdir)
            if os.path.isdir(subdir_path) and subdir.startswith("driver_"):
                for arquivo in os.listdir(subdir_path):
                    arquivo_path = os.path.join(subdir_path, arquivo)
                    if os.path.isfile(arquivo_path):
                        destino = os.path.join(save_dir, arquivo)
                        if os.path.exists(destino):
                            base, ext = os.path.splitext(arquivo)
                            i = 1
                            while os.path.exists(os.path.join(save_dir, f"{base}_{i}{ext}")):
                                i += 1
                            destino = os.path.join(save_dir, f"{base}_{i}{ext}")
                        shutil.move(arquivo_path, destino)
                        logger.info(f"Arquivo movido: {arquivo_path} -> {destino}")
                if not os.listdir(subdir_path):
                    os.rmdir(subdir_path)
                    logger.info(f"Subpasta vazia removida: {subdir_path}")
    except Exception as e:
        logger.error(f"Erro ao mover arquivos para o diretório raiz: {e}")

def run_extraction(excel_path, save_dir, due_day, month_number, progress_callback=None):
    """Função principal para extrair relatórios Hapvida."""
    logger.info("Iniciando extração de relatórios Hapvida")
    os.makedirs(save_dir, exist_ok=True)

    if not os.path.exists(excel_path):
        logger.error("Arquivo Excel não encontrado!")
        raise FileNotFoundError("Arquivo Excel não encontrado!")
    if not os.path.isdir(save_dir):
        logger.error("Diretório de salvamento inválido!")
        raise NotADirectoryError("Diretório de salvamento inválido!")

    try:
        df = pd.read_excel(excel_path, dtype=str)
        df['cnpj'] = df['cnpj'].astype(str).str.strip()
        df['senha'] = df['senha'].astype(str).str.strip()
        df['empresa'] = df['empresa'].astype(str).str.strip()
        total_linhas = len(df)
        logger.info(f"Total de linhas a processar: {total_linhas}")

        gerenciador_sessao = GerenciadorSessao(NUM_INSTANCIAS, save_dir)
        with ThreadPoolExecutor(max_workers=NUM_INSTANCIAS) as executor:
            futures = [
                executor.submit(processar_linha_com_pool, (gerenciador_sessao, (index, row)))
                for index, row in df.iterrows()
            ]
            processed = 0
            for future in as_completed(futures):
                processed += 1
                try:
                    index, sucesso = future.result()
                    percentage = int((processed / total_linhas) * 100)
                    if progress_callback:
                        progress_callback(percentage, processed, total_linhas)
                    logger.info(f"Progresso: {processed}/{total_linhas} ({percentage}%)")
                except Exception as e:
                    logger.error(f"Erro ao process Jonto os processar future: {e}")

        atualizar_status_excel(excel_path)
        mover_arquivos_para_raiz(save_dir)
        logger.info("Extração de relatórios concluída")
    except Exception as e:
        logger.critical(f"Erro crítico na extração de relatórios: {e}")
        raise
    finally:
        gerenciador_sessao.fechar_todos()

if __name__ == "__main__":
    excel_path = r"C:\Users\chrys\Downloads\SINTICLEPEMPa.xlsx"
    save_dir = r"C:\Users\chrys\OneDrive\Desktop\Nova pasta (3)"
    due_day = "20"
    month_number = "04"
    def dummy_progress_callback(percentage, processed, total):
        print(f"Progresso: {processed}/{total} ({percentage}%)")
    run_extraction(excel_path, save_dir, due_day, month_number, dummy_progress_callback)