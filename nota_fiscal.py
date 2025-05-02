import os
import time
import pandas as pd
import logging
import requests
from bs4 import BeautifulSoup
from concurrent.futures import ThreadPoolExecutor, as_completed
from openpyxl import load_workbook
import shutil
import re
from urllib.parse import urlparse, parse_qs
import certifi

# Configuração do logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - Thread %(thread)d - %(message)s',
    handlers=[
        logging.FileHandler('automacao_nfe.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# URLs dos sites
WEBHAP_URL = "https://webhap.hapvida.com.br/pls/webhap/pk_nota_fiscal.login"
WEBHAP_LOGIN_URL = "https://webhap.hapvida.com.br/pls/webhap/pk_nota_fiscal.loginValidacao"
WEBHAP_DATA_URL = "https://webhap.hapvida.com.br/pls/webhap/pk_nota_fiscal.listaNotasFiscais"
ISS_URL = "https://iss.fortaleza.ce.gov.br/grpfor/pagesPublic/validarNota.seam"
ISS_VALIDATE_URL = "https://iss.fortaleza.ce.gov.br/grpfor/pagesPublic/validarNota.seam"
ISS_CONSULT_URL = "https://iss.fortaleza.ce.gov.br/grpfor/pagesPublic/consultarNota.seam"

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
            "Referer": ISS_URL,
            "Origin": "https://iss.fortaleza.ce.gov.br",
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

def atualizar_status_excel(excel_path):
    """Atualiza os status na planilha Excel."""
    try:
        wb = load_workbook(excel_path)
        ws = wb.active
        for index, status in resultados.items():
            ws.cell(row=index + 2, column=7, value=status)
        wb.save(excel_path)
        logger.info("Status atualizados no Excel com sucesso")
    except Exception as e:
        logger.error(f"Erro ao atualizar status no Excel: {e}")

def processar_empresa(sessao, download_dir, empresa, senha, data_vencimento, index, sessao_id, is_retry=False):
    """Processa o download da nota fiscal para uma empresa."""
    try:
        empresa = str(empresa).strip()
        senha = str(senha).strip()
        logger.info(f"Sessão {sessao_id} - Iniciando processamento da empresa: {empresa}")
        # Validação do código da empresa
        if empresa == "00100":
            logger.error(f"Sessão {sessao_id} - Código da empresa inválido: {empresa}")
            resultados[index] = "ERRO: CÓDIGO INVÁLIDO (00100)"
            return
        # Login no Hapvida
        sessao.headers.update({"Referer": WEBHAP_URL, "Origin": "https://webhap.hapvida.com.br"})
        login_payload = {
            "pCodigoEmpresa": empresa,
            "pNm_login": "",
            "pSenha": senha,
            "pTokenCaptcha": ""
        }
        login_response = sessao.post(WEBHAP_LOGIN_URL, data=login_payload, allow_redirects=True, verify=certifi.where(), timeout=(5, 30))
        logger.info(f"Sessão {sessao_id} - Status do login: {login_response.status_code}, URL: {login_response.url}")
        logger.info(f"Sessão {sessao_id} - Cookies após login: {sessao.cookies.get_dict()}")
        if login_response.status_code != 200 or "pk_nota_fiscal.login" in login_response.url:
            logger.error(f"Sessão {sessao_id} - Falha no login para empresa {empresa}, URL: {login_response.url}")
            resultados[index] = "ERRO_LOGIN"
            return
        logger.info(f"Sessão {sessao_id} - Login bem-sucedido para empresa {empresa}")
        # Extrair pIdSessao da resposta do login
        soup_login = BeautifulSoup(login_response.text, 'html.parser')
        link_com_sessao = soup_login.find('a', href=re.compile(r'pIdSessao'))
        pIdSessao = None
        if link_com_sessao:
            parsed_url = urlparse(link_com_sessao['href'])
            params = parse_qs(parsed_url.query)
            pIdSessao = params.get('pIdSessao', [None])[0]
            logger.info(f"Sessão {sessao_id} - pIdSessao extraído: {pIdSessao}")
        else:
            logger.warning(f"Sessão {sessao_id} - Nenhum link com pIdSessao encontrado na página de login")
        # Acessar a lista de notas fiscais
        params = {"pIdSessao": pIdSessao} if pIdSessao else {}
        lista_response = sessao.get(WEBHAP_DATA_URL, params=params, allow_redirects=True, verify=certifi.where(), timeout=(5, 30))
        logger.info(f"Sessão {sessao_id} - Status da requisição: {lista_response.status_code}, URL final: {lista_response.url}")
        logger.debug(f"Sessão {sessao_id} - Primeiros 500 caracteres do HTML: {lista_response.text[:500]}")
        if lista_response.status_code != 200:
            logger.error(f"Sessão {sessao_id} - Erro ao acessar lista de notas fiscais")
            resultados[index] = "ERRO_LISTA_NOTAS"
            return
        # Parsear a tabela
        soup = BeautifulSoup(lista_response.text, 'html.parser')
        tabela = soup.find('table', class_='table-rel')
        if not tabela:
            logger.error(f"Sessão {sessao_id} - Tabela de notas fiscais não encontrada")
            resultados[index] = "ERRO: TABELA NÃO ENCONTRADA"
            return
        # Encontrar o link diretamente pelo texto da data
        link_encontrado = None
        logger.info(f"Sessão {sessao_id} - Buscando data de vencimento: {data_vencimento}")
        for link in tabela.find_all('a'):
            texto_link = link.get_text(strip=True)
            logger.info(f"Sessão {sessao_id} - Verificando link com texto: {texto_link}")
            if texto_link == data_vencimento:
                link_encontrado = link
                logger.info(f"Sessão {sessao_id} - Link com data {data_vencimento} encontrado!")
                break
        if not link_encontrado:
            datas_disponiveis = [link.get_text(strip=True) for link in tabela.find_all('a')]
            logger.error(f"Sessão {sessao_id} - Data de vencimento {data_vencimento} não encontrada. Datas disponíveis: {datas_disponiveis}")
            resultados[index] = "ERRO: DATA NÃO ENCONTRADA"
            return
        # Extrair parâmetros do link
        href = link_encontrado['href']
        parsed_url = urlparse(href)
        params = parse_qs(parsed_url.query)
        pIdSessao = params.get('pIdSessao', [None])[0]
        nfs_e = params.get('pnu_nfse', [None])[0]
        cod_verificacao = params.get('pcd_verificacao', [None])[0]
        if not (pIdSessao and nfs_e and cod_verificacao):
            logger.error(f"Sessão {sessao_id} - Parâmetros pIdSessao, pnu_nfse ou pcd_verificacao não encontrados")
            resultados[index] = "ERRO: PARÂMETROS AUSENTES"
            return
        logger.info(f"Sessão {sessao_id} - Extraído: NFS-e={nfs_e}, Código={cod_verificacao}, pIdSessao={pIdSessao}")
        # Simular clique no link para validar a sessão
        data_response = sessao.get(WEBHAP_DATA_URL, params={
            "pIdSessao": pIdSessao,
            "pnu_nfse": nfs_e,
            "pcd_verificacao": cod_verificacao
        }, allow_redirects=True, verify=certifi.where(), timeout=(5, 30))
        logger.info(f"Sessão {sessao_id} - Status do acesso aos detalhes: {data_response.status_code}")
        if data_response.status_code != 200:
            logger.error(f"Sessão {sessao_id} - Erro ao acessar detalhes da nota fiscal")
            resultados[index] = "ERRO_ACESSO_NOTA"
            return
        # Acessar a página de validação para obter o ViewState
        sessao.headers.update({"Referer": ISS_URL, "Origin": "https://iss.fortaleza.ce.gov.br"})
        init_response = sessao.get(ISS_VALIDATE_URL, allow_redirects=True, verify=False, timeout=(5, 30))
        logger.info(f"Sessão {sessao_id} - Status do GET inicial: {init_response.status_code}")
        if init_response.status_code != 200:
            logger.error(f"Sessão {sessao_id} - Erro ao acessar página de validação inicial")
            resultados[index] = "ERRO_VALIDACAO_INICIAL"
            return
        soup_init = BeautifulSoup(init_response.text, 'html.parser')
        view_state = soup_init.find('input', {'name': 'javax.faces.ViewState'})['value'] if soup_init.find('input', {'name': 'javax.faces.ViewState'}) else None
        if not view_state:
            logger.error(f"Sessão {sessao_id} - ViewState não encontrado na página inicial de validação")
            resultados[index] = "ERRO: VIEWSTATE INICIAL AUSENTE"
            return
        logger.info(f"Sessão {sessao_id} - ViewState inicial extraído: {view_state}")
        # Validar no ISS
        cnpj = "63.554.067/0001-98"  # CNPJ fixo
        validate_payload = {
            "validarNotaForm": "validarNotaForm",
            "validarNotaForm:opConsulta": "0",  # Nota Fiscal Eletrônica
            "validarNotaForm:opPrestadorNF": "1",  # CNPJ
            "validarNotaForm:numNfse": nfs_e,
            "validarNotaForm:numCodVerificacao": cod_verificacao,
            "validarNotaForm:nfseCnpjPrestador": cnpj,
            "validarNotaForm:j_id92": "Consultar",  # Botão Consultar
            "javax.faces.ViewState": view_state
        }
        logger.debug(f"Sessão {sessao_id} - Payload de validação: {validate_payload}")
        validate_response = sessao.post(ISS_VALIDATE_URL, data=validate_payload, allow_redirects=False, verify=False, timeout=(5, 30))
        logger.debug(f"Sessão {sessao_id} - Status do POST de validação: {validate_response.status_code}")
        # Seguir redirecionamento manualmente
        if validate_response.status_code == 302:
            redirect_url = validate_response.headers.get('Location')
            if redirect_url:
                full_redirect_url = f"https://iss.fortaleza.ce.gov.br{redirect_url}" if redirect_url.startswith('/') else redirect_url
                logger.info(f"Sessão {sessao_id} - Seguindo redirecionamento para: {full_redirect_url}")
                validate_response = sessao.get(full_redirect_url, allow_redirects=True, verify=False, timeout=(5, 30))
                logger.debug(f"Sessão {sessao_id} - Status após redirecionamento: {validate_response.status_code}, URL: {validate_response.url}")
            else:
                logger.error(f"Sessão {sessao_id} - Cabeçalho de redirecionamento ausente")
                resultados[index] = "ERRO: REDIRECIONAMENTO AUSENTE"
                return
        # Salvar HTML de validação para debug
    
        # Verificar mensagens de erro
        soup_validate = BeautifulSoup(validate_response.text, 'html.parser')
        mensagens = soup_validate.find('dl', id='mensagens')
        if mensagens and mensagens.get_text(strip=True):
            logger.error(f"Sessão {sessao_id} - Mensagem de erro na validação: {mensagens.get_text(strip=True)}")
            resultados[index] = f"ERRO_VALIDACAO: {mensagens.get_text(strip=True)}"
            return
        # Validar se a página contém a nota fiscal
        panel_acoes = soup_validate.find(lambda tag: tag.name in ['div', 'table'] and tag.get('id', '').endswith('panelAcoes'))
        form_download = soup_validate.find('form', id='j_id32')
        nf_info = soup_validate.find('div', class_='container-fluid preview')
        # Logar presença dos elementos
        logger.info(f"Sessão {sessao_id} - Elementos encontrados: panelAcoes={bool(panel_acoes)}, form_download={bool(form_download)}, nf_info={bool(nf_info)}")
        if panel_acoes:
            logger.info(f"Sessão {sessao_id} - ID do panelAcoes: {panel_acoes.get('id')}")
        # Procurar por qualquer div ou table com 'panel' no ID como fallback
        if not panel_acoes:
            panel_alternativo = soup_validate.find(lambda tag: tag.name in ['div', 'table'] and 'panel' in tag.get('id', '').lower())
            logger.info(f"Sessão {sessao_id} - Panel alternativo encontrado: {panel_alternativo.get('id') if panel_alternativo else 'Nenhum'}")
        if validate_response.status_code != 200:
            logger.error(f"Sessão {sessao_id} - Resposta inválida da validação (status {validate_response.status_code})")
            resultados[index] = "ERRO_VALIDACAO_STATUS"
            return
        if not form_download:
            logger.error(f"Sessão {sessao_id} - Formulário 'j_id32' não encontrado na página de validação")
            resultados[index] = "ERRO: FORMULÁRIO AUSENTE"
            return
        if not nf_info:
            logger.warning(f"Sessão {sessao_id} - Informações da nota fiscal (container-fluid preview) não encontradas na página")
        logger.info(f"Sessão {sessao_id} - Página da NF validada com sucesso")
        # Extrair ViewState para download
        view_state = soup_validate.find('input', {'name': 'javax.faces.ViewState'})['value'] if soup_validate.find('input', {'name': 'javax.faces.ViewState'}) else None
        if not view_state:
            logger.error(f"Sessão {sessao_id} - ViewState não encontrado na página de validação")
            resultados[index] = "ERRO: VIEWSTATE AUSENTE"
            return
        logger.info(f"Sessão {sessao_id} - ViewState para download: {view_state}")
        # Preparar payload para download
        download_payload = {
            "j_id32": "j_id32",
            "j_id32:j_id33": "Exportar PDF",
            "javax.faces.ViewState": view_state
        }
        logger.debug(f"Sessão {sessao_id} - Payload de download: {download_payload}")
        # Baixar o PDF
        for attempt in range(2):
            sessao.headers.update({"Referer": validate_response.url, "Origin": "https://iss.fortaleza.ce.gov.br"})
            download_response = sessao.post(ISS_CONSULT_URL, data=download_payload, stream=True, verify=False, timeout=(5, 30))
            logger.info(f"Sessão {sessao_id} - Status do download (tentativa {attempt + 1}): {download_response.status_code}")
            logger.info(f"Sessão {sessao_id} - Content-Type: {download_response.headers.get('Content-Type', '')}")
            if download_response.status_code == 200 and 'application/pdf' in download_response.headers.get('Content-Type', ''):
                break
            logger.warning(f"Sessão {sessao_id} - Tentativa {attempt + 1} de download falhou (status {download_response.status_code}, Content-Type: {download_response.headers.get('Content-Type', '')})")
            time.sleep(1)
        if download_response.status_code != 200 or 'application/pdf' not in download_response.headers.get('Content-Type', ''):
            logger.error(f"Sessão {sessao_id} - Erro ao baixar PDF após retries (status {download_response.status_code}, Content-Type: {download_response.headers.get('Content-Type', '')})")
            resultados[index] = "ERRO_DOWNLOAD_FINAL"
            return
        arquivo_temp = os.path.join(download_dir, f"temp_{sessao_id}.pdf")
        with open(arquivo_temp, 'wb') as f:
            for chunk in download_response.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
        if os.path.exists(arquivo_temp) and os.path.getsize(arquivo_temp) > 0:
            novo_nome = os.path.join(download_dir, f"{empresa}.pdf")
            if os.path.exists(novo_nome):
                base, ext = os.path.splitext(novo_nome)
                i = 1
                while os.path.exists(f"{base}_{i}{ext}"):
                    i += 1
                novo_nome = f"{base}_{i}{ext}"
            os.rename(arquivo_temp, novo_nome)
            logger.info(f"Sessão {sessao_id} - PDF salvo: {novo_nome}")
            resultados[index] = "SUCESSO"
        else:
            logger.error(f"Sessão {sessao_id} - Arquivo PDF vazio ou não criado")
            resultados[index] = "ERRO_DOWNLOAD_FINAL" if is_retry else "ERRO_DOWNLOAD"
    except Exception as e:
        logger.error(f"Sessão {sessao_id} - Erro ao processar empresa {empresa}: {e}")
        resultados[index] = "ERRO: NF NÃO ENCONTRADA" if not is_retry else "ERRO_FINAL: NF NÃO ENCONTRADA"

def processar_empresa_com_pool(args):
    """Função wrapper para processar uma empresa usando o pool de threads."""
    gerenciador_sessao, dados_empresa = args
    index, row = dados_empresa
    empresa = str(row['empresa']).strip()
    senha = str(row['senha']).strip()
    data_vencimento = row['data_vencimento']
    sessao, download_dir, sessao_id = gerenciador_sessao.obter_sessao_disponivel()
    try:
        processar_empresa(sessao, download_dir, empresa, senha, data_vencimento, index, sessao_id)
        return index, True
    except Exception as e:
        logger.error(f"Erro ao processar empresa {empresa} com sessão {sessao_id}: {e}")
        return index, False
    finally:
        gerenciador_sessao.liberar_sessao(sessao_id)

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
    """Função principal para executar a extração de notas fiscais."""
    logger.info("Iniciando extração de notas fiscais")
    os.makedirs(save_dir, exist_ok=True)
    if not os.path.exists(excel_path):
        logger.error("Arquivo Excel não encontrado!")
        raise FileNotFoundError("Arquivo Excel não encontrado!")
    if not os.path.isdir(save_dir):
        logger.error("Diretório de salvamento inválido!")
        raise NotADirectoryError("Diretório de salvamento inválido!")
    try:
        df = pd.read_excel(excel_path, dtype=str)
        df['empresa'] = df['empresa'].astype(str).str.strip()
        df['senha'] = df['senha'].astype(str).str.strip()
        df['data_vencimento'] = f"{due_day}/{month_number}/25"
        total_empresas = len(df)
        logger.info(f"Total de empresas a processar: {total_empresas}")
        gerenciador_sessao = GerenciadorSessao(NUM_INSTANCIAS, save_dir)
        with ThreadPoolExecutor(max_workers=NUM_INSTANCIAS) as executor:
            futures = [
                executor.submit(processar_empresa_com_pool, (gerenciador_sessao, (index, row)))
                for index, row in df.iterrows()
            ]
            processed = 0
            for future in as_completed(futures):
                processed += 1
                try:
                    index, sucesso = future.result()
                    percentage = int((processed / total_empresas) * 100)
                    if progress_callback:
                        progress_callback(percentage, processed, total_empresas)
                    logger.info(f"Progresso: {processed}/{total_empresas} ({percentage}%)")
                except Exception as e:
                    logger.error(f"Erro ao processar future: {e}")
        atualizar_status_excel(excel_path)
        logger.info("Extração de notas fiscais concluída")
        mover_arquivos_para_raiz(save_dir)
    except Exception as e:
        logger.critical(f"Erro crítico na extração de notas fiscais: {e}")
        raise
    finally:
        gerenciador_sessao.fechar_todos()

if __name__ == "__main__":
    # Exemplo de uso standalone para testes
    excel_path = r"C:\Users\chrys\Downloads\SINTICLEPEMPa.xlsx"
    save_dir = r"C:\Users\chrys\OneDrive\Desktop\Nova pasta (3)"
    due_day = "20"
    month_number = "04"
    def dummy_progress_callback(percentage, processed, total):
        print(f"Progresso: {processed}/{total} ({percentage}%)")
    run_extraction(excel_path, save_dir, due_day, month_number, dummy_progress_callback)