import os
import sys
import json
import logging
from datetime import datetime
import traceback

# Importar os scripts de extração
import relatorio_hapvida
import nota_fiscal
import boleto_hapvida

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('extractus_bridge.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

def atualizar_progresso(tipo, porcentagem, processados, total):
    """Função para atualizar o progresso da extração e enviar para o frontend"""
    progresso = {
        "tipo": tipo,
        "porcentagem": porcentagem,
        "processados": processados,
        "total": total,
        "timestamp": datetime.now().isoformat()
    }
    # Imprimir o progresso como JSON para ser capturado pelo frontend
    print(json.dumps({"progresso": progresso}))
    sys.stdout.flush()

def listar_arquivos_gerados(save_dir):
    """Função para listar os arquivos gerados no diretório de salvamento"""
    try:
        arquivos = [
            {"nome": arquivo, "caminho": os.path.join(save_dir, arquivo)}
            for arquivo in os.listdir(save_dir)
            if os.path.isfile(os.path.join(save_dir, arquivo))
        ]
        return arquivos
    except Exception as e:
        logger.error(f"Erro ao listar arquivos gerados: {str(e)}")
        return []

def executar_extracao(excel_path, save_dir, due_day, month_number, tipos):
    """Função principal que executa a extração dos documentos selecionados"""
    resultados = {
        "boleto": 0,
        "nota": 0,
        "relatorio": 0,
        "erros": [],
        "arquivos_gerados": []
    }
    
    try:
        # Verificar se o arquivo Excel existe
        if not os.path.exists(excel_path):
            raise FileNotFoundError(f"Arquivo Excel não encontrado: {excel_path}")
        
        # Verificar se o diretório de salvamento existe, se não, criar
        os.makedirs(save_dir, exist_ok=True)
        
        # Processar cada tipo de documento selecionado
        for tipo in tipos:
            if tipo == "boleto":
                logger.info("Iniciando extração de boletos")
                try:
                    boleto_hapvida.run_extraction(
                        excel_path=excel_path,
                        save_dir=save_dir,
                        due_day=due_day,
                        month_number=month_number,
                        progress_callback=lambda p, proc, tot: atualizar_progresso("boleto", p, proc, tot)
                    )
                    # Contar arquivos PDF gerados
                    boletos_gerados = sum(1 for arquivo in os.listdir(save_dir) if arquivo.endswith('.pdf'))
                    resultados["boleto"] = boletos_gerados
                    logger.info(f"Extração de boletos concluída. {boletos_gerados} boletos gerados.")
                except Exception as e:
                    logger.error(f"Erro na extração de boletos: {str(e)}")
                    resultados["erros"].append(f"Boletos: {str(e)}")
            
            elif tipo == "nota":
                logger.info("Iniciando extração de notas fiscais")
                try:
                    nota_fiscal.run_extraction(
                        excel_path=excel_path,
                        save_dir=save_dir,
                        due_day=due_day,
                        month_number=month_number,
                        progress_callback=lambda p, proc, tot: atualizar_progresso("nota", p, proc, tot)
                    )
                    # Contar arquivos PDF gerados
                    notas_geradas = sum(1 for arquivo in os.listdir(save_dir) if arquivo.endswith('.pdf'))
                    resultados["nota"] = notas_geradas
                    logger.info(f"Extração de notas fiscais concluída. {notas_geradas} notas geradas.")
                except Exception as e:
                    logger.error(f"Erro na extração de notas fiscais: {str(e)}")
                    resultados["erros"].append(f"Notas Fiscais: {str(e)}")
            
            elif tipo == "relatorio":
                logger.info("Iniciando extração de relatórios")
                try:
                    relatorio_hapvida.run_extraction(
                        excel_path=excel_path,
                        save_dir=save_dir,
                        due_day=due_day,
                        month_number=month_number,
                        progress_callback=lambda p, proc, tot: atualizar_progresso("relatorio", p, proc, tot)
                    )
                    # Contar arquivos PDF e CSV gerados
                    relatorios_gerados = sum(1 for arquivo in os.listdir(save_dir) if arquivo.endswith(('.pdf', '.csv')))
                    resultados["relatorio"] = relatorios_gerados
                    logger.info(f"Extração de relatórios concluída. {relatorios_gerados} relatórios gerados.")
                except Exception as e:
                    logger.error(f"Erro na extração de relatórios: {str(e)}")
                    resultados["erros"].append(f"Relatórios: {str(e)}")
        
        # Listar os arquivos gerados após a extração
        resultados["arquivos_gerados"] = listar_arquivos_gerados(save_dir)
    
    except Exception as e:
        logger.error(f"Erro geral na extração: {str(e)}")
        logger.error(traceback.format_exc())
        resultados["erros"].append(f"Erro geral: {str(e)}")
    
    # Imprimir os resultados como JSON para serem capturados pelo frontend
    print(json.dumps({"resultados": resultados}))
    sys.stdout.flush()
    
    return resultados

if __name__ == "__main__":
    # Verificar se foram passados argumentos suficientes
    if len(sys.argv) < 5:
        print(json.dumps({"erro": "Argumentos insuficientes. Uso: python extractus_bridge.py <excel_path> <save_dir> <due_day> <month_number> <tipos>"}))
        sys.exit(1)
    
    # Obter argumentos da linha de comando
    excel_path = sys.argv[1]
    save_dir = sys.argv[2]
    due_day = sys.argv[3]
    month_number = sys.argv[4]
    tipos = sys.argv[5].split(',') if len(sys.argv) > 5 else ["boleto", "nota", "relatorio"]
    
    # Executar a extração
    resultados = executar_extracao(excel_path, save_dir, due_day, month_number, tipos)
    
    # Imprimir os arquivos gerados para facilitar o download no frontend
    print(json.dumps({"arquivos_gerados": resultados.get("arquivos_gerados", [])}))
    sys.stdout.flush()