/**
 * Este arquivo implementa a integração com o backend Node.js
 * para executar os scripts Python de extração de documentos.
 */

// URL base do servidor backend
const API_BASE_URL = 'http://localhost:3000/api';

// Função para executar a extração
async function executarExtracao(params) {
  const { excelPath, saveDir, dueDay, monthNumber, tipos } = params;
  
  console.log(`Iniciando extração com os seguintes parâmetros:
    Excel: ${excelPath}
    Diretório: ${saveDir}
    Dia de vencimento: ${dueDay}
    Mês: ${monthNumber}
    Tipos: ${tipos.join(', ')}
  `);
  
  // Criar um FormData para enviar o arquivo e os parâmetros
  const formData = new FormData();
  
  // Se excelPath é um objeto File, usá-lo diretamente
  if (excelPath instanceof File) {
    formData.append('file', excelPath);
  } else {
    // Caso contrário, tentar obter o arquivo do input
    const fileInput = document.getElementById('excel_file');
    if (fileInput && fileInput.files[0]) {
      formData.append('file', fileInput.files[0]);
    } else {
      throw new Error('Arquivo Excel não encontrado');
    }
  }
  
  formData.append('dueDay', dueDay);
  formData.append('monthNumber', monthNumber);
  formData.append('tipos', JSON.stringify(tipos));
  
  try {
    // Enviar a requisição para o backend
    const response = await fetch(`${API_BASE_URL}/extractus`, {
      method: 'POST',
      body: formData
    });
    
    if (!response.ok) {
      throw new Error(`Erro na requisição: ${response.status} ${response.statusText}`);
    }
    
    const data = await response.json();
    return data;
  } catch (error) {
    console.error('Erro ao iniciar extração:', error);
    throw error;
  }
}

// Função para criar um EventSource para acompanhar o progresso
function criarEventSource(tipo) {
  // Armazenar o ID da extração atual
  let extractionId = null;
  
  // Criar um objeto que simula um EventSource
  const eventSourceWrapper = {
    onmessage: null,
    close: function() {
      if (this._eventSource) {
        this._eventSource.close();
      }
    },
    _eventSource: null,
    _interval: null,
    
    // Método para iniciar o EventSource real
    connect: function(id) {
      extractionId = id;
      
      if (this._eventSource) {
        this._eventSource.close();
      }
      
      this._eventSource = new EventSource(`${API_BASE_URL}/extractus/progress?extractionId=${extractionId}&tipo=${tipo}`);
      
      this._eventSource.onmessage = (event) => {
        if (this.onmessage) {
          this.onmessage(event);
        }
      };
      
      this._eventSource.onerror = (error) => {
        console.error(`Erro no EventSource para ${tipo}:`, error);
        this._eventSource.close();
      };
    },
    
    // Método para simular o progresso (usado apenas para demonstração)
    simulateProgress: function() {
      // Se temos um ID de extração, usar o EventSource real
      if (extractionId) {
        this.connect(extractionId);
        return null;
      }
      
      // Caso contrário, simular o progresso
      console.warn('Simulando progresso. Em um ambiente real, isso usaria Server-Sent Events.');
      
      let progresso = 0;
      this._interval = setInterval(() => {
        progresso += Math.floor(Math.random() * 10) + 5;
        if (progresso >= 100) {
          progresso = 100;
          clearInterval(this._interval);
          
          if (this.onmessage) {
            this.onmessage({
              data: JSON.stringify({
                progresso: { tipo, porcentagem: progresso },
                resultados: { [tipo]: Math.floor(Math.random() * 100) + 50 }
              })
            });
          }
          
          return;
        }
        
        if (this.onmessage) {
          this.onmessage({
            data: JSON.stringify({
              progresso: { tipo, porcentagem: progresso }
            })
          });
        }
      }, 500);
      
      return this._interval;
    }
  };
  
  return eventSourceWrapper;
}

// Exportar as funções para uso no frontend
if (typeof window !== 'undefined') {
  window.executarExtracao = executarExtracao;
  window.criarEventSource = criarEventSource;
}

// Exportar as funções para uso em Node.js
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    executarExtracao,
    criarEventSource
  };
}