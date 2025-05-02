# Extractus Web Integration

Este projeto implementa uma integração web para o sistema Extractus, permitindo a extração de documentos (boletos, notas fiscais e relatórios) a partir de uma planilha Excel.

## Estrutura do Projeto

O sistema consiste em três componentes principais:

1. **Frontend (React)**: Interface para seleção de parâmetros e visualização do progresso da extração
2. **API Backend (Node.js)**: Ponte entre o frontend e os scripts Python
3. **Scripts Python**: Lógica de extração de documentos

## Instalação e Configuração

### Requisitos

- Node.js (v14 ou superior)
- Python (v3.8 ou superior)
- Playwright para Python (para os scripts de extração)

### Backend (Node.js)

1. Instale as dependências do Node.js:

```bash
npm install
```

2. Inicie o servidor backend:

```bash
npm start
```

O servidor estará disponível em http://localhost:3000.

### Frontend

O frontend está implementado como páginas HTML estáticas que podem ser servidas por qualquer servidor web. Para desenvolvimento, você pode usar o servidor HTTP simples do Python:

```bash
python -m http.server 8000
```

Acesse o frontend em:
- http://localhost:8000/collaborator.html (Interface completa)
- http://localhost:8000/extractus_demo.html (Demo simplificada)

## Como Funciona

### Fluxo de Extração

1. O usuário seleciona um arquivo Excel, data de vencimento, mês e tipos de documentos a serem extraídos
2. O frontend envia esses parâmetros para o backend
3. O backend executa o script Python (extractus_bridge.py) com os parâmetros fornecidos
4. O script Python processa a planilha e extrai os documentos solicitados
5. O progresso é enviado de volta ao frontend em tempo real usando Server-Sent Events (SSE)
6. Ao finalizar, os resultados são disponibilizados para download

### Arquivos Principais

- **server.js**: Implementação do servidor Node.js com Express
- **extractus_api.js**: Cliente JavaScript para comunicação com o backend
- **extractus_bridge.py**: Script Python que coordena a extração de documentos
- **boleto_hapvida.py**, **nota_fiscal.py**, **relatorio_hapvida.py**: Scripts Python para extração de cada tipo de documento
- **collaborator.html**: Interface principal com React
- **extractus_demo.html**: Versão simplificada da interface

## Solução de Problemas

### O servidor backend não inicia

Verifique se todas as dependências foram instaladas corretamente:

```bash
npm install
```

### A extração não funciona

1. Verifique se o servidor backend está rodando em http://localhost:3000
2. Verifique se o arquivo Excel está no formato correto
3. Verifique os logs do servidor para identificar possíveis erros

### Erros nos scripts Python

1. Verifique se o Python e todas as dependências estão instalados:

```bash
pip install playwright
python -m playwright install
```

2. Verifique os logs do servidor para identificar erros específicos nos scripts Python

## Desenvolvimento

Para desenvolvimento, você pode iniciar o servidor em modo de desenvolvimento, que reinicia automaticamente quando há alterações:

```bash
npm run dev
```

## Licença

Este projeto é proprietário e confidencial. Todos os direitos reservados.