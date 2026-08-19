# Automação de PCP com Inteligência Artificial Multimodal e Python

Solução de automação desenvolvida para o setor de Planejamento e Controle de Produção (PCP), permitindo transformar relatórios falados em dados estruturados e inseri-los automaticamente em planilhas do Excel em tempo real.

## Arquitetura da Solução

* **Captura de Áudio:** Gravação dinâmica via microfone através das bibliotecas `sounddevice` e `soundfile`, exportando o sinal para formato wave.
* **Processamento LLM:** Envio do áudio diretamente para a API Multimodal do Gemini. Utilizando engenharia de prompt, o modelo extrai os dados brutos da fala e os sanitiza em um JSON padronizado com campos de data, modelo, medidas e grade de tamanhos.
* **Automação de Formulário:** Uso da biblioteca `PyAutoGUI` para realizar a navegação e a digitação automática na interface do Excel, preenchendo cada célula e confirmando a entrada de dados.
* **Deploy:** Empacotamento de toda a aplicação usando `PyInstaller` para gerar um executável independente.

---

## Desafio Técnico: Resolução do Endpoint da API

Durante a integração, surgiram erros de resposta da API (códigos 404 Not Found e API Key Invalid) ao tentar realizar chamadas usando a SDK padrão.

**Resolução e Diagnóstico:**
1. **Incompatibilidade da SDK:** Identificação de que as chaves mais recentes exigiam uma estrutura de rota diferente da tratada pelas chamadas legadas da biblioteca.
2. **Análise por cURL:** Extração da requisição nativa no formato cURL a partir do console. Identificou-se que o cabeçalho de autenticação exigia a chave enviada especificamente via `X-goog-api-key`, e o modelo ativo era o `gemini-flash-latest`.
3. **Validação Direta:** Construção de um script leve utilizando a biblioteca `urllib` do Python para disparar requisições POST diretamente para a URL do endpoint, isolando a camada de biblioteca de terceiros.
4. **Refatoração:** Conversão do áudio gravado para representação em Base64 e reestruturação da aplicação principal para consumir a API via HTTP direto.

---

## Como Executar

1. Instale as dependências:
```bash
pip install -r requirements.txt
