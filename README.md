# Sistema Integrado de PCP com IA e Automação Python

Solução end-to-end desenvolvida para o setor de Planejamento e Controle de Produção (PCP) têxtil. O ecossistema transforma relatórios falados em dados estruturados, automatiza a leitura e validação de etiquetas de rolos de tecido via DataMatrix/QR Code e disponibiliza um assistente virtual para consultas estratégicas aos dados em tempo real.

---

## 🏛️ Arquitetura e Módulos da Solução

O projeto é dividido em 3 pilares principais interligados à planilha mestra do Excel (`Controle Consumo de Tecidos Duda.xlsm`):

### 1. Lançamento de Formuário por Voz (`app_ia_pcp.py`)
* **Captura de Áudio:** Gravação dinâmica via microfone através das bibliotecas `sounddevice` e `soundfile`, exportando o sinal para formato WAVE e processando matrizes de som em memória com `numpy`.
* **Processamento LLM:** Envio dos bytes de áudio para a API Multimodal do Gemini via SDK oficial `google-genai`. O modelo extrai os dados brutos da fala e os sanitiza em um JSON padronizado com campos de data, modelo, medidas e grade de tamanhos (PP ao 5G e Único).
* **Automação de Interface:** Uso da biblioteca `PyAutoGUI` para navegação e digitação automática nos campos do formulário no Excel.

### 2. Leitor e Auditoria de Rolos (`gravar_excel_direto.py`)
* **Interface Gráfica Desktop:** Desenvolvida em `Tkinter`, permitindo operação ágil pelo cortador.
* **Regex Especializada para DataMatrix/GS1:** Extração e sanitização dos campos de Metragem, Peso Líquido e Número do Rolo (14 dígitos), tratando caracteres invisíveis de controle ASCII (`Group Separator ^] / GS`).
* **Automação Nativa COM (`win32com.client`):** Inserção de dados via motor nativo do Excel no Windows, preservando intactos todos os formulários, botões de ação e macros VBA (`.xlsm`) da planilha sem corromper a estrutura de arquivos.

### 3. Assistente Virtual de Consulta IA (`consulta_ia_pcp.py`)
* **Processamento de Dados com Pandas:** Leitura em tempo real das abas de auditoria e lançamentos da planilha em rede local (`Z:\`).
* **Consulta em Linguagem Natural:** Integração com LLM para responder dúvidas operacionais da gestão (ex: *"Quantos metros do modelo 2904 foram cortados hoje?"* ou *"Qual o peso total dos rolos bipados?"*).

---

## 🛡️ Desafios Técnicos e Resiliência

Durante a homologação da ferramenta, surgiram desafios de infraestrutura, formatação e roteamento de APIs que foram solucionados com as seguintes arquiteturas:

### 1. Migração para a SDK Oficial (`google-genai`)
* **Problema:** Chamadas REST manuais via HTTP (`urllib`) sofriam com alterações frequentes de endpoints e descontinuação de aliases legados (erros 404 e 500).
* **Solução:** Adoção da biblioteca oficial do Google (`google-genai`), garantindo o gerenciamento automático dos cabeçalhos de requisição, suporte nativo ao formato de áudio e roteamento dinâmico para os modelos vigentes em produção.

### 2. Tratamento de Picos de Demanda na Nuvem (Failover e Retry)
* **Problema:** Erros temporários de alta demanda no servidor do Google (código `503 UNAVAILABLE`) interrompiam o fluxo de trabalho do usuário.
* **Solução:** Implementação de uma lógica de resiliência com até 3 tentativas automáticas de reenvio (`time.sleep`) e *fallback* automático para modelos de backup (`gemini-1.5-flash`), evitando falhas por instabilidades momentâneas da nuvem.

### 3. Preservação de Macros e Botões no Excel (`.xlsm`)
* **Problema:** A gravação direta via manipuladores de XML compactados gerava alertas de segurança no Excel ("conteúdo ilegível") e excluía botões de formulário flutuantes associados a macros VBA.
* **Solução:** Substituição do pipeline de escrita por automação COM (`win32com.client`), fazendo o Python interagir de forma invisível com a própria aplicação do Excel instalada na máquina, mantendo a integridade visual e funcional dos arquivos.

### 4. Tratamento Global de Exceções em Modo `noconsole`
* **Problema:** Ao compilar com PyInstaller ocultando a janela do prompt (`--noconsole`), exceções de tempo de execução fechavam a aplicação sem feedback ao operador.
* **Solução:** Implementação de um bloco estruturado de captura geral (`try...except`) que utiliza o `traceback` para exibir pop-ups visuais interativos (`pyautogui.alert` / `tkinter.messagebox`) informando o status preciso de qualquer falha de rede ou execução.


