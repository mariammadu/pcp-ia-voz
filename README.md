# Automação de PCP com Inteligência Artificial Multimodal e Python

Solução de automação desenvolvida para o setor de Planejamento e Controle de Produção (PCP), permitindo transformar relatórios falados em dados estruturados e inseri-los automaticamente em planilhas do Excel em tempo real.

## Arquitetura da Solução

* **Captura de Áudio:** Gravação dinâmica via microfone através das bibliotecas `sounddevice` e `soundfile`, exportando o sinal para formato WAVE e processando matrizes de som em memória com `numpy`.
* **Processamento LLM:** Envio dos bytes de áudio diretamente para a API Multimodal do Gemini utilizando o SDK oficial `google-genai`. O modelo extrai os dados brutos da fala e os sanitiza em um JSON padronizado com campos de data, modelo, medidas e grade de tamanhos.
* **Automação de Formulário:** Uso da biblioteca `PyAutoGUI` para realizar a navegação e a digitação automática na interface do Excel, preenchendo cada célula e confirmando a entrada de dados.
* **Deploy:** Empacotamento de toda a aplicação usando `PyInstaller` para gerar um executável independente (`.exe`).

---

## Desafios Técnicos e Resiliência

Durante o desenvolvimento e homologação da ferramenta, surgiram desafios de infraestrutura e rotas de API que foram solucionados com as seguintes arquiteturas:

### 1. Migração para a SDK Oficial (`google-genai`)
* **Problema:** Chamadas REST manuais via HTTP (`urllib`) sofriam com alterações frequentes de endpoints e descontinuação de aliases legados (erros 404 e 500).
* **Solução:** Adopção da biblioteca oficial do Google (`google-genai`), garantindo o gerenciamento automático dos cabeçalhos de requisição, suporte nativo ao formato de áudio e roteamento dinâmico para os modelos vigentes em produção (como o `gemini-3.6-flash`).

### 2. Tratamento de Picos de Demanda na Nuvem (Failover e Retry)
* **Problema:** Erros temporários de alta demanda no servidor do Google (código `503 UNAVAILABLE`) interrompiam o fluxo de trabalho do usuário.
* **Solução:** Implementação de uma lógica de resiliência com até 3 tentativas automáticas de reenvio (`time.sleep`) e *fallback* automático para modelos de backup (`gemini-1.5-flash`), evitando que a automação falhe por instabilidades momentâneas da nuvem.

### 3. Tratamento Global de Exceções em Modo `noconsole`
* **Problema:** Ao compilar com PyInstaller ocultando a janela do prompt (`--noconsole`), exceções de tempo de execução fechavam a aplicação sem feedback ao operador.
* **Solução:** Implementação de um bloco estruturado de captura geral (`try...except`) que utiliza o `traceback` para exibir pop-ups visuais interativos (`pyautogui.alert`) informando o status preciso de qualquer falha de rede ou execução.

---

## Como Executar

1. Instale as dependências:
```bash
pip install -r requirements.txt
