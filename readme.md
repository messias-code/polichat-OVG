# 🤖 Motor de Extração e Engenharia de Dados - Poli Digital (OVG)

## 1. Propósito do Script

Este projeto é uma **Pipeline de Dados completa (ETL)** desenhada para automatizar a extração e o tratamento dos relatórios de atendimento do Metabase (Poli Digital).

O script atua em duas fases automatizadas e silenciosas (em background):

1. **Fase de Extração (Selenium):** Simula a navegação humana, faz o login no sistema, aplica o filtro de "1 Ano" de histórico e faz o download do relatório em formato CSV.
2. **Fase de Tratamento (Pandas):** Lê o CSV bruto, limpa o "lixo" do sistema, aplica cálculos cronométricos precisos e exporta uma base de dados em Excel (`.xlsx`) Premium, estruturada de forma cronológica e pronta para ser consumida por ferramentas de Business Intelligence, como o Power BI.

---

## 2. Como Rodar o Script

**Pré-requisitos:**

* Ter o Python instalado na máquina.
* Ter o Google Chrome instalado.

**Passo 1: Instalar as bibliotecas necessárias**
Abra o terminal e execute o comando abaixo para instalar as dependências de navegação e tratamento de dados:

```bash
pip install pandas numpy selenium webdriver-manager xlsxwriter

```

**Passo 2: Executar o robô**
Navegue até à pasta onde o script `app.py` se encontra e rode:

```bash
python app.py

```

*O script rodará de forma 100% invisível. Basta acompanhar os logs no terminal. Ao final, o ficheiro `relatorio_chats_pronto.xlsx` estará disponível na pasta `downloads`.*

---

## 3. Explicação dos Dados Tratados e os Motivos

O relatório bruto exportado pelo sistema contém diversas anomalias que corrompem a leitura em ferramentas analíticas. O Python aplica as seguintes "vacinas" aos dados:

* **Quebra da Notação Científica:** O Excel transforma números maiores que 11 dígitos (como IDs e Telefones) em formatos matemáticos (ex: `5.56E+12`). O script converte estas colunas em Texto Puro.
* **Remoção de Alerta de Erro Visual:** Para evitar que o Excel encha a folha com "triângulos verdes" a reclamar dos números guardados como texto, o script desativa esse erro programaticamente em toda a folha de cálculo.
* **Padronização de Fusos Horários (ISO 8601):** O sistema devolve datas no formato `2026-02-01T00:11:00-03:00`. O script desmembra este formato, separando Datas e Horas isoladas para garantir que os **filtros de calendário nativos do Excel e do Power BI** consigam agrupar os dados por Anos, Meses e Dias.
* **Morte das "Mensagens do Cliente":** Como o sistema falha na contagem exata de mensagens trocadas, todas as colunas de quantitativo de mensagens foram removidas. A inteligência da base agora foca-se em **Cronómetros Irrefutáveis** (ex: Subtração da Data de Fechamento pela Data da 1ª Resposta).

---

## 4. Dicionário de Dados: Cenários, Colunas e Insights

A base final foi construída para contar uma **História Cronológica**. Ao ler a tabela da esquerda para a direita, o gestor acompanha o ciclo de vida exato do cliente.

### 📍 ETAPA 1: A Chegada (O Contexto)

* **Id do atendimento / Cliente / Telefone:** Identificação única do ticket e do utilizador.
* **1. Data / Hora de Entrada:** O exato segundo em que o cliente disparou o primeiro "Olá".
* **Período do Dia:** Categoriza o momento do contacto para mapas de calor:
* `Madrugada` (00:00 - 05:59)
* `Manhã` (06:00 - 11:59)
* `Tarde` (12:00 - 17:59)
* `Noite` (18:00 - 23:59)


* **Dentro do Expediente?:** Retorna `SIM` apenas se o contacto ocorreu de Segunda a Sexta, entre as **08:01 e 17:59**. Permite ao gestor isolar facilmente nos gráficos quem "entope" a fila com mensagens fora de horas.

### 📍 ETAPA 2: A Fila (Distribuição e SLA)

* **Houve redirecionamento / Departamento / Atendente:** Onde o chat caiu e quem foi designado para o resolver.
* **Tempo de Espera (Fila):** Cronómetro em `HH:MM:SS` do tempo que o cliente esperou até receber um "Olá" humano.
* **Avaliação da Espera:** O "Semáforo" de SLA do Power BI. Classifica a dor do cliente:
* `🟢 Rápido (< 5 min)` / `🟡 Aceitável (5 a 15 min)` / `🟠 Demorado (> 15 min)`.
* `⚠️ Passou para o Dia Seguinte`: O cliente entrou às 17h50, não foi respondido e a primeira resposta só aconteceu no dia seguinte às 08h.
* `👻 Vácuo Total (Fechado após X min)`: O chat caiu para o Atendente, ele **NUNCA** respondeu e o chat foi simplesmente encerrado.
* `🤖 Sistema/Robô`: O cliente ficou retido no menu eletrónico e o chat foi morto pela URA sem sequer chegar a um humano.



### 📍 ETAPA 3: O Atendimento (A Conversa)

* **2. Data / Hora da 1ª Resposta:** O exato segundo em que o Atendente deu o "Olá".
* **Diagnóstico da Conversa:** Lê a qualidade da interação baseada no tempo, em vez de mensagens:
* `✅ Atendimento com Interação`: Ocorreu uma resposta e a conversa fluiu por mais de 1 minuto.
* `⚡ Fechamento Imediato`: O atendente assumiu o chat e fechou-o em menos de 60 segundos (Forte indicador de encerramento em massa ou tentativa de burlar metas).
* `👻 Ignorado`: O Atendente assumiu, mas a data da primeira resposta é inexistente.
* `🤖 Retido no Robô`: Chat sem dono humano.


* **Tempo de Conversa (Atendimento):** O tempo cronometrado exclusivo da interação. Começa na 1ª resposta do atendente e termina no encerramento (ignora o tempo de fila).

### 📍 ETAPA 4: O Encerramento (A Morte do Ticket)

* **3. Data / Hora de Encerramento:** O segundo exato em que o botão "Encerrar" foi clicado.
* **Fechado por:** Informação letal para descobrir auditorias. Se o Atendente for "Fabricio" e o Fechado por for "Administração", o Fabricio abandonou o chat e o supervisor teve de fechá-lo à força horas/dias depois.
* **Tempo Total (Início ao Fim):** Duração absoluta do ciclo de vida do ticket (Fila + Conversa).
* **Status Final:** `Encerrado` ou `Em Aberto` (se o script não encontrar data de fim).
* **Motivo do serviço / Fechamento:** As Tags de tabulação usadas para os relatórios de volume de cada setor (ex: Quantos foram por questões de Matrícula, Financeiro, etc.).