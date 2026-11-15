# Porsche Sales Intelligence

Dashboard interativa desenvolvida para análise de vendas de veículos Porsche, com foco na identificação de padrões comerciais, preferências regionais e comportamento de compra dos clientes.

## Objetivo do projeto

Transformar dados de vendas em informações estratégicas que facilitem a análise comercial e auxiliem na tomada de decisões.

## 1. Perguntas de negócio escolhidas

As perguntas foram definidas com o objetivo de identificar padrões de consumo e compreender o desempenho comercial dos veículos em diferentes regiões.

**1. Quais são os principais modelos Porsche vendidos por cidade ou estado?**

Essa análise permite identificar as preferências de cada região, compreender a distribuição das vendas e descobrir quais modelos possuem maior procura em determinadas localidades.

**2. Qual método de pagamento foi mais utilizado em determinado período?**

O objetivo é identificar as formas de pagamento predominantes entre os compradores, permitindo compreender o comportamento financeiro das vendas.

**3. Quais são os carros mais populares em cada cidade?**

Essa análise busca identificar os modelos com maior volume de vendas em cada localidade, permitindo reconhecer tendências regionais e oportunidades comerciais.

**4. Qual ano de modelo apresentou maior volume de vendas no período analisado?**

Permite identificar quais anos de fabricação/modelo possuem maior participação nas vendas e observar possíveis tendências de preferência dos consumidores.

## 2. Desenvolvimento com Inteligência Artificial

O ChatGPT e o Claude foi utilizado como ferramenta de apoio na concepção e desenvolvimento da dashboard.

### Prompt inicial

"utilizando o recurso de canvas, renderize uma dashboard em html ao lado

\#filtros 
-modelo da porche
-model year
-city
-pay method

\#perguntas de negocios e kpis
-quais os principais modelos de carros por cidade ou estado
-qual o metodo de pagamento que mais saiu em um periodo
-quero insight de carros populares de vendas com base nos dados e em cada cidade

sobre o ui/ux, se baseie no site oficial da porche brasil:
[https://www.porsche.com/brazil/pt/](https://www.porsche.com/brazil/pt/)

tenha um ar de elegante e refinado, deixe o site visivelmente e esteticamente bonito"

### Evolução do prompt

Durante o desenvolvimento, os requisitos foram refinados para deixar a dashboard mais completa e orientada à análise de negócios.

Foram reforçados os seguintes pontos:

- Organização dos indicadores principais (KPIs).
- Melhor apresentação dos gráficos e informações comerciais.
- Inclusão de análises temporais das vendas.
- Maior destaque para os filtros interativos.
- Aprimoramento da identidade visual inspirada na Porsche Brasil.

O objetivo foi evoluir de uma simples visualização de dados para uma ferramenta de análise comercial.

## 3. Tratamento da base de dados

A fonte de dados utilizada foi uma planilha Excel previamente tratada.

Antes da disponibilização dos dados em uma aplicação pública, é necessário garantir:

- Padronização dos nomes dos modelos e das cidades.
- Consistência dos campos de ano e período.
- Verificação de valores ausentes e registros duplicados.
- Conferência dos métodos de pagamento.
- Remoção de informações pessoais ou confidenciais.

Somente dados comerciais necessários às análises devem ser incorporados ao projeto. A publicação da base deve ocorrer após a validação dessas condições.

## 4. Ferramentas utilizadas

- **ChatGPT:** apoio na elaboração dos prompts, estruturação das análises e desenvolvimento da solução.
- Claude: apoio na elaboração dos prompts e correção de gráficos.
- **HTML:** estrutura da dashboard.
- **GitHub:** armazenamento e versionamento do projeto.
- **GitHub Pages:** publicação e hospedagem da aplicação.

## 5. Dashboard publicada

**Acesso à aplicação:**

[https://davi-rizzo.github.io/Porsches-Sales-Dashboard/](https://davi-rizzo.github.io/Porches-Sales-Dashboard/)

**Repositório do projeto:**

[https://github.com/Davi-Rizzo/Porsches-Sales-Dashboard](https://github.com/Davi-Rizzo/Porches-Sales-Dashboard)

## 6. Considerações finais

O projeto demonstra como a utilização de Inteligência Artificial, análise de dados e visualização interativa pode transformar uma base de vendas em informações relevantes para o negócio.

A proposta é facilitar a identificação de tendências, compreender o comportamento dos consumidores e apoiar decisões comerciais com base nos dados disponíveis.

---

Desenvolvido por **Davi Rizzo**.
