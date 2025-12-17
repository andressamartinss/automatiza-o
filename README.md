# automatiza-o
Automação em Python para limpeza, padronização e comparação de planilhas Excel, gerando um relatório de divergências para facilitar a conferência e correção de dados.

## 📊 Automação de Comparação de Planilhas

Este projeto tem como objetivo automatizar a conferência entre duas planilhas Excel — **Base Interna** e **Planilha de Credenciados** — eliminando a necessidade de verificação manual linha por linha.

A automação realiza a **limpeza, padronização e comparação dos dados**, garantindo que pequenas diferenças de formatação não gerem erros na análise.

---

## 🚀 Funcionalidades

* Padronização automática dos dados:

  * Remove espaços extras no início e no fim das células;
  * Ignora diferenças entre letras maiúsculas e minúsculas;
  * Normaliza acentos e caracteres especiais.
* Comparação precisa entre duas planilhas Excel;
* Geração automática de um **relatório de correção**, exibindo apenas os itens com divergência;
* Identificação de:

  * Códigos divergentes;
  * Códigos ausentes;
  * Códigos novos na planilha credenciada.

---

## 🛠️ Tecnologias Utilizadas

* **Python**
* **Pandas**
* **OpenPyXL**
* **Unidecode**

---

## 📋 Pré-requisitos

* Python instalado (versão 3.12 ou superior);
* Bibliotecas Python necessárias:

  ```
  pip install pandas openpyxl unidecode
  ```

---

## ⚙️ Como Executar

1. Acesse a pasta do projeto pelo terminal:

   ```
   cd downloads\comparar
   ```

2. Execute o script de limpeza e padronização:

   ```
   python limpar_planilha_sem_formatacao.py
   ```

3. Execute o script de comparação para gerar o relatório.

---

## 📄 Resultado

Ao final do processo, o sistema gera automaticamente uma **planilha de correção** na pasta do projeto, indicando exatamente quais campos precisam ser ajustados.

Isso torna a conferência:

* Mais rápida;
* Mais organizada;
* Mais confiável.

---

## 🎯 Objetivo

Reduzir erros humanos, ganhar produtividade e garantir consistência na validação de dados entre diferentes bases de informação.

