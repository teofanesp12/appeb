# Applicativos do EB
**Software for Military**

## DESCRIÇÃO

Este sistema foi criado para facilitar diversos trabalhos nas OMs, para isso foi utilizado as seguintes ferramentas:

* [LibreOffice](https://pt-br.libreoffice.org/baixe-ja/libreoffice-novo/) documentação e outros modelos simples.
* [Python ORG](https://www.python.org/downloads/) é uma linguagem de Programação.

## INSTALAÇÃO 

### SISTEMA

LINUX:

    apt install python3 python3-lxml python3-tablib python3-tk

WINDOWS:

necessário baixar o instalador: [Python ORG](https://www.python.org/downloads/), depois é só executar o arquivo *instalar.bat*

Sugerimos seguir o seguinte: [Manual de Instalação](https://github.com/teofanesp12/appeb/blob/main/doc/install/INSTALL%20-%20WINDOWS.md)

### DEPENDÊNCIAS

BIBLIOTECAS:

    tablib, diff_match_patch, lxml, tk

PYTHON:

execute o seguinte comando pelo terminal:

    pip install -r requirements.txt

## EXECUTAR 

### CONFIGURAR

para iniciar basta, configurar os arquivos *lista.csv* e *configure.ini*.

### INICIAR

executar o arquivo _app.py_

    python documento{nome}/app.py
