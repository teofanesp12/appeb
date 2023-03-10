# Applicativos do EB
**Alterações e Assentamentos de Militares**

## LISTA CSV

### VARIAVEIS

Arquivo Tipo Planilha simples, onde deve conter o seguinte obrigatoriamente:

  Variavel |    Obs      
---------- | ---------------------------------------------- 
nome       | Nome do Militar
identidade | Identidade do Militar
grad       | Graduação do Militar: Sd Ef Vrv
qm         | Qualificação Militar: RECRUTA, INFANTARIA, etc

### VARIAVEIS DO MODO TAF

`modo = taf`

onde *{nome_sessao}* representa a sessão configurada no arquivo *configure.ini*

  Variavel              |    Obs      
----------------------- | ---------------------------------------------- 
{nome_sessao}_MEN       | Menção do TAF: E, MB, B, R, I
{nome_sessao}_PAD       | PAD do TAF: S ou I
{nome_sessao}_PBD       | PBD do TAF
{nome_sessao}_OBS       | Obs do TAF realizado

### VARIAVEIS DO MODO INSPEÇÃO DE SAÚDE

`modo = insp_saude`

onde *{nome_sessao}* representa a sessão configurada no arquivo *configure.ini*

  Variavel              |    Obs      
----------------------- | ---------------------------------------------- 
{nome_sessao}_SESSAO    | SESSÃO que foi Salvo o documento
{nome_sessao}_DATA      | Data da inspeção de saúde
{nome_sessao}_ATA       | Ata do documento da inspeção de saúde

### VARIAVEIS DO MODO INCONSISTÊNCIA BANCÁRIA

`modo = inconsistencia_bancaria`

onde *{nome_sessao}* representa a sessão configurada no arquivo *configure.ini*

  Variavel              |    Obs      
----------------------- | ---------------------------------------------- 
{nome_sessao}_CPF       | CPF do militar
{nome_sessao}_BANCO     | Banco
{nome_sessao}_AGENCIA   | Agência
{nome_sessao}_CONTA     | Conta
{nome_sessao}_CREDITADO | Valor creditado

### VARIAVEIS DO MODO OLIMPIADAS

`modo = olimpiadas`

onde *{nome_sessao}* representa a sessão configurada no arquivo *configure.ini*

  Variavel                |    Obs      
------------------------- | ---------------------------------------------- 
{nome_sessao}_MODALIDADE  | Modalidade
{nome_sessao}_PONTOS      | Pontos
{nome_sessao}_TEMPO       | Tempo
{nome_sessao}_CLASS       | Classificação
{nome_sessao}_SU          | Subunidade

## CONFIGURE INI

este arquivo é importante para a configuração das Alterações...

### SESSÃO

as sessões são definidos *[nome_sessao]*, onde _nome_sessao_ é um nome simbolico e sem espaço, utilizado para separar os eventos.
os nomes das sessões são pesquisadas no arquivo *lista.csv* se forem = 1 ou X ele apresentará a sessão no documento final.

### MODOS

dentro das sessoes, exite uma variavel chamada *modo* onde definimos o modo daquele arquivo, Ex: normal, taf, .

### VARIAVEIS

são as seguintes:

  Variavel |    Obs      
---------- | --------------------------------------------- 
data       | Data que aconteceu o evento 
titulo     | Titulo do Evento 
tipo       | Tipo do Evento, Ex: Distribuição etc
doc_tipo   | Tipo do Documento: BI, BAR, BAEs
doc_numero | Número do Documento
modo       | Modo de arrumação do documemento: normal, taf, insp_saude
assunto    | Toda as descrição do evento.

### VARIAVEIS ALTERNATIVAS

Podemos chamar atraves da variavel *assunto* algum valor do *lista.csv*, assim:

```
assunto =  O referido militar, concluiu com aproveitamento em 15 de julho de 2022 a Fase de Instrução Individual de Qualificação: {qmp_qms}

```

no exemplo acima, vemos que *{qmp_qms}* faz referencia a um valor que se encontra no arquivo *lista.csv*

## EXECUTAR

para isso existe duas maneiras:

* A mais simples: clicar duas vezes no arquivo _app.py_
* Executando o arquivo _documentos_assentamentos.py_

![ScreenShot](https://github.com/teofanesp12/appeb/blob/main/doc/assentamentos/gui_run.png?raw=true)
