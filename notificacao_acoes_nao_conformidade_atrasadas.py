#Importando bibliotecas
import win32com.client
import datetime
import time
import sys
from dateutil.relativedelta import relativedelta

hoje = datetime.datetime.now()
ano = hoje.year
mes = hoje.month
dia = hoje.day
hora = hoje.hour
minuto = hoje.minute

xlapp = win32com.client.DispatchEx("Excel.Application")
xlapp.Visible = 0
wb = xlapp.workbooks.open("C:\\Users\\carlos.junior\\Desktop\\Dashboards\\Notificacoes nao conformidades em aberto.lnk")
wb.RefreshAll()
xlapp.CalculateUntilAsyncQueriesDone()
xlapp.DisplayAlerts = False
wb.Save()
wb.Close()
xlapp.Quit()

# #-------------------------------------------------------------------------------------------------
#importando o pandas
import pandas as pd

#primeira tabela: informações dos atrasos

#lendo o arquivo
notificacoes = pd.read_excel(
    "K:/21 - INDICADORES/01 - Indicadores/5_ Indicadores HKM/Notificacoes_RNC.xlsx", sheet_name="acoes atrasadas")

#segunda tabela: lista de e-mails

#lendo a planilha de e-mails
emails = pd.read_excel(
    "K:/21 - INDICADORES/01 - Indicadores/5_ Indicadores HKM/Notificacoes_RNC.xlsx", sheet_name="emails")

#colhendo as informações das colunas em variáveis
emails_responsaveis = emails["Responsável"].values
emails_grau1 = emails["grau1"].values
emails_grau2 = emails["grau2"].values
emails_grau3 = emails["grau3"].values
#emails.tail()

#filtrando os valores com menos de 7 dias de atraso
grau1 = notificacoes[(notificacoes["Dias de Atraso"]>=1) & (notificacoes["Dias de Atraso"]<7)]

#colocando os valores das tabelas em vetores
grau1_Responsavel = grau1["Responsável pela Ação"].values
grau1_Atraso = grau1["Dias de Atraso"].values
grau1_CodNC = grau1["Cód. Não Conf."].values
grau1_NC = grau1["Não Conformidade"].values
grau1_CodOS = grau1["OS"].values
grau1_Descricao = grau1["Descrição"].values
grau1_Cliente = grau1["Cliente"].values
grau1_Familia = grau1["Família"].values
grau1_Tipo = grau1["Tipo"].values
grau1_Fase = grau1["Fase"].values
grau1_Destino = grau1["Destino"].values
grau1_Acao = grau1["Ação"].values
grau1_mensagens = []
emailparaenvio = []
x = len(grau1_CodNC)

if x != 0:
    #para o intervalo do número de itens da matriz
    for i in range(x):
        #filtrar a tabela notificações baseado no e-mail de cada responsável
        notificacoes_filtro = emails[(emails["Responsável"] == grau1_Responsavel[i])]
        #adicionar este e-mail à matriz de emails para envio
        emailparaenvio.append(notificacoes_filtro["grau1"].values)
        #tratando o texto do e-mail para uma string "limpa"
        emailparaenvio[i] = str(emailparaenvio[i]).replace("[","")
        emailparaenvio[i] = str(emailparaenvio[i]).replace("]","")
        emailparaenvio[i] = str(emailparaenvio[i]).replace("'","")
        #adicionar a mensagem para a matriz de mensagens de grau 1
        grau1_mensagens.append(f'''
        Bom dia,

        Na data de hoje, foi constado no sistema de notificações de não-conformidade um atraso nas ações que deveriam ser tomadas.

        Informações do atraso:
            Responsável:{grau1_Responsavel[i]}
            Dias de atraso: {grau1_Atraso[i]}
            Código de não conformidade:{grau1_CodNC[i]}
            Título da NC: {grau1_NC[i]}
            Código da OS:{grau1_CodOS[i]}
            Descrição: {grau1_Descricao[i]}
            Cliente: {grau1_Cliente[i]}
            Família: {grau1_Familia[i]}
            Tipo: {grau1_Tipo[i]}
            Fase: {grau1_Fase[i]}
            Destino: {grau1_Destino[i]}
            Ação: {grau1_Acao[i]}
        
        É possível acessar as informações completas do registro da não conformidade via Software GRV. Para isso, acesse o menu Cadastro -> Qualidade -> Não Conformidade.
            
        Em caso de dúvidas, divergências ou sugestões, favor entrar em contato.
        Este é um e-mail automático, porém sinta-se livre para respondê-lo.

        Att,
        Carlos Alberto Barbosa Junior''')

for email in emailparaenvio:
    if email == '':
        print('Há funcionários de grau 1 sem e-mail cadastrado')
        sys.exit()
        
print(emailparaenvio)

#filtrando os valores entre 7 e 14 dias de atraso
grau2 = notificacoes[(notificacoes["Dias de Atraso"]>=7) & (notificacoes["Dias de Atraso"]<=14)]

#colocando os valores das tabelas em vetores
grau2_Responsavel = grau2["Responsável pela Ação"].values
grau2_Atraso = grau2["Dias de Atraso"].values
grau2_CodNC = grau2["Cód. Não Conf."].values
grau2_NC = grau2["Não Conformidade"].values
grau2_CodOS = grau2["OS"].values
grau2_Descricao = grau2["Descrição"].values
grau2_Cliente = grau2["Cliente"].values
grau2_Familia = grau2["Família"].values
grau2_Tipo = grau2["Tipo"].values
grau2_Fase = grau2["Fase"].values
grau2_Destino = grau2["Destino"].values
grau2_Acao = grau2["Ação"].values
grau2_mensagens = []
grau2_emailparaenvio = []
y = len(grau2_CodNC)

if y != 0:
    #para o intervalo do número de itens da matriz
    for j in range(y):
        #filtrar a tabela notificações baseado no e-mail de cada responsável
        notificacoes_filtro2 = emails[(emails["Responsável"] == grau2_Responsavel[j])]
        #adicionar este e-mail à matriz de emails para envio
        grau2_emailparaenvio.append(notificacoes_filtro2["grau2"].values)
        #tratando o texto do e-mail para uma string "limpa"
        grau2_emailparaenvio[j] = str(grau2_emailparaenvio[j]).replace("[","")
        grau2_emailparaenvio[j] = str(grau2_emailparaenvio[j]).replace("]","")
        grau2_emailparaenvio[j] = str(grau2_emailparaenvio[j]).replace("'","")
        #adicionar a mensagem para a matriz de mensagens de grau 2
        grau2_mensagens.append(f'''
        Bom dia,

        Na data de hoje, foi constado no sistema de notificações de não-conformidade um atraso nas ações que deveriam ser tomadas.

        Informações do atraso:
            Responsável:{grau2_Responsavel[j]}
            Dias de atraso: {grau2_Atraso[j]}
            Código de não conformidade:{grau2_CodNC[j]}
            Título da NC: {grau2_NC[j]}
            Código da OS:{grau2_CodOS[j]}
            Descrição: {grau2_Descricao[j]}
            Cliente: {grau2_Cliente[j]}
            Família: {grau2_Familia[j]}
            Tipo: {grau2_Tipo[j]}
            Fase: {grau2_Fase[j]}
            Destino: {grau2_Destino[j]}
            Ação: {grau2_Acao[j]}
        
        É possível acessar as informações completas do registro da não conformidade via Software GRV. Para isso, acesse o menu Cadastro -> Qualidade -> Não Conformidade.
            
        Em caso de dúvidas, divergências ou sugestões, favor entrar em contato.
        Este é um e-mail automático, porém sinta-se livre para respondê-lo.

        Att,
        Carlos Alberto Barbosa Junior''')
        
for email2 in grau2_emailparaenvio:
    if email2 == '':
        print('Há funcionários de grau 2 sem e-mail cadastrado')
        sys.exit()
        
print(grau2_emailparaenvio)

#filtrando os valores acima de 14 dias de atraso
grau3 = notificacoes[notificacoes["Dias de Atraso"]>14]

#colocando os valores das tabelas em vetores
grau3_Responsavel = grau3["Responsável pela Ação"].values
grau3_Atraso = grau3["Dias de Atraso"].values
grau3_CodNC = grau3["Cód. Não Conf."].values
grau3_NC = grau3["Não Conformidade"].values
grau3_CodOS = grau3["OS"].values
grau3_Descricao = grau3["Descrição"].values
grau3_Cliente = grau3["Cliente"].values
grau3_Familia = grau3["Família"].values
grau3_Tipo = grau3["Tipo"].values
grau3_Fase = grau3["Fase"].values
grau3_Destino = grau3["Destino"].values
grau3_Acao = grau3["Ação"].values
grau3_mensagens = []
grau3_emailparaenvio = []
z = len(grau3_CodNC)

if z != 0:
    #para o intervalo do número de itens da matriz
    for k in range(z):
        #filtrar a tabela notificações baseado no e-mail de cada responsável
        notificacoes_filtro3 = emails[(emails["Responsável"] == grau3_Responsavel[k])]
        #adicionar este e-mail à matriz de emails para envio
        grau3_emailparaenvio.append(notificacoes_filtro3["grau3"].values)
        #tratando o texto do e-mail para uma string "limpa"
        grau3_emailparaenvio[k] = str(grau3_emailparaenvio[k]).replace("[","")
        grau3_emailparaenvio[k] = str(grau3_emailparaenvio[k]).replace("]","")
        grau3_emailparaenvio[k] = str(grau3_emailparaenvio[k]).replace("'","")
        #adicionar a mensagem para a matriz de mensagens de grau 3
        grau3_mensagens.append(f'''
        Bom dia,

        Na data de hoje, foi constado no sistema de notificações de não-conformidade um atraso nas ações que deveriam ser tomadas.

        Informações do atraso:
            Responsável:{grau3_Responsavel[k]}
            Dias de atraso: {grau3_Atraso[k]}
            Código de não conformidade:{grau3_CodNC[k]}
            Título da NC: {grau3_NC[k]}
            Código da OS:{grau3_CodOS[k]}
            Descrição: {grau3_Descricao[k]}
            Cliente: {grau3_Cliente[k]}
            Família: {grau3_Familia[k]}
            Tipo: {grau3_Tipo[k]}
            Fase: {grau3_Fase[k]}
            Destino: {grau3_Destino[k]}
            Ação: {grau3_Acao[k]}
        
        É possível acessar as informações completas do registro da não conformidade via Software GRV. Para isso, acesse o menu Cadastro -> Qualidade -> Não Conformidade.
            
        Em caso de dúvidas, divergências ou sugestões, favor entrar em contato.
        Este é um e-mail automático, porém sinta-se livre para respondê-lo.

        Att,
        Carlos Alberto Barbosa Junior''')
        
for email3 in grau3_emailparaenvio:
    if email3 == '':
        print('Há funcionários de grau 3 sem e-mail cadastrado')
        sys.exit()
        
print(grau3_emailparaenvio)

# #--------------------------------------------------------------------------------------------------------
# enviando os e-mails

# entrando no e-mail
outlook = win32com.client.Dispatch("Outlook.Application")


#laço de repetição para escrever todos os e-mails
if x != 0:
    for i in range (x):
        #abrindo um e-mail novo
        Msg = outlook.CreateItem(0)
        #escrevendo o destinatário
        Msg.To = str(emailparaenvio[i])
        #acessando o campo assunto e escrevendo-o
        Msg.Subject = "Notificação - Grau 1 - Atraso na ação de não-conformidade"
        #acessando o corpo do e-mail e escrevendo-o
        Msg.Body = str(grau1_mensagens[i])
        #enviando o e-mail
        Msg.Send()

if y != 0:
    for j in range (y):
        #abrindo um e-mail novo
        Msg = outlook.CreateItem(0)
        #escrevendo o destinatário
        Msg.To = str(grau2_emailparaenvio[j])
        #acessando o campo assunto e escrevendo-o
        Msg.Subject = "Notificação - Grau 2 - Atraso na ação de não-conformidade"
        #acessando o corpo do e-mail e escrevendo-o
        Msg.Body = str(grau2_mensagens[j])
        #enviando o e-mail
        Msg.Send()
        
if z != 0:
    for k in range (z):
        #abrindo um e-mail novo
        Msg = outlook.CreateItem(0)
        #escrevendo o destinatário
        Msg.To = str(grau3_emailparaenvio[k])
        #acessando o campo assunto e escrevendo-o
        Msg.Subject = "Notificação - Grau 3 - Atraso na ação de não-conformidade"
        #acessando o corpo do e-mail e escrevendo-o
        Msg.Body = str(grau3_mensagens[k])
        #enviando o e-mail
        Msg.Send()        

