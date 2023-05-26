#331.666.228-61
#331.667.228-62
#413.434.388-70
#287.703.988-95

"""
pip install selenium
pip install webdriver-manager
pip install reportlab
pip install pandas
pip install openpyxl
"""

# Importa as bibliotecas
import time
import os
import sqlite3
import datetime
import pandas as pd
import logging

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Paragraph
from reportlab.lib import colors


# Funções
def enviar_pdf():

    # captura todas as tabelas na página
    try:
        tabelas = driver2.find_elements(By.TAG_NAME, 'table')
        time.sleep(1)
    except:
        logging.info('Erro ao capturar as tabelas')
        driver2.close()
        raise ValueError("error")


    # Cria o pdf
    doc = SimpleDocTemplate("dados.pdf", pagesize=A4)
    elements = []

    # Título da Página
    try:
        dados = [['Pesquisa Cadastral Simplificada']]
        t = Table(dados)
        t.setStyle(TableStyle([('FONTSIZE', (0, 0), (0, 0), 20)]))
        elements.append(t)
        elements.append(Spacer(width=0, height=50))
    except:
        logging.info('Erro ao gerar o título do pdf')
        driver2.close()
        raise ValueError("error")


    # Dados iniciais
    try:
        rows = tabelas[0].find_elements(By.TAG_NAME, "tr")
        cols = rows[0].find_elements(By.TAG_NAME, "td")
        nome = cols[1].text
    except:
        logging.info('Erro ao capturar as linhas e colunas da tabela')
        driver2.close()
        raise ValueError("error")

    try:
        cols = rows[2].find_elements(By.TAG_NAME, "td")
        cpf = cols[1].text
    except:
        logging.info('Erro ao capturar o cpf da tabela')
        driver2.close()
        raise ValueError("error")

    try:
        cols = rows[6].find_elements(By.TAG_NAME, "td")
        data_hora = cols[1].text
    except:
        logging.info('Erro ao capturar a data e hora da tabela')
        driver2.close()
        raise ValueError("error")

    try:
        dados = [['Nome do Cliente:', nome], ['CPF:', cpf], ['Data / Hora:', data_hora]]
        t = Table(dados)
        t.setStyle(TableStyle([('FONTSIZE', (0, 0), (-1, -1), 12)]))
        elements.append(t)
        elements.append(Spacer(width=0, height=40))
    except:
        logging.info('Erro ao gerar o cabeçalho do pdf')
        driver2.close()
        raise ValueError("error")

    # Itera as outras tabelas da segunda até a penúltima
    for i in range(1,len(tabelas) - 1):
        lista_tabela = []
        j = 0
        
        rows = tabelas[i].find_elements(By.TAG_NAME, "tr")
        for row in rows:

            # Transforma a primeira linha da tabela no título
            if j == 0:
                j += 1
                lista_auxiliar = []
                lista_titulo = []
                
                c = row.find_element(By.TAG_NAME, "td")
                try:
                    lista_titulo.append(c.text)
                    lista_auxiliar.append(lista_titulo)
                    t = Table(lista_auxiliar)
                    t.setStyle(TableStyle([('FONTSIZE', (0, 0), (0, 0), 12)]))
                    elements.append(t)
                    elements.append(Spacer(width=0, height=10))
                except:
                    logging.info('Erro ao iterar os títulos das tabelas')
                    pass
                
            # Para o restante da tabela 
            else:
                lista_celula = []    
                cols = row.find_elements(By.TAG_NAME, "td")
                
                for col in cols:
                    try:
                        lista_celula.append(Paragraph(col.text))
                    except:
                        logging.info('Erro ao iterar os dados das tabelas')
                        continue
                
                lista_tabela.append(lista_celula)
            
        try:
            t = Table(lista_tabela)
        
            t.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 1), colors.grey),
                            ('TEXTCOLOR', (0, 0), (-1, 1), colors.whitesmoke),
                            ('ALIGN', (0, 0), (-1, 0), 'LEFT'),
                            ('FONT', (0, 0), (-1, 0), 'Helvetica-Bold', 10),
                            ('FONTSIZE', (0, 1), (-1, -1), 8),
                            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                            ('GRID', (0, 0), (-1, -1), 1, colors.black)]))
            
            elements.append(t)
            elements.append(Spacer(width=0, height=30))
        except:
            logging.info('Erro ao formatar as tabelas')
            continue

    try:
        doc.build(elements)
    except:
        logging.info('Erro ao gerar o pdf')
        driver2.close()
        raise ValueError("error")

    # Fecha a segunda janela
    driver2.close()
    

    # Captura o caminho da pasta
    try:
        diretorio_atual = os.getcwd()
        arquivo = diretorio_atual + '/dados.pdf'
    except:
        logging.info('Erro ao pegar o caminho do arquivo')
        raise ValueError("error")


    # Envia o arquivo
    try:
        driver.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/div').click()
        time.sleep(2)
    except:
        logging.info('Erro ao clicar em anexo')
        raise ValueError("error")
    
    try:
        documento = driver.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/span/div/div/ul/li[4]/button/input')
        documento.send_keys(arquivo)
        time.sleep(10)
    except:
        logging.info('Erro ao carregar o arquivo')
        raise ValueError("error")

    try:
        driver.find_element(By.XPATH, '//*[@id="app"]/div/div/div[3]/div[2]/span/div/span/div/div/div[2]/div/div[2]/div[2]/div/div').click()
        time.sleep(10)
    except:
        logging.info('Erro ao enviar o arquivo')
        raise ValueError("error")


    # Remove o arquivo
    try:
        os.remove(arquivo)
    except:
        logging.info('Erro ao apagar o arquivo')
        pass
    

# Cria o log
logging.basicConfig(level=logging.INFO, filename='app.log', format='%(asctime)s - %(levelname)s - %(message)s')
logging.info('Inicio do programa')


# Abre a planilha com os dados de login
try:
    df = pd.read_excel('login.xlsx') 
except:
    logging.info('Erro ao abrir o arquivo login.xlsx')


# Cria a conexão com o banco de dados
try:
    conn = sqlite3.connect('banco.db')
    cursor = conn.cursor()
except:
    logging.info('Erro ao conectar ao banco de dados')

# Faz um query no banco
try:
    tabelas = cursor.execute("""SELECT * FROM sqlite_master WHERE type='table' AND name='registros';""").fetchall()
except:
    logging.info('Erro ao fazer a query no banco')


# Cria a tabela se ela não existe
if tabelas == []:
    try:
        tabela = """CREATE TABLE registros (
            cpf VARCHAR(20) NOT NULL,
            hora VARCHAR(20) NOT NULL,
            data VARCHAR(20) NOT NULL
            );"""

        cursor.execute(tabela)
    except:
        logging.info('Erro ao criar a tabela')

# Se a tabela existe verifica se tem registro para apagar
else:
    try:
        data_atual = datetime.date.today()
        nova_data = data_atual - datetime.timedelta(days=2)
    except:
        logging.info('Erro ao pegar a data')

    try:
        dados = cursor.execute(f"SELECT * FROM registros;").fetchall()
    except:
        logging.info('Erro ao fazer a query na tabela')

    if dados != []:

        # Itera os registros
        for dado in dados:

            try:
                data = datetime.datetime.strptime(dado[2], "%Y-%m-%d").date()
            except:
                logging.info('Erro ao pegar a data do registro')
            
            # Apaga o registro depois de dois dias
            if data <= nova_data:
                
                try:
                    cursor.execute(f"DELETE FROM registros WHERE data='{dado[2]}'")
                    conn.commit()
                except:
                    logging.info('Erro ao apagar o registro')

# Início do programa
whatsapp = 'https://web.whatsapp.com/'
caminho_pasta_atual = os.getcwd()

options = Options()
options.add_argument("--profile-directory=Default")
options.add_argument(f"--user-data-dir={caminho_pasta_atual}/cookies")


driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
driver.maximize_window()

driver.get(whatsapp)
time.sleep(40)


while True:
    
    # Pega as últimas conversas
    try:
        conversas = driver.find_elements(By.CLASS_NAME, 'aprpv14t')     
    except:
        logging.info('Erro ao pegar as conversas')
        time.sleep(10)
        continue

    i = 0

    # Itera as conversas encontradas
    for conversa in conversas:

        try:  
            conversa.text
        except:
            logging.info('Erro na conversa')
            continue 


        # Verifica se as conversas tem menos de 24h
        if len(conversa.text) == 5 and conversa.text[0].isnumeric():

            # Clica no conversa
            try:
                conversa.click()
            except:
                logging.info('Erro ao clicar na conversa')
                time.sleep(10)
                continue

            # Captura o usuario da conversa
            try:
                usuario = driver.find_elements(By.CLASS_NAME, '_21nHd')
                nome = usuario[0].find_element(By.TAG_NAME, 'span').text 
            except:
                logging.info('Erro ao pegar o nome do usuário')
                time.sleep(10)
                continue
            
            
            # Captura as últimas mensagens
            try:
                mensagens = driver.find_elements(By.CLASS_NAME, 'ItfyB')
            except:
                logging.info('Erro ao pegar as mensagens')
                time.sleep(10)
                continue
            
            # Itera as mensagens encontradas
            for mensagem in mensagens:

                # variáveis
                resultado = ''
                numero = 'a'

                # Captura alguns dados da mensagem
                try:
                    mensagem_numero_span = mensagem.find_elements(By.CLASS_NAME, '_11JPr')
                except:
                    logging.info('Erro ao pegar dados da mensagen')
                    continue


                 # Captura o horário da mensagem
                try:
                    mensagem_horario = mensagem.find_elements(By.CLASS_NAME, 'l7jjieqr')
                    mensagem_horario_correto = mensagem_horario[0].text
                except:
                    logging.info('Erro ao pegar dados do horário')
                    continue
                    
    
                if len(mensagem_numero_span) == 1:
                    
                    # Trata o número quando não for link
                    try:
                        mensagem_numero_span_span = mensagem_numero_span[0].find_elements(By.TAG_NAME, 'span')
                        mensagem_numero_span_texto = mensagem_numero_span_span[0]
                        numero = mensagem_numero_span_texto.text.replace('.', '').replace('-', '').replace('/', '').replace(' ', '') 
                    except:
                        logging.info('Erro ao tratar o número quando não for link')
                        continue
    
                elif len(mensagem_numero_span) == 2:

                    # Trata o número quando for link
                    try:
                        mensagem_numero_span_texto = mensagem_numero_span[1]
                        numero = mensagem_numero_span_texto.text.replace('.', '').replace('-', '').replace('/', '').replace(' ', '')
                    except:
                        logging.info('Erro ao tratar o número quando for link')
                        continue

                else:
                    continue

                
                # Verifica se o número é um CNPJ
                if len(numero) == 14 and numero.isnumeric():

                    resultado = 'Foi digitado um CNPJ, Por Favor digite um CPF!'

            
                # Verifica se o número é um CPF
                elif len(numero) == 11 and numero.isnumeric():
                    
                    # Faz a consulta no banco
                    try:
                        existe = cursor.execute(f"SELECT cpf FROM registros WHERE cpf='{numero}' AND hora='{mensagem_horario_correto}';").fetchall()
                    except:
                        logging.info('Erro ao verificar se existe o cpf no banco')
                        continue

                    # Verifica se a resposta é vazia
                    if existe == []:

                        # Verifica se o nome do contato está salvo
                        if nome[-4:-1].isnumeric():
                            resultado = 'Identifiquei que você não tem autorização para acesso a consulta. Solicite acesso ao administrador!'

                            try:
                                driver.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p').send_keys(resultado)
                                time.sleep(1)
                                driver.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button').click()
                                time.sleep(4)
                            except:
                                logging.info('Erro ao enviar mensagem de contato não salvo')
                                continue

                            #Salvando no banco
                            try:
                                cursor.execute(f"INSERT INTO registros VALUES ('{numero}', '{mensagem_horario_correto}','{datetime.date.today()}')")
                                conn.commit()
                            except:
                                logging.info('Erro ao salvar o cpf do contato não salvo')
                                break
                        else:

                            # Abre uma segunda janela e acessa o site da Caixa
                            try:
                                caixa = 'https://caixaaqui.caixa.gov.br/caixaaqui/CaixaAquiController/index'

                                driver2 = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
                                driver2.maximize_window()
                                driver2.get(caixa)
                                time.sleep(10)
                            except:
                                logging.info('Erro ao abrir a segunda janela')
                                break

                            # Faz login na Caixa
                            try:
                                driver2.find_element(By.XPATH, '//*[@id="convenio"]').send_keys(df['Valores'][0])
                                time.sleep(1)
                                driver2.find_element(By.XPATH, '//*[@id="login"]').send_keys(df['Valores'][1])
                                time.sleep(1)
                                driver2.find_element(By.XPATH, '//*[@id="password"]').send_keys(df['Valores'][2])
                                time.sleep(1)
                                driver2.find_element(By.XPATH, '//*[@id="btLogin"]/input').click()
                                time.sleep(5) 
                            except:
                                logging.info('Erro ao fazer login na segunda janela')
                                driver2.close()
                                time.sleep(1)
                                break

                            # Navega no site
                            try:
                                driver2.find_element(By.XPATH, '//*[@id="menu-principal"]/tbody/tr[1]/td/a').click()
                                time.sleep(2)

                                driver2.find_element(By.XPATH, '//*[@id="menu-principal"]/tbody/tr[1]/td/a').click()
                                time.sleep(2)

                                driver2.find_element(By.XPATH, '//*[@id="menu-principal"]/tbody/tr[2]/td/a').click()
                                time.sleep(2)
                                
                            except:
                                logging.info('Erro ao navegar no site')
                                driver2.close()
                                time.sleep(1)
                                break

                            # Escreve o CPF e busca no site
                            try:      
                                driver2.find_element(By.XPATH, '//*[@id="cpf"]').send_keys(numero)
                                time.sleep(1)

                                botao = driver2.find_elements(By.CLASS_NAME, 'btn-azul')
                                botao[1].click()
                                time.sleep(10)
                                
                            except:
                                logging.info('Erro ao buscar cpf no site')
                                driver2.close()
                                time.sleep(1)
                                break

                            # Variáveis de controle
                            alert_texto = ''
                            pc = ''

                            # Verifica se tem alert
                            try:
                                # Muda o foco para o alerta
                                alert = Alert(driver2)
                                
                                # Obtem o texto do alerta
                                alert_texto = alert.text
                                
                                # Aceite o alerta
                                alert.accept()
                                time.sleep(1)

                            except:
                                # Verifica se gerou um link
                                try:
                                    pesquisa_cadastral = driver2.find_element(By.ID, 'pesquisa-cadastral')
                                    link_pesquisa_cadastral = pesquisa_cadastral.find_element(By.TAG_NAME, 'a')
                                    pc = link_pesquisa_cadastral.text
                                    
                                except:
                                    try:
                                        cpf_cliente = driver2.find_element(By.ID, 'tdcpf').text
                                        nome_cliente = driver2.find_element(By.ID, 'nome').text
                                        regularidade = driver2.find_element(By.ID, 'regularidade').text
                                        pesquisa_cadastral = driver2.find_element(By.ID, 'pesquisa-cadastral')
                                        mensagem_avaliacao = driver2.find_element(By.ID, 'mensagem-avaliacao-risco').text
                                        pc = pesquisa_cadastral.text
                                        
                                    except:
                                        logging.info('Erro no servidor da caixa')
                                        pc = 'Erro no servidor da Caixa!'

                            # Verifica o texto do alert
                            if alert_texto != '':

                                if alert_texto == 'CPF do cliente inválido.':
                                    resultado = 'CPF inválido!'

                                elif alert_texto == 'For input string: " "':

                                    try:
                                        driver2.get('https://caixaaqui.caixa.gov.br/caixaaqui/CaixaAquiController/consulta_cadastral/consulta_cadastral1')
                                        time.sleep(5)
                                        driver2.find_element(By.XPATH, '//*[@id="dataCpf"]').send_keys(numero)
                                        driver2.find_element(By.XPATH, '//*[@id="spanCPF"]/a').click()
                                        time.sleep(2)
                                    except:
                                        logging.info('Erro ao pegar dados do alert')
                                        driver2.close()
                                        time.sleep(1)
                                        break

                                    try:
                                        enviar_pdf()
                                    except:
                                        logging.info('Erro ao enviar o pdf do alert')
                                        driver2.close()
                                        break

                                    try:
                                        #Salvando no banco
                                        cursor.execute(f"INSERT INTO registros VALUES ('{numero}', '{mensagem_horario_correto}','{datetime.date.today()}')")
                                        conn.commit()
                                    except:
                                        logging.info('Erro ao salvar o cpf do alert no banco')
                                        driver2.close()
                                        break

                                    continue
                                            
                                else:
                                    resultado = 'Erro com o CPF'
                                    

                            elif pc == 'Nada consta':
                                resultado = f"""Nome do Cliente: {nome_cliente}
                                Situação CPF na Receita: {regularidade}
                                Pesquisa Cadastral: {pc}
                                {mensagem_avaliacao}"""

                            elif pc == 'Constam Ocorrências':
                                link_pesquisa_cadastral.click()
                                time.sleep(5)

                                try:
                                    enviar_pdf()
                                except:
                                    logging.info('Erro ao salvar o dado do alert no banco')
                                    driver2.close()
                                    break

                                try:
                                    #Salvando no banco
                                    cursor.execute(f"INSERT INTO registros VALUES ('{numero}', '{mensagem_horario_correto}','{datetime.date.today()}')")
                                    conn.commit()
                                except:
                                    logging.info('Erro ao salvar o cpf do constam ocorrências no banco')
                                    driver2.close()
                                    break

                                continue

                            else:
                                resultado = 'Erro no servidor da Caixa!'

                            driver2.close()

                            try:
                                driver.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p').send_keys(resultado)
                                time.sleep(1)
                                driver.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button').click()
                                time.sleep(4)
                                logging.info('Mensagem enviada')
                            except:
                                logging.info('Erro ao enviar a mensagem')
                                break

                            try:
                                #Salvando no banco
                                cursor.execute(f"INSERT INTO registros VALUES ('{numero}', '{mensagem_horario_correto}','{datetime.date.today()}')")
                                conn.commit()
                            except:
                                logging.info('Erro ao salvar cpf no banco no fim do looping')
                                break

    time.sleep(10)