from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import datetime
import time
import glob
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import NoSuchElementException
import pandas as pd
import PyPDF2
import re
import sqlite3
import os
class BancoDados:

    # criar banco

    def CriaTabelaJus():
        conn = sqlite3.connect('Jus.db')
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS Registro (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                dataHoraEnvio TEXT,
                dataHoraProcess TEXT,
                numProcess TEXT,
                pagina INTEGER,
                caderno TEXT
            )
        ''')
        conn.commit()

    def Cadastrar(dataHoraEnvio, dataHoraProcess, numProcess, pagina, caderno):
        conn = sqlite3.connect('Jus.db')
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO Registro (dataHoraEnvio, dataHoraProcess, numProcess, pagina, caderno)
            VALUES (?, ?, ?, ?, ?)
        ''', (dataHoraEnvio, dataHoraProcess, numProcess, pagina, caderno))
        conn.commit()

        # criar tabela pdf

    def CriarTabelaPdf():
        conn = sqlite3.connect('Jus.db')
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS Pdf (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                Nome TEXT
            )
        ''')
        conn.commit()

        # verifica pdf

    def VerificarPdf(pdfNome):
        conn = sqlite3.connect('Jus.db')
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM Pdf")
        pdfs = cursor.fetchall()
        for pdf in pdfs:
            if pdf[1] == pdfNome:
                return False
        return True
    
    # cadstra pdf

    def CadastrarPdf( pdfNome):
        conn = sqlite3.connect('Jus.db')
        cursor = conn.cursor()
        if BancoDados.VerificarPdf(pdfNome):
            cursor.execute("INSERT INTO Pdf (Nome) VALUES (?)", (pdfNome,))
            conn.commit()
            print("PDF cadastrado com sucesso!")
        else:
            print("Já existe um PDF com o mesmo nome!")
    def FecharConexao():
        conn = sqlite3.connect('Jus.db')
        conn.close()

        #criar tabelas

BancoDados.CriaTabelaJus()
BancoDados.CriarTabelaPdf()
BancoDados.FecharConexao()

# acessa forms

class GoogleForms:
    def AbrirGoogleForm():
        UrlGoogle = 'https://docs.google.com/forms/d/e/1FAIpQLSc6qpe_qr6gYa5ckCcxXv-2jO7DL4Qwi_ovFgcmee59uFDe6A/viewform'
        
        #adicionar os 10 registros no google

        NavegadorGoogleForm = webdriver.Chrome()
        NavegadorGoogleForm.get(UrlGoogle)
        return NavegadorGoogleForm
    
    def Atedez(limpar='Nao'):
        if not hasattr(GoogleForms.Atedez, 'AteDezz'):
            GoogleForms.Atedez.AteDezz = 0
        GoogleForms.Atedez.AteDezz += 1
        if limpar == 'Sim':
            GoogleForms.Atedez.AteDezz = 0
        return GoogleForms.Atedez.AteDezz
    
    def Inserir(DataProcessoo, HoraProcessoo, MinutoProcessoo, NumeroProcessoo, PaginaProcessoo, Cadernoo, Obss):
        limite = GoogleForms.Atedez()
        if limite <= 10:
            NavegadorGoogleForm = GoogleForms.AbrirGoogleForm()
        
            #Meu Nome
            Nome = WebDriverWait(NavegadorGoogleForm, 100).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, 'input.whsOnd.zHQkBf[jsname="YPqjbf"][autocomplete="off"][tabindex="0"][aria-labelledby="i1"][aria-describedby="i2 i3"][required][dir="auto"][data-initial-dir="auto"][data-initial-value=""]'))
            )
            meuNome = "Alexandre"
            Nome.send_keys(meuNome)

            #data de envio
            dataEnvio = WebDriverWait(NavegadorGoogleForm, 100).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, 'input[type="date"].whsOnd.zHQkBf[jsname="YPqjbf"][aria-labelledby="i9"]'))
            )
            dataEnvio.clear()
            dataEnvio.send_keys(datetime.date.today().strftime('%d/%m/%Y'))

            #hora atual
            horaAtual = WebDriverWait(NavegadorGoogleForm, 100).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, 'input[aria-label="Hora"]'))
            )
            horaAtual.clear()
            horaAtual.send_keys(datetime.datetime.today().strftime("%H"))

            #minuto Atual
            MinutoAtual = WebDriverWait(NavegadorGoogleForm, 100).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, 'input[aria-label="Minuto"]'))
            )
            MinutoAtual.clear()
            MinutoAtual.send_keys(datetime.datetime.today().strftime("%M"))

            #data processo
            DataProcesso = WebDriverWait(NavegadorGoogleForm, 100).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, 'input[type="date"].whsOnd.zHQkBf[jsname="YPqjbf"][aria-labelledby="i18"]'))
            )
            DataProcesso.clear()
            DataProcesso.send_keys(DataProcessoo.strftime('%d/%m/%Y'))

            #Hora processo
            HoraProcesso =  WebDriverWait(NavegadorGoogleForm, 100).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, 'input.whsOnd.zHQkBf[jsname="YPqjbf"][autocomplete="off"][tabindex="0"][aria-label="Hora"][maxlength="2"][min="0"][max="23"][role="combobox"][data-initial-value=""][jsaction="click:.CLIENT;keydown:.CLIENT;change:.CLIENT;blur:.CLIENT"]'))
            )
            HoraProcesso.clear()
            HoraProcesso.send_keys(HoraProcessoo)

            #Numero processo
            MinutoProcesso = WebDriverWait(NavegadorGoogleForm, 100).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, 'input.whsOnd.zHQkBf[jsname="YPqjbf"][autocomplete="off"][tabindex="0"][aria-label="Minuto"][maxlength="2"][min="0"][max="59"][role="combobox"][data-initial-value=""][jsaction="click:.CLIENT;keydown:.CLIENT;change:.CLIENT;blur:.CLIENT"]'))
            )
            MinutoProcesso.clear()
            MinutoProcesso.send_keys(MinutoProcessoo)


            #NumeroProcesso
            NumeroProcesso = WebDriverWait(NavegadorGoogleForm, 100).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, 'input.whsOnd.zHQkBf[jsname="YPqjbf"][autocomplete="off"][tabindex="0"][aria-labelledby="i23"][aria-describedby="i24 i25"][required][dir="auto"][data-initial-dir="auto"][data-initial-value=""]'))
            )
            NumeroProcesso.clear()
            NumeroProcesso.send_keys(NumeroProcessoo)

            # PaginaProcesso
            PaginaProcesso = WebDriverWait(NavegadorGoogleForm, 100).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, 'input[aria-labelledby="i27"]'))
            )
            PaginaProcesso.clear()
            PaginaProcesso.send_keys(PaginaProcessoo)

            # Caderno
            Caderno = WebDriverWait(NavegadorGoogleForm, 100).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, 'input[aria-labelledby="i31"]'))
            )
            Caderno.clear()
            Caderno.send_keys(Cadernoo)
                
            # Obs
            Obs = WebDriverWait(NavegadorGoogleForm, 100).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, 'textarea[aria-labelledby="i35"]'))
            )
            Obs.clear()
            Obs.send_keys(Obss)

            button = WebDriverWait(NavegadorGoogleForm, 100).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "span[class='NPEfkd RveJvd snByac']")))
            time.sleep(5)

            # Clicar no botão

            button.click()
            time.sleep(5)
            NavegadorGoogleForm.quit()


# UrlJus para abrir jus.br

UrlJus = 'https://dejt.jt.jus.br/dejt/f/n/diariocon'

# Configura as opções do Chrome para download

chromeOptions = webdriver.ChromeOptions()
pasta = "C:\\Users\\rocha\\OneDrive - Fatec Centro Paula Souza\\Desktop\\one4_Fim"
chromeOptions.add_experimental_option("prefs", {"download.default_directory": pasta})

# Cria o objeto Serviço e Chrome

chromedriver_path = ChromeDriverManager().install()
service = Service(chromedriver_path)
navegador = webdriver.Chrome(service=service, options=chromeOptions)

# Acessa o site

navegador.get(UrlJus)

# data início = anteontem

datainicial = datetime.date.today() - datetime.timedelta(days=2)
datainicial_input =  WebDriverWait(navegador, 20).until(
    EC.visibility_of_element_located((By.CSS_SELECTOR, "input[name='corpo:formulario:dataIni']")))
datainicial_input.clear()
datainicial_input.send_keys(datainicial.strftime('%d/%m/%Y'))


# data fim = hoje

DataFinal = WebDriverWait(navegador, 20).until(
    EC.visibility_of_element_located((By.CSS_SELECTOR,"input[name='corpo:formulario:dataFim']")))
DataFinal.clear()
DataFinal.send_keys(datetime.date.today().strftime('%d/%m/%Y'))

# caderno = judiciário

caderno = Select(WebDriverWait(navegador, 20).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, "select[name='corpo:formulario:tipoCaderno']"))))
caderno.select_by_visible_text('Judiciário')


# órgão = TST

orgao = Select(WebDriverWait(navegador, 20).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, "select[name='corpo:formulario:tribunal']"))))
orgao.select_by_visible_text('TST')


# botão pesquisar

BtnPesquisar = WebDriverWait(navegador, 10).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, "button[name='corpo:formulario:botaoAcaoPesquisar']")))
BtnPesquisar.click()

time.sleep(10)

try:
    # Pega tabela
    table = WebDriverWait(navegador, 60).until(
    EC.visibility_of_element_located((By.CSS_SELECTOR, "div#diarioCon table")))

    # Pega as linhas
    linhas = table.find_elements(By.TAG_NAME, "tr")

    # Itera sobre cada caderno
    # lê linhas

    for linha in linhas[1:]:

        # Encontrar o link para baixar o caderno

        link_baixar = linha.find_element(By.CSS_SELECTOR, "button.bt")
        pdfNome =""

        # Clicar no link para baixar o caderno
        # botão
        link_baixar.click()

       # lista arquivos

        arquivosPDF = os.listdir(pasta)
        pdfs = [arquivo for arquivo in arquivosPDF if arquivo.endswith(".pdf")]
        quantidade_pdfs = len(pdfs)
        while True:
            time.sleep(5)
            arquivosPDFAtu = os.listdir(pasta)
            pdfsAtu = [arquivo for arquivo in arquivosPDFAtu if arquivo.endswith(".pdf")]
            quantidade_pdfsAtu = len(pdfsAtu)
            if quantidade_pdfsAtu > quantidade_pdfs:
                break

        #Loop para percorrer cada arquivo baixado

        #fazer loop para verificar se o pdf foi baixado

        for file_name in glob.glob(pasta + '/*.pdf'):
            if BancoDados.VerificarPdf(file_name.split("\\")[-1]):

                # Abrir arquivo pdf

                with open(file_name, 'rb') as pdf_file:
                    pdf_reader = PyPDF2.PdfReader(pdf_file)

                        # Loop para percorrer cada página do arquivo pdf

                    for pagina in range(len(pdf_reader.pages)):
                        page = pdf_reader.pages[pagina]
                        text = page.extract_text()

                        texto = text.split('\n')
                        for linha in texto:
                            for match in re.finditer(r'PROCESSO N[ºN°.]\s*TST\S*\b\d+\.\d+\.\d+\.\d+\b', linha):

                                # Pega o Processo e o Numero

                                processo = match.group(0)

                                # Verifica se ele termina com 0

                                if processo.endswith('0'):

                                    #Print para fins de teste

                                    print('Processo válido: ', processo)

                                    # Salvar informações na planilha de log (Criei essa planilha pra ser alimentada
                                    
                                    # e conter as informações que está pedindo no exercicio)

                                    # Criar uma nova linha com as informações desejadas

                                    datarecebida = datetime.datetime.now()
                                    pdfNome = file_name.split("\\")[-1]

                                    planilha = pd.read_excel("C:\\Users\\rocha\\OneDrive - Fatec Centro Paula Souza\\Desktop\\one4_Fim\\log.xlsx")
                                    nova_linha = pd.DataFrame([[datarecebida.strftime("%Y-%m-%d %H:%M:%S"), processo, pagina+1, pdfNome]], columns=['Data/Hora', 'Número do Processo', 'Página', 'Caderno'])

                                    # Adicionar a nova linha ao DataFrame existente

                                    planilha = pd.concat([planilha, nova_linha], ignore_index=True)

                                    # Salvar o DataFrame atualizado no arquivo Excel

                                    planilha.to_excel("C:\\Users\\rocha\\OneDrive - Fatec Centro Paula Souza\\Desktop\\one4_Fim\\log.xlsx", index=False)


                                    time.sleep(10)
                                    
                                    #google

                                    GoogleForms.Inserir(datarecebida, datarecebida.strftime("%H"),datarecebida.strftime("%M"),processo,pagina+1,pdfNome, "Alexandre")
                                    BancoDados.Cadastrar(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), datarecebida.strftime("%Y-%m-%d %H:%M:%S"),processo,pagina+1,pdfNome)

                                else:
                                    print('Processo inválido: ', processo)

                Limpar = GoogleForms.Atedez(limpar='Sim')
                BancoDados.CadastrarPdf(pdfNome)            


    navegador.quit()

except NoSuchElementException as e:
    print(f"Aconteceu um erro durante tentar realizar algum processo, ERRO: {str(e)}")

