import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.select import Select
import pyautogui as pgui
import warnings

warnings.filterwarnings('ignore')

chrome_options = Options()
chrome_options.add_argument('--headless')

#versão mais nova não precisa
#Chromedriver = "chromedriver.exe"

#chrome invisivel
#driver = webdriver.Chrome(options=chrome_options)

#Abre o chrome visivel
driver = webdriver.Chrome()

caminhoBase = "arquivosBase/baseExemplo.csv"
df = pd.read_csv("{}".format(caminhoBase), converters={'cnpj': lambda x: str(x)}, on_bad_lines='skip', sep=";", encoding="utf-8")
dataRefN = df['data'].tolist()

mesRef = "072022"

confirmarIndice = pgui.confirm(text='Selecione o Índice que será utilizado', title='Índices suportados', buttons=[
    'IGP-M (FGV) - a partir de 06/1989',
    'IGP-DI (FGV) - a partir de 02/1944',
    'INPC (IBGE) - a partir de 04/1979',
    'IPCA (IBGE) - a partir de 01/1980',
    'IPCA-E (IBGE) - a partir de 01/1992',
    'IPC-BRASIL (FGV) - a partir de 01/1990',
    'IPC-SP (FIPE) - a partir de 11/1942'
    ])

if confirmarIndice == 'IGP-M (FGV) - a partir de 06/1989':
    indice = '28655IGP-M'
elif confirmarIndice == 'IGP-DI (FGV) - a partir de 02/1944':
    indice = '00190IGP-DI'
elif confirmarIndice == 'INPC (IBGE) - a partir de 04/1979':
    indice = '00188INPC'
elif confirmarIndice == 'IPCA (IBGE) - a partir de 01/1980':
    indice = '00433IPCA'
elif confirmarIndice == 'IPCA-E (IBGE) - a partir de 01/1992':
    indice = '10764IPC-E'
elif confirmarIndice == 'IPC-BRASIL (FGV) - a partir de 01/1990':
    indice = '00191IPC-BRASIL'
elif confirmarIndice == 'IPC-SP (FIPE) - a partir de 11/1942':
    indice = '00193IPC-SP'
else:
    pgui.alert(text='Selecione um índice antes da execução', title='Aviso', button="Ok")
    driver.quit()
    quit()



valorAtualN = df['valor']
nomeEmpresaN = df['nome']
notaFiscalN = df['nf']

contador = 1

listaNome = []
listaNF = []
listaDebitoAnterior = []
listaPorcentual = []
listaDebitoCorrigido = []
listaDataRef = []

for i in range(len(nomeEmpresaN)):
    dataRef = dataRefN[i]
    dataOriginal = dataRefN[i]
    valorAtual = valorAtualN[i]
    nomeEmpresa = nomeEmpresaN[i]
    notaFiscal = notaFiscalN[i]

    hora = datetime.today().strftime('%H:%M')

    print('{0} | Nome: {1} | Débito: {2}| Hora: {3}'.format(contador, nomeEmpresa, valorAtual, hora))

    dataRef = dataRef.replace("/", "")
    dataRef = dataRef[2:]

    driver.get("https://www3.bcb.gov.br/CALCIDADAO/publico/corrigirPorIndice.do?method=corrigirPorIndice")

    select_element = driver.find_element(By.XPATH,'/html/body/div[6]/table/tbody/tr[3]/td/div/form/div[1]/table/tbody/tr[3]/td[2]/select')
    
    select_object = Select(select_element)
    
    select_object.select_by_value(indice)

    driver.find_element(By.XPATH, '//*[@id="corrigirPorIndiceForm"]/div[1]/table/tbody/tr[4]/td[2]/input').send_keys(dataRef)

    driver.find_element(By.XPATH, '//*[@id="corrigirPorIndiceForm"]/div[1]/table/tbody/tr[5]/td[2]/input').send_keys(mesRef)

    driver.find_element(By.XPATH, '//*[@id="corrigirPorIndiceForm"]/div[1]/table/tbody/tr[6]/td[2]/input').send_keys(valorAtual)

    driver.find_element(By.XPATH, '//*[@id="corrigirPorIndiceForm"]/div[2]/input[1]').click()

    valorPercentual = driver.find_element(By.XPATH, '/html/body/div[6]/table/tbody/tr/td/div[2]/table[1]/tbody/tr[7]/td[2]').text
    valorCorrigido = driver.find_element(By.XPATH, '/html/body/div[6]/table/tbody/tr/td/div[2]/table[1]/tbody/tr[8]/td[2]').text

    valorPercentual = valorPercentual.replace(" ", "")
    valorCorrigido = valorCorrigido.replace("R$", "").replace("REAL", "").replace(" ", "").replace("(", "").replace(")", "")

    print('''  | Valor percentual: {0} | Débito Corrigido: R$ {0}
-------------------------------------------------------------'''.format(valorPercentual, valorCorrigido))

    listaNome.append(nomeEmpresa)
    listaNF.append(notaFiscal)
    listaDataRef.append(dataOriginal)
    listaDebitoAnterior.append(valorAtual)
    listaPorcentual.append(valorPercentual)
    listaDebitoCorrigido.append(valorCorrigido)

    contador += 1

baseListagem = []

for i in range(len(listaNome)):
    baseListagem.append("{};{};{};{};{};{}\n".format(listaNome[i], listaNF[i], listaDataRef[i],listaDebitoAnterior[i], listaPorcentual[i], listaDebitoCorrigido[i]))

f = open('arquivosBase/baseExemploFinalizada.csv', 'w')
f.write("NOME;NF;REF;DEBITO;PORCENTAGEM;CORRECAO;\n")
f.writelines(baseListagem)
f.close()

driver.quit()

pgui.alert(text='Correções foram finalizadas com sucesso!', title='Aviso', button="Fechar")