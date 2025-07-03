from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pandas as pd
import os

driver = webdriver.Chrome()

caminho = os.getcwd()
arquivo_html = caminho + r"\login.html"
driver.get(arquivo_html)
# limpar o campo de login e senha
driver.find_element(By.XPATH, '/html/body/div/form/input[1]').clear()
driver.find_element(By.XPATH, '/html/body/div/form/input[2]').clear()
# preencher login
driver.find_element(By.XPATH, '/html/body/div/form/input[1]').send_keys('admin')
# preencher senha
driver.find_element(By.XPATH, '/html/body/div/form/input[2]').send_keys('123456')
# clicar no botao
driver.find_element(By.XPATH, '/html/body/div/form/button').click()
dados_form = pd.read_excel('NotasEmitir.xlsx')


from selenium.webdriver.support.ui import Select
contador = 0
while  True: 
    form_nome = driver.find_element(By.XPATH, '//*[@id="nome"]')
    form_nome.clear()
    # preencher o campo nome
    for nome in range(len(dados_form['Cliente'])):
        nome = dados_form['Cliente'][contador]
        form_nome.send_keys(nome)
        break 
        
    # preencher o campo endereço
    form_endereco = driver.find_element(By.NAME, 'endereco')
    form_endereco.clear()
    for ender in range(len(dados_form['Endereço'])):
        ender = dados_form['Endereço'][contador]
        form_endereco.send_keys(ender)
        break

    #preencher bairro
    form_bairro = driver.find_element(By.NAME, 'bairro')
    form_bairro.clear()
    for bairro in range(len(dados_form['Bairro'])):
        bairro = dados_form['Bairro'][contador]
        form_bairro.send_keys(bairro)
        break

    # preencher muinícipio
    form_municipio = driver.find_element(By.NAME, 'municipio')
    form_municipio.clear()
    for municipio in range(len(dados_form['Municipio'])):
        municipio = dados_form['Municipio'][contador]
        form_municipio.send_keys(municipio)
        break

    # preencher cep
    form_cep = driver.find_element(By.NAME, 'cep')
    form_cep.clear()
    for cep in range(len(dados_form['CEP'])):
        cep = str(dados_form['CEP'][contador])
        form_cep.send_keys(cep)
        break

    # preencher uf
    uf = driver.find_element(By.TAG_NAME, 'select')
    uf_select = Select(uf)
    for estado in range(len(dados_form['UF'])):
        estado = dados_form['UF'][contador]
        uf_select.select_by_visible_text(estado)
        break

    # preencher cnpj/cpf
    form_cnpj = driver.find_element(By.NAME, 'cnpj')
    form_cnpj.clear()
    for cnpj in range(len(dados_form['CPF/CNPJ'])):
        cnpj = str(dados_form['CPF/CNPJ'][contador])
        form_cnpj.send_keys(cnpj)
        break

    # preencher incrição estadual
    form_ie = driver.find_element(By.NAME, 'inscricao')
    form_ie.clear()
    for inscricao in range(len(dados_form['Inscricao Estadual'])):
        inscricao = int(dados_form['Inscricao Estadual'][contador])
        form_ie.send_keys(inscricao)
        break

    # preencher descrição produto
    form_produto = driver.find_element(By.NAME, 'descricao')
    form_produto.clear()
    for produto in range(len(dados_form['Descrição'])):
        produto = dados_form['Descrição'][contador]
        form_produto.send_keys(produto)
        break

    # preencher quantidade
    form_quantidade = driver.find_element(By.NAME, 'quantidade')
    form_quantidade.clear()
    for quantidade in range(len(dados_form['Quantidade'])):
        quantidade = int(dados_form['Quantidade'][contador])
        form_quantidade.send_keys(quantidade)
        break

    # preencher valor unitário
    form_valor = driver.find_element(By.NAME, 'valor_unitario')
    form_valor.clear()
    for valor in range(len(dados_form['Valor Unitario'])):
        valor = float(dados_form['Valor Unitario'][contador])
        form_valor.send_keys(valor)
        break

    # preencher valor total
    form_valor_total = driver.find_element(By.NAME, 'total')
    form_valor_total.clear()
    for total in range(len(dados_form['Valor Total'])):
        total = float(dados_form['Valor Total'][contador])
        form_valor_total.send_keys(total)
        break
    # clicar no botao emitir
    driver.find_element(By.XPATH, '//*[@id="formulario"]/button').click()
    
    contador += 1
    if contador == len(dados_form['Cliente']):
        break
