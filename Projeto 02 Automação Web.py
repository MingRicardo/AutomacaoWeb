#!/usr/bin/env python
# coding: utf-8

# ### Inicialização do Navegador

# In[ ]:


from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas as pd
import time
import email.message

#criar navegador
nav = webdriver.Chrome()

#importar e visualizar banco de dados
tabela_produtos = pd.read_excel('buscas.xlsx')
display(tabela_produtos)


# ### Definição das funções do Google Shopping e Buscapé

# In[ ]:


def verificar_tem_termos_banidos(lista_termos_banidos, nome):
    tem_termos_banidos = False
    for palavra in lista_termos_banidos:
        if palavra in nome:
            tem_termos_banidos = True
    return tem_termos_banidos

def verificar_tem_todos_termos_produto(lista_termos_nome_produto, nome):
    tem_todos_termos_produtos = True
    for palavra in lista_termos_nome_produto:
        if palavra not in nome:
            tem_todos_termos_produtos = False  
    return tem_todos_termos_produtos

def busca_google_shopping(nav, produto, termos_banidos, preco_minimo, preco_maximo):
    produto = produto.lower()
    termos_banidos = termos_banidos.lower()
    lista_termos_banidos = termos_banidos.split(' ')
    lista_termos_nome_produto = produto.split(' ')
    preco_minimo = float(preco_minimo)
    preco_maximo = float(preco_maximo)
        
    lista_ofertas = []

    # entrar no google
    nav.get('https://www.google.com.br/')
    nav.find_element('xpath', '//*[@id="APjFqb"]' ).send_keys(produto, Keys.ENTER)
    time.sleep(2)

    # entrar na aba shopping
    elementos = nav.find_elements('class name', 'YmvwI')

    for item in elementos:
        if "Shopping" in item.text:
            item.click()
            break

    lista_resultados = nav.find_elements('class name', 'i0X6df')

    for resultado in lista_resultados:
        nome = resultado.find_element('class name', 'tAxDx').text
        nome = nome.lower()

        #analisar se não tem nenhum termo banido
        tem_termos_banidos = verificar_tem_termos_banidos(lista_termos_banidos, nome)

        # analisar se tem todos os nomes do produto
        tem_todos_termos_produtos = verificar_tem_todos_termos_produto(lista_termos_nome_produto, nome)

        #selecionar somente os elementos que: tem_termos_banidos = False   e ao mesmo tempo e tem_todos_termos_produtos = True
        if not tem_termos_banidos and tem_todos_termos_produtos:
            preco = resultado.find_element('class name', 'XrAfOe').text
            preco = preco.replace('R$', '').replace(' ', '').replace('.', '').replace(',', '.')
            preco = preco[:(preco.find(".")+3)]
            preco = float(preco)

            #se o preço tá no parâmetro estabelecido
            if preco_minimo <= preco <= preco_maximo:
                elemento_referencia = resultado.find_element('class name', 'EI11Pd')
                elemento_pai = elemento_referencia.find_element('xpath', '..')
                link = elemento_pai.get_attribute('href')

                lista_ofertas.append((nome, preco, link))
    return lista_ofertas

def busca_buscape(nav, produto, termos_banidos, preco_minimo, preco_maximo):
    produto = produto.lower()
    termos_banidos = termos_banidos.lower()
    lista_termos_banidos = termos_banidos.split(' ')
    lista_termos_nome_produto = produto.split(' ')
    preco_minimo = float(preco_minimo)
    preco_maximo = float(preco_maximo)
    
    lista_ofertas = []
    
    # entrar no buscapé
    nav.get('https://www.buscape.com.br/')
    nav.maximize_window()
    #nav.find_element('class name', 'Button_button__bSiTN Button_system__Z0Mv0 Button_outline__CHLBt PrivacyPolicy_Button__1RxwB').send_keys(Keys.ENTER)
    nav.find_element('xpath', '//*[@id="new-header"]/div[1]/div/div/div[3]/div/div/div[2]/div/div[1]/input').send_keys(produto, Keys.ENTER)
    #time.sleep(5)
    while len(nav.find_elements('class name', 'Select_Select__1HNob')) < 1:
       time.sleep(1)
    
    lista_resultados = nav.find_elements('class name', 'ProductCard_ProductCard_Inner__gapsh')
    
    #time.sleep(10)
        
    for resultado in lista_resultados:
        preco = resultado.find_element('class name', 'Text_MobileHeadingS__HEz7L').text
        nome = resultado.find_element('class name', 'ProductCard_ProductCard_Name__U_mUQ').text
        nome = nome.lower()
        link = resultado.get_attribute('href')

        tem_termos_banidos = verificar_tem_termos_banidos(lista_termos_banidos, nome)
        tem_todos_termos_produtos = verificar_tem_todos_termos_produto(lista_termos_nome_produto, nome)
        
        if not tem_termos_banidos and tem_todos_termos_produtos:
            preco = preco.replace('R$', '').replace(' ', '').replace('.', '').replace(',', '.')
            preco = preco[:(preco.find(".")+3)]
            preco = float(preco)
            
            if preco_minimo <= preco <= preco_maximo:
                lista_ofertas.append((nome, preco, link))
    return lista_ofertas


# ### Construção da lista de ofertas

# In[ ]:


tabela_ofertas = pd.DataFrame()

for linha in tabela_produtos.index:
    # pesquisar pelo produto
    produto = tabela_produtos.loc[linha, 'Nome']
    termos_banidos = tabela_produtos.loc[linha, 'Termos banidos']
    preco_minimo = tabela_produtos.loc[linha, 'Preço mínimo']
    preco_maximo = tabela_produtos.loc[linha, 'Preço máximo']
    
    lista_ofertas_google_shopping = busca_google_shopping(nav, produto, termos_banidos, preco_minimo, preco_maximo)
    if lista_ofertas_google_shopping:
        tabela_google_shopping = pd.DataFrame(lista_ofertas_google_shopping, columns = ['Produto', 'Preço', 'Link'])
        tabela_ofertas = pd.concat([tabela_ofertas, tabela_google_shopping])
    else:
        tabela_google_shopping = None
        
    lista_ofertas_buscape = busca_buscape(nav, produto, termos_banidos, preco_minimo, preco_maximo)
    if lista_ofertas_buscape:
        tabela_buscape = pd.DataFrame(lista_ofertas_buscape, columns = ['Produto', 'Preço', 'Link'])
        tabela_ofertas = pd.concat([tabela_ofertas, tabela_buscape])
    else:
        tabela_buscape = None
        
display(tabela_ofertas)
    


# ### Exportar para excel

# In[ ]:


tabela_ofertas.to_excel('Ofertas.xlsx', index = False)


# ### Enviar resultado final para e-mail

# In[ ]:


if len(tabela_ofertas) > 0:

    def enviar_email():  
        corpo_email = f"""
        <p> Prezados, </p>
        <p> Segue tabela com ofertas encontradas, </p>
        <p> Atenciosamente, </p>
        <p> Ricardo Ming, </p>

        {tabela_ofertas.to_html(index=False)}

        """

        msg = email.message.Message()
        msg['Subject'] = "Tabela ofertas"
        msg['From'] = 'ricardoxming@gmail.com'
        msg['To'] = 'ricardoxming@gmail.com'
        password = '***************' 
        msg.add_header('Content-Type', 'text/html')
        msg.set_payload(corpo_email )

        s = smtplib.SMTP('smtp.gmail.com: 587')
        s.starttls()
        # Login Credentials for sending the mail
        s.login(msg['From'], password)
        s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
        print('Email enviado')
enviar_email()

nav.quit()

