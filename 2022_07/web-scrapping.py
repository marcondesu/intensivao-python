from typing import KeysView
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

import pandas as pd

## Passos
# Pegar a cotação do dólar
# Pegar a cotação do euro
# Pegar a cotação do ouro
# Atualizar a base de dados
# Recalcular os preços
# Exportar a base de dados

def main():
    # abrir o navegador
    navegador = webdriver.Chrome()

    dolar = getCotacaoDolar(navegador)
    euro = getCotacaoEuro(navegador)
    ouro = getCotacaoOuro(navegador)

    # fecha o navegador
    navegador.quit()

    # carrega tabela
    tabela = pd.read_excel('./Produtos.xlsx')

    # atualiza tabela com o valor dos preços atualizados
    tabela = recalcularPrecos(tabela, dolar, euro, ouro)

    # corrige detalhes de formatação
    tabela = corrigirFormatacoes(tabela)

    # exportar tabela
    tabela.to_excel('./Produtos - Atualizado.xlsx', index = False)

def corrigirFormatacoes(tabela):
    tabela['Preço Original'] = tabela['Preço Original'].map('R$ {:.2f}'.format)
    tabela['Cotação'] = tabela['Cotação'].map('R$ {:.2f}'.format)
    tabela['Preço de Compra'] = tabela['Preço de Compra'].map('R$ {:.2f}'.format)
    tabela['Preço de Venda'] = tabela['Preço de Venda'].map('R$ {:.2f}'.format)

    return tabela

def recalcularPrecos(tabela, dolar, euro, ouro):
    tabela = atualizarCotacoes(tabela, dolar, euro, ouro)
    
    # atualizar preço de compra (Preço Original * Cotação)
    tabela['Preço de Compra'] = tabela['Preço Original'] * tabela['Cotação']

    # atualizar preço de venda (Preço de Compra * Margem)
    tabela['Preço de Venda'] = tabela['Preço de Compra'] * tabela['Margem']

    return tabela

def atualizarCotacoes(tabela, dolar, euro, ouro):
    # na linha em que a coluna 'Moeda' tiver o valor 'Dólar', substituir a coluna 'Cotação' para o valor da variável dolar
    tabela.loc[tabela['Moeda'] == 'Dólar', 'Cotação'] = float(dolar)

    # na linha em que a coluna 'Moeda' tiver o valor 'Euro', substituir a coluna 'Cotação' para o valor da variável euro
    tabela.loc[tabela['Moeda'] == 'Euro', 'Cotação'] = float(euro)

    # na linha em que a coluna 'Moeda' tiver o valor 'Ouro', substituir a coluna 'Cotação' para o valor da variável ouro
    tabela.loc[tabela['Moeda'] == 'Ouro', 'Cotação'] = float(ouro)

    return tabela

def getCotacaoOuro(navegador):
    navegador.get('https://www.melhorcambio.com/ouro-hoje')

    # pega a cotação do ouro
    cotacao_ouro = navegador.find_element('xpath', '//*[@id="comercial"]').get_attribute('value')

    # formatando valor recebido da cotação com vírgula para ponto
    cotacao_ouro = cotacao_ouro.replace(',', '.')

    return cotacao_ouro

def getCotacaoEuro(navegador):
    # abre o site do google
    navegador.get('https://www.google.com.br/')

    # pesquisar cotação do euro
    navegador.find_element('xpath', '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys('cotação euro')

    # clica no enter para pesquisar
    navegador.find_element('xpath', '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)

    # pega a cotação do euro
    cotacao_euro = navegador.find_element('xpath', '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
    
    return cotacao_euro    

def getCotacaoDolar(navegador):
    # abre o site do google
    navegador.get('https://www.google.com.br/')

    # pesquisar cotação do dólar
    navegador.find_element('xpath', '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys('cotação dólar')

    # clica no enter para pesquisar
    navegador.find_element('xpath', '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)

    # pega a cotação do dólar
    cotacao_dolar = navegador.find_element('xpath', '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
    
    return cotacao_dolar

main()
