import pyautogui, pyperclip, time
import pandas as pd

def main():
    # este projeto foi feito para ser executado pelo terminal do windows, também pode ser executado pelo terminal interno do seu IDE
    # quando for executar, o browser deve estar em segundo plano por conta do alt+tab da função DownloadDatabase
    # downloadDatabase()
    writeEmail()

def writeEmail():
    print('antes')
    tabela = pd.read_excel(r'C:/Users/user/Downloads/Vendas - Dez.xlsx')
    print('depois')

    # variável que contém o e-mail destinatário
    email = 'rkatchau@gmail.com'

    faturamento = getFaturamento(tabela)
    quantidade_de_vendas = getQuantidadeVendas(tabela)

    # nova aba
    pyautogui.hotkey('ctrl', 't')

    # escreve url do gmail e enter
    pyautogui.write('https://mail.google.com/mail/u/0/#inbox')
    enter()
    time.sleep(5)

    # aperta no botão + (escrever) do gmail
    pyautogui.click(42, 211)
    time.sleep(6)

    # escreve o destinatário
    pyautogui.write(email)

    # tab para confirmar o email e selecionar o próximo campo
    pyautogui.hotkey('tab')
    pyautogui.hotkey('tab')

    # escreve o assunto e seleciona o próximo campo
    pyperclip.copy('Relatório de Vendas')
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.hotkey('tab')

    # escreve o corpo e envia e-mail
    corpo = f'''
    Prezados, bom dia

    O faturamento de ontem foi de: R${faturamento:,.2f}
    Ontem foram vendidos {quantidade_de_vendas:,} produtos

    Abs.

    Marcos
    '''
    pyautogui.write(corpo)
    pyautogui.hotkey('ctrl', 'enter')

def downloadDatabase():

    # trocar janela pro browser
    pyautogui.hotkey('alt', 'tab')

    # nova aba
    pyautogui.hotkey('ctrl', 't')

    # escreve o link na url e aperta enter
    pyautogui.write('https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga')
    enter()
    time.sleep(5)

    # abre a pasta Exportar
    pyautogui.click(487, 338, clicks = 2)
    time.sleep(4)

    # botão direito no arquivo a ser baixado
    pyautogui.click(487, 338, button = 'right')
    time.sleep(3)

    # aperta em 'Fazer download'
    pyautogui.click(562, 932)
    time.sleep(4)

    # ATENÇÃO: sua pasta padrão para downloads deve ser C:/Users/seu_user/Downloads
    # aperta em 'Salvar' quando a janela do Explorador de Arquivos é chamada
    pyautogui.click(750, 558)


def getQuantidadeVendas(tabela):
    return tabela['Quantidade'].sum()

def getFaturamento(tabela):
    return tabela['Valor Final'].sum()

def enter():
    pyautogui.press('enter')

# função para ser utilizada quando quiser pegar a posição de algo na tela, comente a chamada da main e chame ela
def getPosition(timer):
    time.sleep(timer)

    posicao = pyautogui.position()
    print(posicao)

main()
# getPosition(3)
