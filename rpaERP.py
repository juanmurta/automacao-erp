import pyautogui
import subprocess
import time
import pandas as pd


def esperar(caminho):  # função que aguarda determinada imagem aparecer no sistema.
    while True:  # esperando a tela do sistema aparecer
        try:
            pyautogui.locateOnScreen(f'{caminho}', confidence=0.8)
            break
        except pyautogui.ImageNotFoundException:
            time.sleep(1)
    encontrou = pyautogui.locateCenterOnScreen(fr'{caminho}', confidence=0.8)
    return encontrou


pyautogui.alert('O código vai começar. Não mexa em NADA enquanto o código tiver rodando. Quando finalizar, eu te aviso')

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.5

# abrir o programa no nosso computador
subprocess.Popen([r"C:\Makemoney\Gestor.exe"])

# preencher o login do ERP
encontrou = esperar(r'C:\Users\Juan\Desktop\Cursos\Impressionador\Rpa\Arquivo\login.png')
pyautogui.write('2')
pyautogui.press('tab')
pyautogui.write('Avisnet')
pyautogui.press('tab')
pyautogui.write('0379')
pyautogui.press('enter')

encontrou = esperar(r'C:\Users\Juan\Desktop\Cursos\Impressionador\Rpa\Arquivo\produto.png')
pyautogui.click(encontrou)

encontrou = esperar(r'C:\Users\Juan\Desktop\Cursos\Impressionador\Rpa\Arquivo\consulta.png')
pyautogui.click(encontrou)

#  caso queira clicar a direita da imagem imagem = imagem[0] + imagem [2], imagem[1] + imagem[3]/2

# abrindo a planilha com os produtos
produtos = pd.read_excel(r'Produtos.xlsx')

#preenchendo as informações no ERP
for linha in produtos.index:
    encontrou = esperar(r'C:\Users\Juan\Desktop\Cursos\Impressionador\Rpa\Arquivo\incluir.png')
    pyautogui.click(encontrou)
    imagem = produtos.loc[linha, 'Imagem']
    venda = produtos.loc[linha, 'Venda']
    venda_texto = f'{venda:.2f}'.replace(',', '.')
    custo = produtos.loc[linha, 'Custo']
    custo_texto = f'{custo:.2f}'.replace(',', '.')
    produto = produtos.loc[linha, 'Nome']

    pyautogui.press('tab')
    pyautogui.write(str(produtos.loc[linha, 'GTIN']))
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.write(str(produtos.loc[linha, 'Nome']))
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.write(str(produtos.loc[linha, 'NCM']))
    pyautogui.press('enter')
    pyautogui.press('enter')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.write(produtos.loc[linha, 'Categoria'])
    pyautogui.press('tab')
    pyautogui.write(produtos.loc[linha, 'SubCategoria'])

    # encontrar imagem
    encontrou = esperar(r'C:\Users\Juan\Desktop\Cursos\Impressionador\Rpa\Arquivo\imagem.png')
    pyautogui.rightClick(encontrou)
    encontrou = esperar(r'C:\Users\Juan\Desktop\Cursos\Impressionador\Rpa\Arquivo\click_imagem.png')
    pyautogui.click(encontrou)
    encontrou = esperar(r'C:\Users\Juan\Desktop\Cursos\Impressionador\Rpa\Arquivo\caminho.png')
    pyautogui.write(fr'C:\Users\Juan\Desktop\Cursos\Impressionador\Rpa\Imagens Produtos\{imagem}')
    pyautogui.press('enter')
    pyautogui.click(encontrou)
    pyautogui.press('f3')

    #tratando possiveis mensagens de erros.
    if esperar(r'C:\Users\Juan\Desktop\Cursos\Impressionador\Rpa\Arquivo\erro_ean.png'):
        pyautogui.press('enter')
    else:
        pass

    pyautogui.press('f8')
    encontrou = esperar(r'C:\Users\Juan\Desktop\Cursos\Impressionador\Rpa\Arquivo\esperar_estoque.png')
    pyautogui.press('f7')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.write(str(custo_texto))
    pyautogui.press('tab')
    pyautogui.write(str(custo_texto))
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.write(str(venda_texto))
    pyautogui.press('f3')
    pyautogui.press('f10')
    esperar(r'C:\Users\Juan\Desktop\Cursos\Impressionador\Rpa\Arquivo\estoque.png')

pyautogui.alert('O código terminou, pode pegar o seu computador de volta')

