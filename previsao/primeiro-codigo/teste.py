import asyncio
from pyppeteer import launch
import pandas as pd
import time

# Função para aguardar o overlay de "aguarde" desaparecer
async def wait_for_overlay_to_disappear(page):
    overlay_selector = '#progress'
    try:
        await page.waitForSelector(overlay_selector, {'hidden': True})
        print('Overlay desapareceu.')
    except:
        print('Overlay não encontrado seguindo o fluxo')

async def wait_a_moment(tempo):
    await asyncio.sleep(tempo)

async def seta_login(page):
    username_selector = 'input[name="box1:usuario_tela."]'
    await page.waitForSelector(username_selector)
    await page.type(username_selector, 'MATHEUS')

    password_selector = 'input[name="box1:senha_tela."]'
    await page.waitForSelector(password_selector)
    await page.type(password_selector, 'matteus')

    await wait_a_moment(5)

    await page.keyboard.press('F9')

def get_dados_tabela(caminho_arquivo):
    df = pd.read_excel(caminho_arquivo)
    return df.to_dict(orient='records')

async def atualiza_estoque():
    local_arquivo = '../planilha.xlsx'
    dados = get_dados_tabela(local_arquivo)
    print(dados)

    opcoes_inicio = {'headless': False, 'args': ['--start-maximized']}
    print('Iniciando Puppeteer com as opções:', opcoes_inicio)
    browser = await launch(options=opcoes_inicio, executablePath='C:\Program Files\Google\Chrome\Application\chrome.exe')
    # browser = await launch(options=opcoes_inicio)
    page = await browser.newPage()

    # Ajustando o site com o tamanho da tela
    await page.setViewport({'width': 1370, 'height': 1024})

    file_url = 'https://divero.systextil.com.br/systextil/'
    await page.goto(file_url)

    await asyncio.sleep(2)

    await seta_login(page)

    await wait_for_overlay_to_disappear(page)

    await asyncio.sleep(2)

    await page.keyboard.type('estq_f950')

    await page.keyboard.press('Enter')

    await asyncio.sleep(2)

    for item in dados:
        await wait_for_overlay_to_disappear(page)
        await asyncio.sleep(2)

        print('--------------------------------------------------')
        print(item['Nivel'])
        await page.type('input[name="nivel."]', str(item['Nivel']))
        await page.keyboard.press('Enter')
        await asyncio.sleep(2)
        await wait_for_overlay_to_disappear(page)

        print(item['Grupo'])
        await page.type('input[name="grupo."]', str(item['Grupo']))
        await page.keyboard.press('Enter')
        await asyncio.sleep(2)
        await wait_for_overlay_to_disappear(page)

        print(item['Subgrupo'])
        await page.type('input[name="subgrupo."]', str(item['Subgrupo']))
        await page.keyboard.press('Enter')
        await asyncio.sleep(2)
        await wait_for_overlay_to_disappear(page)

        print(item['Item'])
        await page.type('input[name="item."]', str(item['Item']))
        await page.keyboard.press('Enter')
        await asyncio.sleep(2)
        await wait_for_overlay_to_disappear(page)

        print(item['Deposito'])
        await page.type('input[name="cod_deposito."]', str(item['Deposito']))
        await page.keyboard.press('Enter')
        await asyncio.sleep(2)
        await wait_for_overlay_to_disappear(page)

        print(item['Lote'])
        await page.type('input[name="lote."]', str(item['Lote']))
        await page.keyboard.press('Enter')
        await asyncio.sleep(2)
        await wait_for_overlay_to_disappear(page)

        print('Quantidade:', item['Quantidade'])
        await page.type('input[name="quantidade."]', str(item['Quantidade']))
        await page.keyboard.press('Enter')
        await asyncio.sleep(2)
        await wait_for_overlay_to_disappear(page)

        print(item['Confirmacao'])
        await page.type('input[name="confirma."]', 'S')
        await page.keyboard.press('Enter')
        await asyncio.sleep(2)

    await browser.close()

asyncio.get_event_loop().run_until_complete(atualiza_estoque())
