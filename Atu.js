const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');


async function waitForOverlayToDisappear(page) {
    const overlaySelector = '#progress';    
    const overlayElemento = await page.$(overlaySelector);

    if(overlayElemento) {
        console.log('Overlay encontrado esperando desaparecer...');
        await page.waitForSelector(overlaySelector, {hidden: true});
        console.log('Overlay desapareceu.');
        
    } else {
       // console.log('Overlay não encontrado seguindo o fluxo');
    }
}

async function setaLogin(page){
    const usernameSelector = 'input[name="box1:usuario_tela."]';
    await page.waitForSelector(usernameSelector);
    await page.type(usernameSelector, 'MATHEUS');       
    
    const passwordSelector = 'input[name="box1:senha_tela."]';
    await page.waitForSelector(passwordSelector);
    await page.type(passwordSelector, 'MATTEUS');
    
    await new Promise(resolve => setTimeout(resolve, 2000));

    await page.keyboard.press('F9');
    }

    function GetDadosTabela(Caminho_arquivo) {
        const workbook = XLSX.readFile(localArquivo);
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
    
        const dados = XLSX.utils.sheet_to_json(sheet);
        return dados;
    }

async function Atualiza_estoque_deposito_1() {

    localArquivo = '../Codigos.xlsx';
    GetDadosTabela(localArquivo)
    
    const dados = GetDadosTabela(localArquivo);
    console.log(dados);

    const opcoesInicio = { headless: false, args: ['--start-maximized'] };
    
    console.log('Iniciando Puppeteer com as opções:', opcoesInicio);
    const browser = await puppeteer.launch(opcoesInicio);
    const page = await browser.newPage();

    // Ajustando o site com o tamanho da tela
    await page.setViewport({width: 1370, height: 1024});

    fileURL =  `https://divero.systextil.com.br/systextil/`;
    await page.goto(fileURL);

    await new Promise(resolve => setTimeout(resolve, 2000));
    
    await setaLogin(page);
    
    await waitForOverlayToDisappear(page);

    await new Promise(resolve => setTimeout(resolve, 2000));       
    
    await page.keyboard.type('estq_f010');
    
    page.keyboard.press('Enter');
    await new Promise(resolve => setTimeout(resolve, 4000));

    try {
        
        const codigo_transacao = 'input[name="codigo_transacao."]';
        const deposito_saida = 'input[name="deposito_saida."]';
        const codigo_barras = 'input[name="codigo_barras."]';
        let codigos_alerta = new Array();

        // page.on('dialog', async dialog => {
        //     console.log(dialog.message());                
            // console.log('dialogo aceito');
        //     await dialog.accept();
        // });
        await page.type(codigo_transacao, '003');
        console.log('inserido valor 003');
        await new Promise(resolve => setTimeout(resolve, 2000));
        page.keyboard.press('Enter');
        console.log('pressionado Enter');
        waitForOverlayToDisappear(page);
        
        console.log('inserido valor 001');
        await page.type(deposito_saida, '001');
        await new Promise(resolve => setTimeout(resolve, 2000));
        page.keyboard.press('Enter');
        console.log('pressionado Enter');
        waitForOverlayToDisappear(page);
        
        page.keyboard.press('F2');
        console.log('pressionado F2');
        await new Promise(resolve => setTimeout(resolve, 2000));
        page.keyboard.press('F2');
        console.log('pressionado F2');


        for( let i = 0; i < dados.length; i++){
            const item = dados[i];

            // page.on('dialog', async dialog => {
            //     console.log(dialog.message());                
            //     console.log('dialogo aceito');
            //     codigos_alerta.push(String(item)); 
            //     await dialog.accept();
            // });
            
            waitForOverlayToDisappear(page);
            await new Promise(resolve => setTimeout(resolve, 2000));
            
            console.log('--------------------------------------------------');
            console.log(item);
            await page.type(codigo_barras, String(item));
            page.keyboard.press('F9');
            await new Promise(resolve => setTimeout(resolve, 2000));
            waitForOverlayToDisappear(page);

        }

        for (let i = 0; i < codigos_alerta.length; i++) {
            console.log(codigos_alerta [i]);            
        }

        // process.exit(0);

    } catch (error) {
        console.error('Erro ao inserir o valor ', error.message);
    }

}

Atualiza_estoque_deposito_1().catch(console.error);    