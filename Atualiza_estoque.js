const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

     // Função para aguardar o overlay de "aguarde" desaparecer
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
        await page.type(passwordSelector, 'matteus');
        
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

    async function AtualizaEstoque() {

        localArquivo = '../planilha.xlsx';
        GetDadosTabela(localArquivo)
        
        const dados = GetDadosTabela(localArquivo);
        console.log(dados);

        const opcoesInicio = { headless: false, args: ['--start-maximized'] };
         
        
        console.log('Iniciando Puppeteer com as opções:', opcoesInicio);
        const browser = await puppeteer.launch(opcoesInicio);
        const page = await browser.newPage();

        // Ajustando o site com o tamanho da tela
        await page.setViewport({width: 1370, height: 1024});

        // fileURL =  `file://${path.join(__dirname, 'index.html')}`;
        fileURL =  `https://divero.systextil.com.br/systextil/`;
        await page.goto(fileURL);

        await new Promise(resolve => setTimeout(resolve, 2000));
        
        await setaLogin(page);
        
        await waitForOverlayToDisappear(page);

        await new Promise(resolve => setTimeout(resolve, 2000));       
       
        await page.keyboard.type('estq_f950');
       
        page.keyboard.press('Enter');
       
        await new Promise(resolve => setTimeout(resolve, 2000));

        try {
          
            const campo_nivel = 'input[name="nivel."]';
            const campo_grupo = 'input[name="grupo."]';
            const campo_subgrupo = 'input[name="subgrupo."]';
            const campo_item = 'input[name="item."]';
            const campo_deposito = 'input[name="cod_deposito."]';
            const campo_lote = 'input[name="lote."]';
            const campo_quantidade = 'input[name="quantidade."]';
            const campo_confirmacao = 'input[name="confirma."]';

            page.on('dialog', async dialog => {
                console.log(dialog.message());
                console.log('dialogo aceito');
                await dialog.accept();
            });

            for( let i = 0; i < dados.length; i++){
                const item = dados[i];
                
                waitForOverlayToDisappear(page);
                await new Promise(resolve => setTimeout(resolve, 2000));
                

                console.log('--------------------------------------------------');
                //console.log(item.Nivel);
                await page.type(campo_nivel, String(item.Nivel));
                page.keyboard.press('Enter');
                await new Promise(resolve => setTimeout(resolve, 2000));
                waitForOverlayToDisappear(page);
                
                console.log(item.Grupo);
                await page.type(campo_grupo, String(item.Grupo));
                page.keyboard.press('Enter');
                await new Promise(resolve => setTimeout(resolve, 2000));
                waitForOverlayToDisappear(page);
               
                //console.log(item.Subgrupo);
                await page.type(campo_subgrupo, String(item.Subgrupo));
                page.keyboard.press('Enter');
                await new Promise(resolve => setTimeout(resolve, 2000));
                waitForOverlayToDisappear(page);

                console.log(item.Item);
                await page.type(campo_item, String(item.Item));
                page.keyboard.press('Enter');
                await new Promise(resolve => setTimeout(resolve, 2000));
                waitForOverlayToDisappear(page);

                //console.log(item.Deposito);
                await page.type(campo_deposito, String(item.Deposito));
                page.keyboard.press('Enter');
                await new Promise(resolve => setTimeout(resolve, 2000));
                waitForOverlayToDisappear(page);

                //console.log(item.Lote);
                await page.type(campo_lote, String(item.Lote));
                page.keyboard.press('Enter');
                await new Promise(resolve => setTimeout(resolve, 2000));
                waitForOverlayToDisappear(page);

                console.log('Quantidade: ', item.Quantidade);
                await page.type(campo_quantidade, String(item.Quantidade));
                page.keyboard.press('Enter');
                await new Promise(resolve => setTimeout(resolve, 2000));
                waitForOverlayToDisappear(page);

               // console.log(item.Confirmacao);
                await page.type(campo_confirmacao, 'S');
                page.keyboard.press('Enter');
                await new Promise(resolve => setTimeout(resolve, 2000));
            }

        } catch (error) {
            console.error('Erro ao inserir o valor ', error);
        }

    }

     AtualizaEstoque().catch(console.error);    
