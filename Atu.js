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

    async function wait_a_moment(tempo) {
        tempo_ok = tempo * 1000;
        await new Promise(resolve => setTimeout(resolve, tempo_ok));
    }

    async function setaLogin(page){
        const usernameSelector = 'input[name="box1:usuario_tela."]';
        await page.waitForSelector(usernameSelector);
        await page.type(usernameSelector, 'MATHEUS');       
        
        const passwordSelector = 'input[name="box1:senha_tela."]';
        await page.waitForSelector(passwordSelector);
        await page.type(passwordSelector, 'matteus');
        
       wait_a_moment(5);

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
        
        await new Promise(resolve => setTimeout(resolve, 2000));
        
        page.keyboard.press('Enter');

        try {
          
            const campo_transacao = 'input[name="codigo_transacao."]';
            const campo_deposito = 'input[name="deposito_saida."]';
            const campo_codigo_barras = 'input[name="codigo_barras."]';
            // const campo_item = 'input[name="item."]';
            // const campo_deposito = 'input[name="cod_deposito."]';
            // const campo_lote = 'input[name="lote."]';
            // const campo_quantidade = 'input[name="quantidade."]';
            // const campo_confirmacao = 'input[name="confirma."]';

           
            const item1 = dados[0];
            const item2 = dados[1];

            waitForOverlayToDisappear(page);
            await new Promise(resolve => setTimeout(resolve, 5000));

            console.log('--------------------------------------------------');
            console.log(item1.Codigos);
            await page.type(campo_transacao, String(item1.Codigos));
            console.log('Valor inserido');
            page.keyboard.press('Enter');
    
            waitForOverlayToDisappear(page);
            await new Promise(resolve => setTimeout(resolve, 5000));

            console.log('--------------------------------------------------');
            console.log(item2.Codigos);
            await page.type(campo_deposito, String(item2.Codigos));
            console.log('Valor inserido');
            page.keyboard.press('Enter');
            await new Promise(resolve => setTimeout(resolve, 2000));
            waitForOverlayToDisappear(page);
            
            page.keyboard.press('F2');
            await new Promise(resolve => setTimeout(resolve, 2000));
            
            page.keyboard.press('F2');
            await new Promise(resolve => setTimeout(resolve, 2000));

            page.on('dialog', async dialog => {
                console.log(dialog.message());
                if (dialog.message() === 'ATENÇÃO! Produto não cadastrado.') {
                    console.log('Ocorreu um erro. Item com problema: ');
                    await dialog.accept();
                    // process.exit(0);
                    console.log('pulando para a próxima posição...');
                    await new Promise(resolve => setTimeout(resolve, 4000));       
                    page.keyboard.press('F8');
                    return;
                };                
                console.log('dialogo aceito');
                await dialog.accept();
            });
            
            codigosAlerta = new Array();

            for( let i = 2; i < dados.length; i++){
                const item = dados[i];

                page.on('dialog', async dialog => {
                    console.log(dialog.message());
                    if (dialog.message() === 'ATENÇÃO! Código de barras inválido.') {
                        console.log('Ocorreu um erro. Item com problema: ', item.Codigos);
                        codigosAlerta.push(item.Codigos);
                        await dialog.accept();

                        return;
                    };                
                    // console.log('dialogo aceito');
                    // await dialog.accept();
                });
                
                waitForOverlayToDisappear(page);
                await new Promise(resolve => setTimeout(resolve, 3000));
                
                console.log('--------------------------------------------------');
                console.log(item.Codigos);
                await page.type(campo_codigo_barras, String(item.Codigos));
                page.keyboard.press('Enter');
                await new Promise(resolve => setTimeout(resolve, 2000));
                waitForOverlayToDisappear(page);
                
            }

            for (let i = 0; i < codigosAlerta.length; i++) {
                const codigo = codigosAlerta[i];
                
                console.log(codigo);
            }

            console.log('Códigos com erro imopressos.');

            // process.exit(0);
        } catch (error) {
            console.error('Erro ao inserir o valor ', error.message);
        }

    }

     AtualizaEstoque().catch(console.error);    
