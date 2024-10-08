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
       const tempo_ok = tempo * 1000;
        await new Promise(resolve => setTimeout(resolve, tempo_ok));
        return;
    }

    async function setaLogin(page){
        const usernameSelector = 'input[name="box1:usuario_tela."]';
        await page.waitForSelector(usernameSelector);
        await page.type(usernameSelector, 'MATHEUS');       
        
        const passwordSelector = 'input[name="box1:senha_tela."]';
        await page.waitForSelector(passwordSelector);
        await page.type(passwordSelector, 'matteus');
        
       await wait_a_moment(1);

        await page.keyboard.press('F9');
        }

    function GetDadosTabela(Caminho_arquivo) {
        const workbook = XLSX.readFile(Caminho_arquivo);
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
    
        const dados = XLSX.utils.sheet_to_json(sheet);
        return dados;
    }

    async function salvaCodigoForaEstoque(item) {
        codigosForaEstoque.push(item);
        return;
    }
    
    async function salvaCodigoEstoqueZero(item) {
        codigosEstoqueZero.push(item);
        return;
    }

    async function salvaCodigoFormatoInvalido(item) {
        codigosFormatoInvalido.push(item);
        return;
    }
   
    async function salvaCodigosErrosNaoIdentificados(item) {
        codigosFormatoInvalido.push(item);
        return;
    }

    async function AtualizaEstoque() {

        localArquivo = '../Codigos.xlsx';
        
        const dados = GetDadosTabela(localArquivo);
        // console.log(dados);

        const opcoesInicio = { headless: false, args: ['--start-maximized'] };
         
        console.log('Iniciando Puppeteer com as opções:', opcoesInicio);
        const browser = await puppeteer.launch(opcoesInicio);
        const page = await browser.newPage();

        // Ajustando o site com o tamanho da tela
        await page.setViewport({width: 1370, height: 1024});

        fileURL =  `https://divero.systextil.com.br/systextil/`;
        await page.goto(fileURL);
    
        let codigo_atual = null;

        page.on('dialog', async dialog => {
            console.log(dialog.message());
            message = dialog.message();

            switch(message) {

                case 'ATENÇÃO! A quantidade solicitada ultrapassou a quantidade disponível para o produto. Estoque Disponível: 0.0':
                    salvaCodigoEstoqueZero(codigo_atual);
                    dialog.accept();
                    return;
                
                case 'ATENÇÃO! TAG está com situação 4 faturada ou fora de estoque  não pode ser movimentada.':
                    salvaCodigoForaEstoque(codigo_atual);
                    dialog.accept();
                    return;
                
                case 'ATENÇÃO! Código de barras inválido.':
                    console.log('Ocorreu um erro. Item com problema: ', codigo_atual);
                    salvaCodigoFormatoInvalido(codigo_atual);
                    dialog.accept();
                    return; 

                case 'O número de conexões simultâneas excedeu o limite do ambiente Systêxtil. Favor entrar em contato com seu time de tecnologia ou diretamente com a Systêxtil pelo e-mail comercial@systextil.com.br.':
                    console.log('Sem usuários disponíveis no momento, fechando o programa.  ');
                    dialog.accept();
                    process.exit(0);
                    break; 

                default: 
                    salvaCodigosErrosNaoIdentificados(codigo_atual);
                    dialog.accept();
                    return;

            };                
            // console.log('dialogo aceito');
            // await dialog.accept();                                                                               
        });

        await wait_a_moment(2);        
        await setaLogin(page);
        
        await waitForOverlayToDisappear(page);
        
        
        await wait_a_moment(4);        
       
        await page.keyboard.type('estq_f010');
       
        page.keyboard.press('Enter');
        
        await wait_a_moment(2);        
        page.keyboard.press('Enter');
        
        try {
          
            const campo_transacao = 'input[name="codigo_transacao."]';
            const campo_deposito = 'input[name="deposito_saida."]';
            const campo_documento = 'input[name="nr_documento."]';
            const campo_codigo_barras = 'input[name="codigo_barras."]';

            waitForOverlayToDisappear(page);
            await wait_a_moment(5);

            // console.log('--------------------------------------------------');
            // console.log(item1.Codigos);
            await page.type(campo_transacao, '003');
            console.log('Valor 003 inserido');
            page.keyboard.press('Enter');
    
            waitForOverlayToDisappear(page);
            await wait_a_moment(5);

            // console.log('--------------------------------------------------');
            // console.log(item2.Codigos);
            await page.type(campo_deposito, '010');
            console.log('Valor 010 inserido');
            page.keyboard.press('Enter');
            await wait_a_moment(2);        
            waitForOverlayToDisappear(page);
            
            page.keyboard.press('F2');
            await wait_a_moment(2);         
            
            await page.type(campo_documento, '5');
            page.keyboard.press('F9');
            
            await wait_a_moment(2);         

            page.keyboard.press('F2');
            await wait_a_moment(2);           

            codigosEstoqueZero = new Array();
            codigosFormatoInvalido = new Array();
            codigosErrosNaoIdentificados = new Array();
            codigosForaEstoque = new Array();

            // Criando a variével para salvar o código de barras atual do loop  

            for( let i = 0; i < dados.length; i++){
                const item = dados[i];
                codigo_atual = item.Codigos;
                
                waitForOverlayToDisappear(page);
                await wait_a_moment(2);
                
                console.log('--------------------------------------------------');
                await page.type(campo_codigo_barras, String(codigo_atual));
                console.log(codigo_atual);
                console.log('Controle: ', i);
                page.keyboard.press('Enter');
                await wait_a_moment(2);
                waitForOverlayToDisappear(page);                
            }

            console.log('Iniciando impressão dos códigos que geraram alerta: ');
            console.log(' ');
            
            console.log('Códigos formato inválido : ----------------------------------');
            
            for (let i = 0; i < codigosFormatoInvalido.length; i++) {
                const codigo = codigosFormatoInvalido[i];
                
                console.log(codigo);
            }
            
            console.log(' ');
            console.log('Códigos com estoque zerado: ----------------------------------');
            
            for (let i = 0; i < codigosEstoqueZero.length; i++) {
                const codigo = codigosEstoqueZero[i];
                console.log(codigo);
            }
           
            console.log(' ');
            console.log('Códigos fora do estoque ou com situação 4 faturada: ----------------------------------');
            
            for (let i = 0; i < codigosForaEstoque.length; i++) {
                const codigo = codigosForaEstoque[i];
                console.log(codigo);
            }
            
            console.log(' ');
            console.log('Códigos fora do estoque ou com situação 4 faturada: ----------------------------------');
            
            for (let i = 0; i < codigosErrosNaoIdentificados.length; i++) {
                const codigo = codigosErrosNaoIdentificados[i];
                console.log(codigo);
            }

            console.log('Códigos com erro impressos.');

            // process.exit(0);
        } catch (error) {
            console.error('Erro ao inserir o valor ', error.message);
        }

    }

     AtualizaEstoque().catch(console.error);    
