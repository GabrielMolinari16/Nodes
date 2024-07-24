const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');




     // Função para aguardar o overlay de "aguarde" desaparecer
     async function waitForOverlayToDisappear(page) {
        const overlaySelector = '#overlay';
        const overlayElemento = await page.$(overlaySelector);

        if(overlayElemento) {
            console.log('Overlay encontrado esperando desaparecer...');
            await page.waitForSelector(overlaySelector, {hidden: true});
            console.log('Overlay desapareceu.');
            
        } else {
            console.log('Overlay não encontrado seguindo o fluxo s');
        }
    }

    async function setaLogin(page){
        const usernameSelector = 'input[name="box1:usuario_tela."]';
        await page.waitForSelector(usernameSelector);
        await page.type(usernameSelector, 'DIVERO01');       
        
        const passwordSelector = 'input[name="box1:senha_tela."]';
        await page.waitForSelector(passwordSelector);
        await page.type(passwordSelector, '1234');
        
        await new Promise(resolve => setTimeout(resolve, 2000));

        await page.keyboard.press('F9');
        }

    function GetDadosTabela(Caminho_arquivo) {
        const workbook = XLSX.readFile(Caminho_arquivo);
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets['Planilha1'];

        const dados = XLSX.utils.sheet_to_json(sheet);
        return dados;
    }


    async function AtualizaEstoque() {

        
        localArquivo = '../planilha.xlsx';
       /*  const data = GetDadosTabela(localArquivo); */

        const workbook = XLSX.readFile(localArquivo);
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];

        const dados = XLSX.utils.sheet_to_json(sheet);
        const headers = Object.keys(dados[0]);

        dados.forEach((row, rowIndex) => {
            console.log(`Linha ${rowIndex + 1}:`);
            headers.forEach((header) => {
              console.log(`  Coluna ${header}: ${row[header]}`);
            });
          });

    /*     const linha1Coluna1 = dados[0]['A1'];
        console.log('ate aqui tudo certo');
        console.log(linha1Coluna1);
 */
        /* const headers = Object.keys(dados[0]);
        console.log(`Linha 1, ${headers[0]}: ${dados[0][headers[0]]}`);
        console.log(`Linha 1, ${headers[1]}: ${dados[0][headers[1]]}`); */

        //const values = fs.readFileSync('../valoresTeste.txt','utf-8').split('\n');

       /*  const opcoesInicio = { headless: false, args: ['--start-maximized'] };
         
        
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
            const codigo_Estagio_Baixa =  'input[name="codigo_estagio_baixa."]';
            await page.waitForSelector(codigo_Estagio_Baixa);
            console.log('Campo código de baixa encontrado.');
        
            await new Promise(resolve => setTimeout(resolve, 2000));
            await page.type(codigo_Estagio_Baixa, '85');

        } catch (error) {
            console.error('Erro ao inserir o valor "85" ', error);
        }

        for (let i = 0; i <= 3; i++) {
            await new Promise(resolve => setTimeout(resolve, 2000));
            await waitForOverlayToDisappear(page);
            
            if (i == 2) {

            const nr_Operadores = 'input[name="nr_operadores."]';
            await page.waitForSelector(nr_Operadores);
            console.log('Campo número de operadores encontrado.');

            await new Promise(resolve => setTimeout(resolve, 2000));
            await page.type(nr_Operadores, '0,0');

            } else {
                console.log('Pressionando Enter', i);
                page.keyboard.press('Enter');          
            }
        }

        const deposito =  'input[name="codigo_deposito."]';
        await page.waitForSelector(deposito);
        console.log('Campo código do depósito encontrado.\n');
 
        await new Promise(resolve => setTimeout(resolve, 2000));
        await page.type(deposito, values[0]); 
        
        await waitForOverlayToDisappear(page);
        page.keyboard.press('F9');    
 
        await new Promise(resolve => setTimeout(resolve, 1000));

        const btnEstornarPecas = '[name="cmdEstornarPecas"]';
        await page.waitForSelector(btnEstornarPecas);
        await page.click(btnEstornarPecas);

        await waitForOverlayToDisappear(page);


        // Repetição para bipar todos os códigos automaticamente: ---------------------------------------
        
        for (let i = 1; i < values.length; i++) {
            console.log('Posição atual: ', i);

            const campo_Leitura = 'input[name="campo_leitura."]';
            await page.waitForSelector(campo_Leitura);
            console.log('Campo código do leitura encontrado.\n');
            
            await new Promise(resolve => setTimeout(resolve, 2000));
            await page.type(campo_Leitura, values[i]);    

            await waitForOverlayToDisappear(page);
            
            page.keyboard.press('Enter');

            page.on('dialog', async dialog => {
                console.log(dialog.message()); // Imprime a mensagem do diálogo
                console.log(values[i]); // Imprime o código da peça 
        
                // Aceita o diálogo
                await dialog.accept();
            });
            
        } */
    }

    AtualizaEstoque().catch(console.error);    