const puppeteer = require('puppeteer-extra');
const StealthPlugin = require('puppeteer-extra-plugin-stealth')
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx')


const Tab = require('./modules/Pup_modules/tab');
const Delay = require('./modules/Pup_modules/delay');

function lerDuasPrimeirasColunasDoExcel(caminhoArquivo) {
    const workbook = XLSX.readFile(caminhoArquivo);
    const sheetName = workbook.SheetNames[0]; // Assume que os dados estão na primeira aba
    const sheet = workbook.Sheets[sheetName];

    const dados = XLSX.utils.sheet_to_json(sheet, { header: 1 }); // Converte a aba em um array de arrays
    return dados.map(row => [row[0], row[1]]); // Retorna as duas primeiras colunas de cada linha
}
function salvarDadosExcel(dados, cnpj) {
    const headers = ['CNPJ', 'Placa', 'Status'];
    let worksheet = XLSX.utils.aoa_to_sheet([headers]);
    dados.forEach(dado => {
        const row = Object.values(dado);
        XLSX.utils.sheet_add_aoa(worksheet, [row], { origin: -1 });
    });
    let workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Dados");
    XLSX.writeFile(workbook, `./${cnpj} Retorno.xlsx`);
}
function getFileCount(directory) {
    return fs.readdirSync(directory).length;
}
async function waitForNewFile(downloadPath, initialCount) {
    return new Promise(resolve => {
        let interval = setInterval(() => {
            const currentCount = getFileCount(downloadPath);
            if (currentCount > initialCount) {
                clearInterval(interval);
                resolve();
            }
        }, 1000); // Verifica a cada 1 segundo
    });
}

async function runAutomation() {

    const linkGov = 'https://www.gov.br/pt-br/servicos/consultar-online-suas-infracoes-de-transito'
    const downloadPath = path.resolve(__dirname, 'downloads');

    puppeteer.use(StealthPlugin())
    // Inicia Chrome com atributos para poder ter a alteração do lugar de download
    const browser = await puppeteer.launch({
        headless: false,
        // executablePath: 'C:\Program Files\Mozilla Firefox\firefox.exe',
        protocolTimeout: 90000,
        args: [
            '--no-sandbox',
            '--disable-setuid-sandbox',
            `--disable-dev-shm-usage`,
            `--disable-accelerated-2d-canvas`,
            `--disable-gpu`
        ],
    });

    const page = await browser.newPage();

    // Alteração do lugar de download
    const client = await page.target().createCDPSession()
    await client.send('Page.setDownloadBehavior', {
        behavior: 'allow',
        downloadPath: downloadPath,
    })

    await page.goto(linkGov)

    // Login GOV 
    const shadowHost = await page.waitForSelector('#barra-sso');
    const shadowRoot = await page.evaluateHandle(host => host.shadowRoot, shadowHost);
    const signInButton = await shadowRoot.$('#sso-status-bar');
    await signInButton.click();
    await page.waitForXPath('//*[@id="cert-digital"]/button')
    await Delay(2000)
    await page.click('xpath/' + '//*[@id="cert-digital"]/button')
    await Delay(5000)

    // Aceitando cookies
    await Delay(2000)
    await page.waitForXPath('//*[@id="content-core"]/div[2]/div[1]/a', { timeout: 150000 })
    const elements = await page.$x('//*[@id="content-core"]/div[2]/div[1]/a');
    if (elements.length > 0) {
        await elements[0].click();
        Delay(1000)
        await elements[0].click();
        Delay(1000)
        // Acessando "Consultar online suas infrações de trânsito"
        await page.click('xpath/' + '//html/body/div[5]/div/div/div/div/div[2]/button[3]')
        Delay(1000)
        await page.click('xpath/' + '//*[@id="content-core"]/div[2]/div[1]/a')
    } else {
        console.log('Elemento não encontrado');
    }

    // Aceitando cookies
    await page.waitForXPath('//*[@id="cookiebar"]/div[1]/div/div/div/div[2]/button[2]')
    await Delay(3000)
    await page.click('xpath/' + '//*[@id="cookiebar"]/div[1]/div/div/div/div[2]/button[2]')
    await page.waitForXPath('//*[@id="card-servicos"]/div/div/slide[4]/div/div')
    await Delay(2000)
    // Clica "Consultar Meus Veículos"
    await page.waitForSelector('#card-servicos > div > div > slide:nth-child(4) > div > div')
    await page.click('#card-servicos > div > div > slide:nth-child(4) > div > div')
    await Delay(2000)
    // Escolhendo CNPJ
    await page.waitForXPath('/html/body/modal-container/div/div/div[2]/div[2]/ul/li[2]/a/span[2]')
    const tabelaDados = lerDuasPrimeirasColunasDoExcel('./teste.xlsx')

    const cnpj = (tabelaDados[1][0]).toString()

    await Delay(10000)
    await page.click('xpath/' + '/html/body/modal-container/div/div/div[2]/div[2]/ul/li[2]/a/span[1]')
    await page.waitForSelector('.input-custom')
    await Delay(2000)
    const selector = 'select[formcontrolname="cnpjEntidade"]';
    const cnpjValue = '0' + cnpj.toString(); // Substitua com o valor real de cnpj

    console.log(cnpjValue)

    await page.waitForSelector(selector, { visible: true });
    await page.select(selector, cnpjValue);
    await Delay(1000)
    await page.click('xpath/' + '/html/body/modal-container/div/div/div[2]/div[5]/div/button[1]')
    await Delay(2000)
    // Clicando Não Aderir ao SNE
    await page.waitForXPath('/html/body/modal-container/div/div/div[3]/div/button[2]')
    await Delay(2000)
    await page.click('xpath/' + '/html/body/modal-container/div/div/div[3]/div/button[2]')

    let Dados = []
    loop: for (let i = 1; i < tabelaDados.length; i++) {
        let placaVeiculo = (tabelaDados[i][1]).toString()

        // Digita Placa
        await page.waitForXPath('/html/body/app-root/form/br-main-layout/div/div/main/app-veiculo/app-veiculos-list/div/div/div/form/br-tab-set/div/nav/br-tab/div/div[2]/div[1]/br-input/div/div/input', { visible: true })
        await Delay(10000)

        await page.click('xpath/' + '/html/body/app-root/form/br-main-layout/div/div/main/app-veiculo/app-veiculos-list/div/div/div/form/br-tab-set/div/nav/br-tab/div/div[2]/div[1]/br-input/div/div/input')

        await page.keyboard.down('Control')
        await page.keyboard.down('A')
        await page.keyboard.up('Control')
        await page.keyboard.up('A')
        await page.keyboard.press('Backspace')

        await Delay(2000)
        await page.keyboard.type(placaVeiculo)
        await page.click('xpath/' + '/html/body/app-root/form/br-main-layout/div/div/main/app-veiculo/app-veiculos-list/div/div/div/form/br-tab-set/div/nav/br-tab/div/div[2]/div[3]/button[1]')
        await Delay(5000)

        const naoTemPlaca = await page.$('.description');

        if (naoTemPlaca) {
            console.log("Não tem placa")
            console.log('-------------------------------------------------------------')
            Dados.push({ 'CNPJ': cnpj , 'Placa': placaVeiculo, 'Status': 'Inválida'})
            await page.reload({ waitUntil: 'networkidle0' });
            continue loop
        }
        await page.waitForXPath('/html/body/app-root/form/br-main-layout/div/div/main/app-veiculo/app-veiculos-list/div/div/div/form/br-tab-set/div/nav/br-tab/div/div[3]/div/div')
        await Delay(2000)
        await page.click('xpath/' + '/html/body/app-root/form/br-main-layout/div/div/main/app-veiculo/app-veiculos-list/div/div/div/form/br-tab-set/div/nav/br-tab/div/div[3]/div/div')
        await page.waitForXPath('//*[@id="header-small"]/br-tab-set/div/nav/br-tab[1]/div/app-veiculo-dados/div/div/table/tbody/tr[2]/td/a', { visible: true })
        await Delay(5000)
        await page.click('xpath/' + '//*[@id="header-small"]/br-tab-set/div/nav/br-tab[1]/div/app-veiculo-dados/div/div/table/tbody/tr[2]/td/a')


        let initialFileCount = getFileCount(downloadPath);

        // Aguarda a conclusão do download
        await waitForNewFile(downloadPath, initialFileCount);
        initialFileCount = getFileCount(downloadPath);

        await Delay(5000)
        Dados.push({ 'CNPJ': cnpj , 'Placa': placaVeiculo, 'Status': 'Baixada'})

        await page.goBack()
        await page.waitForNavigation()

    }

    salvarDadosExcel(Dados, cnpj)
    await browser.close()

}

runAutomation()