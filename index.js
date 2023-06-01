const puppeteer = require('puppeteer');
const xlsx = require('xlsx');

async function run() {
    const browser = await puppeteer.launch({
        headless: false
    });
    const page = await browser.newPage();

    await page.goto('https://www2.correios.com.br/enderecador/encomendas/default.cfm');

    const workbook = xlsx.readFile('Planilha.xlsx');
    const worksheet = workbook.Sheets[workbook.SheetNames[workbook.SheetNames.length - 1]];
    const data = xlsx.utils.sheet_to_json(worksheet);

    try {
        const grupos = 4
        const numGrupos = Math.ceil(data.length / grupos)

        for (let grupo = 0; grupo < numGrupos; grupo++) {
            const inicio = grupo * grupos
            const fim = inicio + grupos
            const dadosGrupos = data.slice(inicio, fim)

            for (let i = 0, j = 1; i < dadosGrupos.length && j <= 4; i++, j++) {
                const nomeDest = dadosGrupos[i]['Nome'];
                const rua = dadosGrupos[i]['RUA']
                const numeroDest = rua.match(/\d+/)
                var complementoDest = dadosGrupos[i]['COMPLEMENTO']
                const cepDest = dadosGrupos[i]['CEP']
                const tipo = dadosGrupos[i]['TIPO']

                if (tipo === 'Correios') {

                    var cep = `input[name="cep_${j}"]`
                    var name = `input[name="nome_${j}"]`
                    var numero = `input[name="numero_${j}"]`
                    var complemento = `input[name="complemento_${j}"]`
                    var cepDes = `input[name="desCep_${j}"]`
                    var nomeDes = `input[name="desNome_${j}"]`
                    var numeroDes = `input[name="desNumero_${j}"]`
                    var complementoDes = `input[name="desComplemento_${j}"]`
                    var gerarEtiqueta = '#btGerarEtiquetas'
                    var gerarAR = '#btGerarAR'

                    console.log(nomeDest)

                    if (cep != '20031-040') {
                        await page.type(cep, '20031-040')
                        await page.click(name)
                        await page.waitForSelector(name)
                        await page.type(name, 'Gestão de Ativos Stone')
                        await page.type(numero, '65')
                        await page.type(complemento, 'Torre 2, Sala 201')
                    }

                    await page.type(cepDes, String(cepDest))
                    await page.click(nomeDes)
                    await page.waitForSelector(nomeDes)
                    await page.type(nomeDes, nomeDest)
                    await page.type(numeroDes, numeroDest)

                    if (complementoDest != null) {
                        await page.type(complementoDes, String(complementoDest))
                    } else {
                        complementoDest = "Sem Complemento"
                        await page.type(complementoDes, complementoDest)
                    }
                }
            }
            for (i = 1; i < 5; i++) {
                var mp = `[name="mp_${i}"]`
                if (await page.$(mp)) {
                    await page.click(mp)
                }
            }

            await page.click(gerarEtiqueta)

            await browser.waitForTarget(target => target.url().includes('gerarEtiqueta'))
            const pages = await browser.pages()
            const novaPagina = pages[pages.length - 1]
            await novaPagina.close();

            await browser.waitForTarget(target => target.url().includes('gerarEtiqueta') && target.type() === 'page', { state: 'detached' });

            await page.click(gerarAR)

            await browser.waitForTarget(target => target.url().includes('gerarAR') && target.type() === 'page', { state: 'detached' });

            await page.goto('https://www2.correios.com.br/enderecador/encomendas/default.cfm');

        }
    } catch (error) {
        console.error('Ocorreu um erro:', error);
    } finally {
        await browser.close();
    }
}


run();
