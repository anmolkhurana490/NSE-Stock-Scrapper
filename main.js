import * as cheerio from 'cheerio';
import axios from 'axios';
import XLSX from 'xlsx';
import nodeCron from 'node-cron';

// import saveDataToSheet from './saveSheets.js';
// const sheetId = '1iHqOA-36AyBcF3an-8d3Zb0pOLpLA_njYTPNLB6Q47A'; // Replace with your Google Sheet ID

const companies = [
    {
        name: 'RELIANCE',
        link: 'https://www.moneycontrol.com/india/stockpricequote/oil-gas/relianceindustries/RI'
    },
    {
        name: 'INFY',
        link: 'https://www.moneycontrol.com/india/stockpricequote/computers-software/infosys/IT'
    },
    {
        name: 'TCS',
        link: 'https://www.moneycontrol.com/india/stockpricequote/computers-software/tataconsultancyservices/IT'
    },
    {
        name: 'HDFCBANK',
        link: 'https://www.moneycontrol.com/india/stockpricequote/banks-private-sector/hdfcbank/HDF01'
    },
    {
        name: 'ICICIBANK',
        link: 'https://www.moneycontrol.com/india/stockpricequote/banks-private-sector/icicibank/ICI02'
    },
    {
        name: 'HINDUNILVR',
        link: 'https://www.moneycontrol.com/india/stockpricequote/fast-moving-consumer-goods-fmcg/hindustan-unilever-ltd/HU'
    },
    {
        name: 'SBIN',
        link: 'https://www.moneycontrol.com/india/stockpricequote/banks-public-sector/statebankofindia/SBI'
    },
    {
        name: 'KOTAKBANK',
        link: 'https://www.moneycontrol.com/india/stockpricequote/banks-private-sector/kotak-mahindra-bank-ltd/KMB'
    },
    {
        name: 'ITC',
        link: 'https://www.moneycontrol.com/india/stockpricequote/fast-moving-consumer-goods-fmcg/itc/ITC'
    },
    {
        name: 'LT',
        link: 'https://www.moneycontrol.com/india/stockpricequote/construction-engineering/larsentoubro/LT'
    },
]

const fetchCompanyData = async (company) => {
    try {
        const { data } = await axios.get(company.link)
        const $ = cheerio.load(data)
        const swot = $('div.swot_feature');

        const fetchData = {
            name: company.name,
            strength: swot.find('#swot_ls strong').text().match(/\d+/)[0],
            weakness: swot.find('#swot_lw strong').text().match(/\d+/)[0],
            opportunity: swot.find('#swot_lo strong').text().match(/\d+/)[0],
            threats: swot.find('#swot_lt strong').text().match(/\d+/)[0],
            mc_score: $('#mcessential_div div.esbx').text().slice(0, 2)
        }

        return fetchData;
    }
    catch (e) {
        console.log(e.message)
        return null;
    }
}


// saveDataToSheet(extractedData, sheetId);

const saveToXLSX = (data, day) => {
    const ws = XLSX.utils.json_to_sheet(data);

    let wb;
    try {
        wb = XLSX.readFile('nse_stock_analysis.xlsx');
    } catch (err) {
        wb = XLSX.utils.book_new();
    }

    XLSX.utils.book_append_sheet(wb, ws, `Day ${day}`);
    XLSX.writeFile(wb, 'nse_stock_analysis.xlsx');

    console.log('Data saved to nse_stock_analysis.xlsx');
}

let day = 1;
// Schedule the job to run daily at midnight
nodeCron.schedule('0 0 * * *', async () => {
    console.log('Running stock analysis...', day);
    const extractedData = await Promise.all(companies.map((company) => fetchCompanyData(company)));
    saveToXLSX(extractedData, day);
    day++;
});