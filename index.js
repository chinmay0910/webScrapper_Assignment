const puppeteer = require('puppeteer');
const excel = require('exceljs');
const stringSimilarity = require("string-similarity");

// Load Excel file
(async () => {
    // set puppeteer
    const browser = await puppeteer.launch({headless: false});
    const page = await browser.newPage();

    // locating and loading Excel file
    const workbook = new excel.Workbook();
    await workbook.xlsx.readFile('Input.xlsx');
    const worksheet = workbook.getWorksheet('Sheet1');
    // Iterating through Rows of Excel
    for (let i = 2; i <= worksheet.actualRowCount; i++) {
        const productTitleCell = worksheet.getCell(`B${i}`).value;
        const isbnCell = worksheet.getCell(`C${i}`).value.toString();
        const foundCell = worksheet.getCell(`D${i}`);
        let pageUrlCell = worksheet.getCell(`E${i}`);
        let priceCell = worksheet.getCell(`F${i}`);
        let authorCell = worksheet.getCell(`G${i}`);
        let publisherCell = worksheet.getCell(`H${i}`);
        let inStock = worksheet.getCell(`I${i}`);
        console.log(isbnCell);

        // navigate to snapdeal
        await page.goto("https://www.snapdeal.com/");
        await page.type("#inputValEnter", isbnCell);
        await page.keyboard.down("Enter");

        // Updating the Found Column
        // Wait for search results
        try {
            await page.waitForSelector('.search-li', {timeout: 4000});
        } catch (error) {
            foundCell.value = 'No';
            inStock.value = 'N.A';
            pageUrlCell.value = 'N.A'
            priceCell.value = 'N.A'
            authorCell.value = 'N.A'
            publisherCell.value = 'N.A'
            console.error("Timeout waiting for search results");
            continue; // Moving on to the next book
        }

        const searchResults = await page.$$('.search-li');
        if (searchResults.length == 0) {
            foundCell.value = 'No';
            inStock.value = 'N.A';
        } else {
            let productTitle_Web = await page.$eval('.product-title', tx => (((tx.innerText).replace(/ *\([^)]*\) */g, "")).split('-')[0]).split(':')[0].split('&')[0]);
            // let match = stringSimilarity.compareTwoStrings(productTitleCell,productTitle_Web);
            let match = similarity(productTitleCell, productTitle_Web);
            // console.log(productTitle_Web, productTitleCell, match);
            let price, publisher, author, pageUrl;

            if (match >= 0.9 || checkMatch(productTitle_Web, productTitleCell)) { // checkMatch is put to avoid some bugs
                try {
                    await page.waitForSelector('.sort-selected');
                    await page.click('.sort-selected');
                    author = await page.$eval('.product-author-name', tx => tx.title);
                    await page.waitForSelector('.search-li');
                    await page.click('.search-li:nth-child(3)');
                    await page.waitForNavigation();
                    await page.waitForSelector('.favDp a:nth-child(1)');
                    pageUrl = await page.$eval('.favDp a:nth-child(1)', li => li.href);
                    await page.goto(pageUrl);
                    await page.waitForSelector('.payBlkBig');
                    price = await page.$eval('.payBlkBig', tx => tx.innerText);
                    await page.waitForSelector('.p-keyfeatures li:nth-child(3)');
                    publisher = await page.$eval('.p-keyfeatures li:nth-child(3)', tx => (tx.innerText).replace("Publisher:", ""));

                    pageUrlCell.value = pageUrl;
                    priceCell.value = price;
                    authorCell.value = author;
                    publisherCell.value = publisher;
                    inStock.value = 'Yes';
                    foundCell.value = 'Yes';
                } catch (error) {
                    console.error("Error: While Clicking for whole Details");
                    continue;
                }

            } else {
                foundCell.value = 'No';
                inStock.value = 'No';
            }
        }


    }
    console.log("Done! :)");
    await workbook.xlsx.writeFile('output.xlsx');
    browser.close()
})();

// Function for checking if it includes in the string or not
function checkMatch(str1, str2) {
    let txt1 = str1.toLowerCase();
    let txt2 = str2.toLowerCase();
    let isMatch = txt1.includes(txt2);
    return isMatch;
}

// Function to calculate similarity between two strings
function similarity(str1, str2) {
    const longer = str1.length > str2.length ? str1 : str2;
    const shorter = str1.length > str2.length ? str2 : str1;
    const longerLength = longer.length;

    if (longerLength === 0) {
        return 1.0;
    }

    return (longerLength - editDistance(longer, shorter)) / parseFloat(longerLength);
}

// Function to calculate edit distance between two strings
function editDistance(str1, str2) {
    str1 = str1.toLowerCase();
    str2 = str2.toLowerCase();

    const costs = new Array();
    for (let i = 0; i <= str1.length; i++) {
        let lastValue = i;
        for (let j = 0; j <= str2.length; j++) {
            if (i === 0) {
                costs[j] = j;
            } else {
                if (j > 0) {
                    let newValue = costs[j - 1];
                    if (str1.charAt(i - 1) !== str2.charAt(j - 1)) {
                        newValue = Math.min(Math.min(newValue, lastValue), costs[j]) + 1;
                    }
                    costs[j - 1] = lastValue;
                    lastValue = newValue;
                }
            }
        }
        if (i > 0) {
            costs[str2.length] = lastValue;
        }
    }
    return costs[str2.length];
}