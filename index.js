const puppet = require('puppeteer');
const xlsx = require('xlsx');



(async () => {

    const browser = await puppet.launch({
        args:[
        '--disable-web-security'
        ],
        headless: true,
        executablePath: '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'
    });

    // create new page
    const page = await browser.newPage();
    await page.goto('https://www.optune.com/hcp/treatment-centers/list')

    let links = await page.$$eval('.textLink', a => a.map(text => text.href))
    
    let data = [];
    for(var i=0; i<links.length; i++){
        
        await page.goto(links[i].toString())
        
        let q = await page.evaluate(() => document.querySelectorAll('#treatmentCenterName')[0].textContent);
        let r = await page.evaluate(() => document.querySelectorAll('#treatmentCenterAddress1')[0].textContent);
        let s = await page.evaluate(() => document.querySelectorAll('#treatmentCenterCity')[0].textContent);
        let t = await page.evaluate(() => document.querySelectorAll('#treatmentCenterState')[0].textContent);
        let u = await page.evaluate(() => document.querySelectorAll('#treatmentCenterZipCode')[0].textContent);
        let v = await page.$$eval('#treatmentCenterPhysicianTable td', v => v.map(text => text.textContent.trim().replace(/(\r\n|\n|\r)/gm, " ")
        ))
        let w = await page.$$eval('#treatmentCenterWebsiteList td', v => v.map(text => text.textContent.trim().replace(/(\r\n|\n|\r)/gm, " ")
        ))
        let x = await page.$$eval('#treatmentCenterEmailList td', v => v.map(text => text.textContent.trim().replace(/(\r\n|\n|\r)/gm, " ")
        ))
        let y = await page.$$eval('#treatmentCenterPhoneNumberList td', v => v.map(text => text.textContent.trim().replace(/(\r\n|\n|\r)/gm, " ")
        ))

        data.push({
            'Center':q,
            'Address': r,
            'City': s,
            'State': t,
            'Zip': u,
            'Physician(s)': v.join(' \n'),
            'Website(s)': w.join(' \n'),
            'Email(s)': x.join(' \n'),
            'Phone Number(s)':y.join(' \n')
        })
 
    }

    const wb = xlsx.utils.book_new();

 for (const x in data){
     if(x==0) {
        var ws = xlsx.utils.json_to_sheet([data[x]]);
        
     } else if(x > 0  && x != data.length-1){
        if(data[x].State === data[x-1].State ){
            console.log('old')
            xlsx.utils.sheet_add_json(ws, [data[x]], {skipHeader: true, origin: -1})
        } else {
            console.log('new')
            xlsx.utils.book_append_sheet(wb, ws, data[x-1].State);
            ws = xlsx.utils.json_to_sheet([data[x]]);
        }
     } else if(x == data.length-1){
        xlsx.utils.sheet_add_json(ws, [data[x]], {skipHeader: true, origin: -1})
        xlsx.utils.book_append_sheet(wb, ws, data[x].State);
     }      
}
    xlsx.writeFile(wb, "text.xlsx");

})();