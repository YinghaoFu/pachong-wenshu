'use strict'

const fs = require('fs');

const xlsx = require('node-xlsx').default;
const puppeteer = require('puppeteer');

(async () => {
  const browser = await puppeteer.launch({
    // headless: false, // 开启界面
    // defaultViewport: null,
    // slowMo: 80,
  });
  const page = await browser.newPage();
  // 文书网
  await page.goto('https://wenshu.court.gov.cn/');
  await page.waitForSelector('#_view_1540966814000 > div > div.search-wrapper.clearfix > div.advenced-search');
  // 高级检索
  await page.click('#_view_1540966814000 > div > div.search-wrapper.clearfix > div.advenced-search');
  const searchText = '村规民约';
  // 全文检索关键字
  await page.type('#qbValue', searchText);
  // 全文检索类型
  await page.click('#qbType');
  // 理由
  await page.click('#qwTypeUl > li:nth-child(6)');
  // // 案件类型
  // await page.click('#selectCon_other_ajlx');
  // 民事案件 
  // await page.click('#gjjs_ajlx > li:nth-child(4)');
  // // 行政案件
  // await page.click('#gjjs_ajlx > li:nth-child(5)');
  // // 文书类型
  // await page.click('#_view_1540966814000 > div > div.advencedWrapper > div.inputWrapper.clearfix > div:nth-child(9) > div > div > div');
  // // 判决书
  // await page.click('#gjjs_wslx > li:nth-child(3)');
  // 裁决书
  // await page.click('#gjjs_wslx > li.on');
  // 年份开始（2017-01-01）
  await page.type('#cprqStart', '2017-01-01');
  // 年份结束（2020-12-31）
  await page.type('#cprqEnd', '2020-12-31');
  // 当事人
  await page.type('#s17', '村民委员会');

  // 检索
  await page.click('#searchBtn');
  await page.waitForSelector('#_view_1545184311000 > div.left_7_3 > div > select');

  // 高级检索2
  await page.click('#_view_1545034775000 > div > div.search-wrapper.clearfix > div.advenced-search');

  // 全文检索关键字
  await page.type('#qbValue', '');
  for (let index = 0; index < searchText.length; index++) {
    await page.keyboard.press('Backspace');
  }

  await page.type('#qbValue', '处罚');

  // 全文检索类型
  await page.click('#qbType');
  // 事实
  await page.click('#qwTypeUl > li:nth-child(5)');


  // 检索
  await page.click('#searchBtn');
  await page.waitForSelector('#_view_1545184311000 > div.left_7_3 > div > select');
  // 页容量改为15
  await page.select('#_view_1545184311000 > div.left_7_3 > div > select', '15');
  await page.waitFor(500);
  let pageNum = 1;
  const data = [['案号', '标题', '案件类型', '当事人', '案由', 'pdf内容']];
  let i = 1;
  while (pageNum < 6) {
    pageNum++;
    const view = await page.$('#_view_1545184311000');
    const lists = await view.$$('.LM_list');
    for (const list of lists) {
      try {
        const href = await list.$('div.list_title.clearfix > h4 > a');
        let ah = await list.$('div.list_subtitle > span.ah');
        ah = await ah.evaluate(node => node.innerText);
        await href.click();
        await page.waitFor(500);
        const page2 = (await browser.pages())[2];

        let title = await page2.$('#_view_1541573883000 > div > div.PDF_box > div.PDF_title');
        title = title !== null ? await title.evaluate(node => node.innerText) : '';
        let ajlx = await page2.$('#_view_1541573889000 > div:nth-child(1) > div.right_fixed > div.gaiyao_box > div.gaiyao_center > ul > li:nth-child(1) > h4:nth-child(2) > a');
        ajlx = ajlx !== null ? await ajlx.evaluate(node => node.innerText) : '';
        let reason = await page2.$('#_view_1541573889000 > div:nth-child(1) > div.right_fixed > div.gaiyao_box > div.gaiyao_center > ul > li:nth-child(1) > h4:nth-child(3) > a');
        reason = reason !== null ? await reason.evaluate(node => node.innerText) : '';
        let client = await page2.$('#_view_1541573889000 > div:nth-child(1) > div.right_fixed > div.gaiyao_box > div.gaiyao_center > ul > li:nth-child(1) > h4:nth-child(6) > b');
        client = client !== null ? await client.evaluate(node => node.innerText) : '';
        let content = await page2.$('#_view_1541573883000 > div > div.PDF_box');
        content = content !== null ? await content.evaluate(node => node.innerText) : '';
        data.push([ah, title, ajlx, client, reason, content]);
        console.log(`${i++}:${ah}`);
        await page2.close();
      } catch (error) {
        console.log(i++);
        console.error('error:', error);
        continue;
      }
    }
    try {
      await page.click(`#_view_1545184311000 > div.left_7_3 > a:nth-child(${pageNum + 1})`);
    } catch (error) {
      console.error('error:', error);
      continue;
    }
    await page.waitFor(500);
  }
  await browser.close();
  const buffer = xlsx.build([
    {
      name: 'sheet1',
      data,
    }
  ]);
  fs.writeFileSync('文书.xlsx', buffer, { 'flag': 'w' });
})();