const PRODUCT_TAB = '网卖商品';
const FIELDS = ['price', 'purchaseTotal', 'packaging', 'shipping', 'rate', 'other'];
const HEADERS = ['商品 ID', '商品名称', 'SKU', '售价', '进货成本', '包装费', '卖家运费', '平台费率', '其他费用', '平台费', '单件总成本', '预估利润', '利润率', '照片文件 ID', '更新版本', '目标利润率', '建议售价 (RM)', '进货总金额 (RM)', '进货数量', '进货确认时间', '累计出货数量', '实际销售收入', '已记销售费用', '补货记录'];

function setup() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) throw new Error('请从目标 Google 表格的「扩展程序 → Apps Script」打开此项目。');
  const properties = PropertiesService.getScriptProperties();
  const original = properties.getProperty('SPREADSHEET_ID');
  if (original && original !== spreadsheet.getId()) throw new Error('此项目已连接另一份表格。');
  const sheet = spreadsheet.getSheetByName(PRODUCT_TAB) || spreadsheet.insertSheet(PRODUCT_TAB);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setBackground('#12213b').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.getRange('H2:H').setNumberFormat('0.00%');
    sheet.getRange('M2:M').setNumberFormat('0.00%');
    sheet.getRange('D2:G').setNumberFormat('0.00');
    sheet.getRange('I2:L').setNumberFormat('0.00');
    sheet.setColumnWidth(2, 180);
    sheet.hideColumns(14, 2);
  } else {
    const existing = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
    if (existing.slice(0, 15).join('|') !== HEADERS.slice(0, 15).join('|') ||
        existing.slice(15).some((value, index) => value && value !== HEADERS[index + 15])) {
      throw new Error('已有同名工作表，请更名后重新执行 setup。');
    }
    if (!existing[15] && !existing[16]) {
      const rows = productRows_(sheet);
      if (rows.length) sheet.getRange(2, 16, rows.length, 2).setValues(rows.map((row, index) =>
        row[0] ? [0.2, suggestionFormula_(index + 2)] : ['', '']));
      sheet.getRange(1, 16, 1, 2).setValues([HEADERS.slice(15, 17)]).setBackground('#12213b').setFontColor('#ffffff').setFontWeight('bold');
    }
  }
  sheet.getRange(1, 18, 1, 2).setValues([HEADERS.slice(17, 19)]).setBackground('#12213b').setFontColor('#ffffff').setFontWeight('bold');
  sheet.getRange(1, 20, 1, 4).setValues([HEADERS.slice(19, 23)]).setBackground('#12213b').setFontColor('#ffffff').setFontWeight('bold');
  ensureRestockHeader_(sheet);
  sheet.getRange('T2:T').setNumberFormat('yyyy-mm-dd hh:mm');
  sheet.getRange('V2:W').setNumberFormat('"RM "0.00');
  sheet.getRange('R2:R').setNumberFormat('"RM "0.00');
  sheet.getRange('S2:S').setNumberFormat('0');
  sheet.getRange('P2:P').setNumberFormat('0.00%');
  sheet.getRange('Q2:Q').setNumberFormat('"RM "0.00');
  sheet.getRange('D2:G').setNumberFormat('"RM "0.00');
  sheet.getRange('I2:L').setNumberFormat('"RM "0.00');
  if (!properties.getProperty('PHOTO_FOLDER_ID')) {
    const folder = DriveApp.createFolder('商品成本台账照片');
    properties.setProperty('PHOTO_FOLDER_ID', folder.getId());
  }
  properties.setProperty('SPREADSHEET_ID', spreadsheet.getId());
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('商品成本台账')
    .addMetaTag('viewport', 'width=device-width,initial-scale=1');
}

function shopSheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('请先在 Apps Script 编辑器运行 setup。');
  return SpreadsheetApp.openById(id).getSheetByName(PRODUCT_TAB);
}

function productRows_(sheet) {
  const length = sheet.getLastRow() - 1;
  return length > 0 ? sheet.getRange(2, 1, length, HEADERS.length).getValues() : [];
}

function product_(row) {
  const cents = value => Math.round(Number(value) * 100);
  return {
    id: row[0], name: row[1], sku: String(row[2]), price: row[3] === '' ? null : cents(row[3]), purchase: cents(row[4]),
    purchaseTotal: row[17] === '' ? cents(row[4]) : cents(row[17]), quantity: Number(row[18]) || 1,
    received: Boolean(row[19]), sold: Number(row[20]) || 0, revenue: cents(row[21]), sellingCosts: cents(row[22]),
    restocks: row[23] ? JSON.parse(row[23]) : [],
    packaging: cents(row[5]), shipping: cents(row[6]), rate: Math.round(Number(row[7]) * 10000),
    other: cents(row[8]), photo: Boolean(row[13]), revision: row[14], target: row[15] === '' ? 2000 : Math.round(Number(row[15]) * 10000),
  };
}

function listProducts() {
  return productRows_(shopSheet_()).filter(row => row[0] && row[1]).map(product_);
}

function photoForProduct(id) {
  const row = productRows_(shopSheet_()).find(row => row[0] === id);
  if (!row || !row[13]) return '';
  const blob = DriveApp.getFileById(row[13]).getBlob();
  return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
}

function saveProduct(input, photoData) {
  if (!input || !/^[0-9a-f-]{36}$/.test(input.id) ||
      typeof input.name !== 'string' || !input.name.trim() || input.name.length > 100 ||
      typeof input.sku !== 'string' || input.sku.length > 60 ||
      !Number.isInteger(input.quantity) || input.quantity < 1 || input.quantity > 999999999 ||
      !Number.isInteger(input.target) || input.target < 0 || input.target > 10000 ||
      FIELDS.some(key => !(key === 'price' && input.price === null) &&
        (!Number.isSafeInteger(input[key]) || input[key] < 0 || input[key] > (key === 'rate' ? 10000 : 99999999900)))) throw new Error('请检查商品名称和金额。');
  const photo = photoData ? /^data:(image\/(?:webp|png|jpeg));base64,([A-Za-z0-9+/=]+)$/.exec(photoData) : null;
  if (photoData && (!photo || photoData.length > 5 * 1024 * 1024)) throw new Error('照片格式无效或文件过大。');

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = shopSheet_();
    const rows = productRows_(sheet);
    const index = rows.findIndex(row => row[0] === input.id);
    const current = rows[index];
    if ((current ? current[14] : null) !== (input.revision || null)) {
      throw new Error('商品已更新，请关闭编辑后刷新。');
    }
    if (current && current[19]) {
      const recorded = product_(current);
      if (input.purchaseTotal !== recorded.purchaseTotal || input.quantity !== recorded.quantity) {
        throw new Error('进货已确认，请使用「补货」更新进货数量和金额。');
      }
    }
    const rowNumber = index >= 0 ? index + 2 : sheet.getLastRow() + 1;
    if (rowNumber > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), 1);
    const oldFile = current ? current[13] : '';
    let newFile = '';
    if (photo) {
      const folder = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('PHOTO_FOLDER_ID'));
      newFile = folder.createFile(Utilities.newBlob(Utilities.base64Decode(photo[2]), photo[1], input.id)).getId();
    }
    const file = newFile || (input.keepPhoto ? oldFile : '');
    const textCell = value => "'" + value;
    const row = [
      input.id, textCell(input.name.trim()), input.sku.trim() ? textCell(input.sku.trim()) : '', input.price === null ? '' : input.price / 100,
      '=IF(OR(R' + rowNumber + '="",S' + rowNumber + '<1),"",ROUND(R' + rowNumber + '/S' + rowNumber + ',2))',
      input.packaging / 100, input.shipping / 100, input.rate / 10000, input.other / 100,
      '=IF(D' + rowNumber + '=,,ROUND(D' + rowNumber + '*H' + rowNumber + ',2))',
      '=ROUND(SUM(E' + rowNumber + ':G' + rowNumber + ',I' + rowNumber + ':J' + rowNumber + '),2)',
      '=IF(D' + rowNumber + '=,,ROUND(D' + rowNumber + '-K' + rowNumber + ',2))',
      '=IF(D' + rowNumber + '=0,"",L' + rowNumber + '/D' + rowNumber + ')',
      file, Utilities.getUuid(), input.target / 10000, suggestionFormula_(rowNumber), input.purchaseTotal / 100, input.quantity,
      ...(current ? current.slice(19, 24) : ['', 0, 0, 0, '']),
    ];
    sheet.getRange(rowNumber, 1, 1, row.length).setValues([row]);
    SpreadsheetApp.flush();
    if (oldFile && oldFile !== file) trashPhoto_(oldFile);
    return product_(sheet.getRange(rowNumber, 1, 1, row.length).getValues()[0]);
  } finally {
    lock.releaseLock();
  }
}

function suggestionFormula_(row) {
  return '=IF(H' + row + '+P' + row + '>=1,"",MAX(0.01,ROUNDUP((SUM(E' + row + ':G' +
    row + ',I' + row + ')+IF(H' + row + '>0,0.005,0))/(1-H' + row + '-P' + row + '),2)))';
}

function ensureRestockHeader_(sheet) {
  const header = sheet.getRange(1, 24);
  const value = header.getValue();
  if (value === HEADERS[23]) return;
  if (value !== '') throw new Error('工作表的 X 列已有其他标题，无法自动启用补货记录。');
  header.setValue(HEADERS[23]).setBackground('#12213b').setFontColor('#ffffff').setFontWeight('bold');
  sheet.hideColumns(24);
}

function recordRestock(id, revision, restock) {
  if (!restock) throw new Error('请输入补货件数和金额。');
  return updateInventory_(id, revision, null, restock);
}

function recordInventory(id, revision, sale) {
  return updateInventory_(id, revision, sale, null);
}

function updateInventory_(id, revision, sale, restock) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = shopSheet_();
    const rows = productRows_(sheet);
    const index = rows.findIndex(row => row[0] === id);
    if (index < 0 || rows[index][14] !== revision) throw new Error('商品已更新，请刷新后再试。');
    const row = rows[index];
    if (restock) ensureRestockHeader_(sheet);
    const product = product_(row);
    if (restock) {
      if (!product.received) throw new Error('请先确认首次进货。');
      const {quantity, total} = restock;
      const newQuantity = product.quantity + quantity;
      const newTotal = product.purchaseTotal + total;
      if (!Number.isSafeInteger(quantity) || quantity < 1 || newQuantity > 999999999 ||
          !Number.isSafeInteger(total) || total < 0 || newTotal > 99999999900) {
        throw new Error('请检查补货件数和金额。');
      }
      const history = JSON.stringify([...product.restocks, {date: new Date().toISOString(), quantity, total}]);
      if (history.length > 49000) throw new Error('此商品的补货记录已达工作表容量。');
      row[17] = newTotal / 100;
      row[18] = newQuantity;
      row[23] = history;
    } else if (!sale) {
      if (product.received) return product;
      row[19] = new Date();
      // Convert legacy unit-only entries to a one-item batch when confirmed.
      if (row[17] === '') row[17] = product.purchaseTotal / 100;
      if (row[18] === '') row[18] = product.quantity;
    } else {
      if (!product.received) throw new Error('请先确认进货。');
      if (!Number.isInteger(sale.quantity) || sale.quantity < 1 || sale.quantity > product.quantity - product.sold) {
        throw new Error('出货数量必须大于零，且不能超过库存。');
      }
      const revenue = product.price === null ? sale.revenue : product.price * sale.quantity;
      if (!Number.isSafeInteger(revenue) || revenue < 0 || revenue > 99999999900 ||
          !Number.isSafeInteger(sale.packaging) || sale.packaging < 0 || sale.packaging > 99999999900) {
        throw new Error('请检查本次销售收入和包装费。');
      }
      const fee = Math.round(revenue * product.rate / 10000);
      const costs = fee + sale.packaging + sale.quantity * (product.shipping + product.other);
      if (!Number.isSafeInteger(costs) || costs > 99999999900 ||
          product.revenue + revenue > 99999999900 || product.sellingCosts + costs > 99999999900) {
        throw new Error('销售金额或费用超出可记录范围。');
      }
      row[20] = product.sold + sale.quantity;
      row[21] = (product.revenue + revenue) / 100;
      row[22] = (product.sellingCosts + costs) / 100;
    }
    row[14] = Utilities.getUuid();
    // Preserve the existing spreadsheet formulas between version and inventory fields.
    const values = [row[14], ...sheet.getRange(index + 2, 16, 1, 2).getValues()[0]];
    values[2] = sheet.getRange(index + 2, 17).getFormula() || values[2];
    values.push(...row.slice(17, 24));
    sheet.getRange(index + 2, 15, 1, values.length).setValues([values]);
    SpreadsheetApp.flush();
    return product_(sheet.getRange(index + 2, 1, 1, HEADERS.length).getValues()[0]);
  } finally {
    lock.releaseLock();
  }
}

function deleteProduct(id, revision) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = shopSheet_();
    const rows = productRows_(sheet);
    const index = rows.findIndex(row => row[0] === id);
    if (index < 0) return;
    if (rows[index][14] !== revision) throw new Error('商品已更新，请刷新后再试。');
    sheet.deleteRow(index + 2);
    SpreadsheetApp.flush();
    if (rows[index][13]) trashPhoto_(rows[index][13]);
  } finally {
    lock.releaseLock();
  }
}

function trashPhoto_(id) {
  try { DriveApp.getFileById(id).setTrashed(true); }
  catch (error) { console.warn('Could not move old photo to Trash: ' + error.message); }
}

function getCompanySettings() {
  const properties = PropertiesService.getScriptProperties();
  const name = properties.getProperty('COMPANY_NAME') || '';
  const logoId = properties.getProperty('COMPANY_LOGO_ID');
  if (!logoId) return {name, logo: ''};
  const blob = DriveApp.getFileById(logoId).getBlob();
  return {name, logo: 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes())};
}

function saveCompanySettings(name, logoData, removeLogo) {
  if (typeof name !== 'string' || name.length > 80) throw new Error('公司名称最多可输入 80 个字符。');
  const logo = logoData ? /^data:(image\/(?:webp|png|jpeg));base64,([A-Za-z0-9+/=]+)$/.exec(logoData) : null;
  if (logoData && (!logo || logoData.length > 1024 * 1024)) throw new Error('无法保存此 Logo，请选择较小的图片。');

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const properties = PropertiesService.getScriptProperties();
    const previous = properties.getProperty('COMPANY_LOGO_ID') || '';
    let next = removeLogo ? '' : previous;
    if (logo) {
      const folderId = properties.getProperty('PHOTO_FOLDER_ID');
      if (!folderId) throw new Error('请先在 Apps Script 编辑器运行 setup。');
      const folder = DriveApp.getFolderById(folderId);
      next = folder.createFile(Utilities.newBlob(Utilities.base64Decode(logo[2]), logo[1], '公司 Logo')).getId();
    }
    properties.setProperties({COMPANY_NAME: name.trim(), COMPANY_LOGO_ID: next});
    if (previous && previous !== next) trashPhoto_(previous);
    return getCompanySettings();
  } finally {
    lock.releaseLock();
  }
}
