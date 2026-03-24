// ======================= モジュール読み込み =======================
const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const { google } = require('googleapis');
const { GoogleAuth } = require('google-auth-library');
const axios = require('axios');
const cors = require('cors');
const express = require('express'); // ✅ NEW

// ======================= USER AGENT =======================
const userAgent = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/116.0.0.0 Safari/537.36";

let keyPath;
if (app.isPackaged) {
  keyPath = path.join(process.resourcesPath, 'service-account.json');
} else {
  keyPath = path.join(process.cwd(), 'service-account.json');
}

// ======================= Puppeteer =======================
const pie = require(app.isPackaged
  ? path.join(process.resourcesPath, 'app.asar.unpacked/node_modules/puppeteer-in-electron')
  : 'puppeteer-in-electron');

const puppeteer = require(app.isPackaged
  ? path.join(process.resourcesPath, 'app.asar.unpacked/node_modules/puppeteer-core')
  : 'puppeteer-core');

// ======================= Electron Flags =======================
app.commandLine.appendSwitch('disable-bluetooth');
app.commandLine.appendSwitch('disable-features', 'Bluetooth,WinrtBluetooth,CrOSBluetooth');
app.commandLine.appendSwitch('bluetooth-low-energy-scanner-disabled');
app.commandLine.appendSwitch('disable-logging');
app.commandLine.appendSwitch('log-level', '3');

// ======================= グローバル変数 =======================
let activeSessions = new Map();
let pendingRowsQueue = [];
let isProcessingQueue = false;
let sessionCounter = 0;

let mainWindow;
let runningStatus = false;
let sheets = google.sheets('v4');

let url = 'https://www.apple.com/jp/shop/buy-iphone/';
let pcSelection = '1';
let modelOption, colorOption, storageOption, quantityOption;
let confirmOption, payOption, deliveryOption, storeOption;
let storeMonitoringInterval, workerId;
let spreadsheetKey1, spreadsheetKey2, spreadsheetKey3, spreadsheetKey4;

const stockSpreadsheetId = '1tY-3_3ClN26w_Z0AYRmUofgQ0rarJsh8epxPAuPw0D0';

// ======================= セッション状態（モバイル監視用） =======================
const sessionStates = new Map();

const STEP_LABELS_JA = {
  navigation_started: 'Apple公式サイトにアクセス中',
  select_model: 'モデルを選択中',
  select_color: 'カラーを選択中',
  select_storage: 'ストレージ容量を選択中',
  add_to_cart: 'カートに追加中',
  go_to_checkout: 'チェックアウト画面へ移動中',
  input_shipping: '配送先住所を入力中',
  input_pickup: '受取情報を入力中',
  input_payment: '支払い情報を入力中',
  review_order: '注文内容を確認中',
  complete: '注文完了',
  error: 'エラー発生',
};

function updateSessionState(sessionId, patch) {
  const prev = sessionStates.get(sessionId) || {};
  const next = {
    ...prev,
    ...patch,
    updatedAt: new Date().toISOString(),
  };
  sessionStates.set(sessionId, next);
}

// ======================= 🔥 API SERVER =======================
function startApiServer() {
  const server = express();
  server.options('*', cors());
  server.use(express.json());

  const API_KEY = "578d28153ca5fc4f7b20c1e4df7c51f87638627b3261165ed8a9803129d85d97";

   server.use(
    cors({
      origin: '*',
      methods: ['GET', 'POST', 'OPTIONS'],
      allowedHeaders: ['x-api-key', 'Content-Type'],
      credentials: false,
      maxAge: 86400,
    })
  );


  server.post('/start', async (req, res) => {
    try {
      if (runningStatus) {
        return res.json({ success: false, message: "Already running" });
      }

      const args = req.body;
            console.log(args, 'ref');

      runningStatus = true;

      url = args.itemUrl || url;
      pcSelection = args.pcSelection || '1';
      modelOption = args.modelOption;
      colorOption = args.colorOption;
      storageOption = args.storageOption;
      quantityOption = args.quantityOption;
      confirmOption = args.confirmOption;
      payOption = args.payOption;
      deliveryOption = args.deliveryOption || 'delivery';
      storeOption = args.storeOption;
      storeMonitoringInterval = Math.max(args.storeMonitoringInterval || 5, 5);
      workerId = args.workerId || 'api';

      spreadsheetKey1 = args.spreadsheetKey;
      spreadsheetKey2 = args.spreadsheetKey2;
      spreadsheetKey3 = args.spreadsheetKey3;
      spreadsheetKey4 = args.spreadsheetKey4;

      await initializeQueue();
      processQueue();

      res.json({ success: true, message: "Started" });

    } catch (err) {
      runningStatus = false;
      res.status(500).json({ error: err.message });
    }
  });

  server.post('/stop', (req, res) => {
    runningStatus = false;

    for (const [id, s] of activeSessions) {
      if (s.window && !s.window.isDestroyed()) s.window.close();
    }
    activeSessions.clear();
    pendingRowsQueue = [];

    res.json({ success: true });
  });

  server.get('/status', (req, res) => {
    const sessions = Array.from(sessionStates.values()).map(s => ({
      ...s,
      statusLabelJa:
        s.status === 'running'
          ? '実行中'
          : s.status === 'completed'
          ? '完了'
          : s.status === 'error'
          ? 'エラー'
          : '待機中',
      stepLabelJa: s.step ? STEP_LABELS_JA[s.step] || '' : '',
    }));

    res.json({
      running: runningStatus,
      queueCount: pendingRowsQueue.length,
      activeSessionCount: sessions.length,
      sessions,
    });
  });

  server.listen(3000, '0.0.0.0', () => {
    console.log("🚀 API READY http://0.0.0.0:3000");
  });
}

// ======================= Queue =======================
function safeStr(v){ return v ? String(v) : '' }

async function initializeQueue() {
  pendingRowsQueue = [];
  const targets = getCandidateTargets();

  for (const t of targets) {
    const rows = await getRowsFromSpreadsheet(t.id, t.tab);
    for (const r of rows) {
      if (!r.status || r.status === '') {
        pendingRowsQueue.push({
          spreadsheetId: t.id,
          sheetName: t.tab,
          rowData: r,
          rowNum: r.rowNum
        });
      }
    }
  }

  pendingRowsQueue.sort(() => Math.random() - 0.5);
}

async function processQueue() {
  if (isProcessingQueue || !runningStatus) return;
  isProcessingQueue = true;

  while (pendingRowsQueue.length > 0) {
    const item = pendingRowsQueue.shift();
    try {
      await lockRow(item.spreadsheetId, item.sheetName, item.rowNum);
      const { sessionId, browserWindow } = await createBrowserSession(item);
      scrapeWebsite(sessionId, browserWindow, item.rowData);
    } catch (e) {
      pendingRowsQueue.push(item);
    }
  }

  isProcessingQueue = false;
}


// ======================= スプレッドシート操作関数 =======================
async function lockAndMarkDoneInSpreadsheet(rowNum, spreadsheetId, sheetName, orderNumber = '') {
  const auth = new GoogleAuth({ 
    keyFile: keyPath, 
    scopes: ['https://www.googleapis.com/auth/spreadsheets'] 
  });

  const client = await auth.getClient();
  google.options({ auth: client });

  const currentTime = new Date().toISOString();
  const newStatus = `Done`;

  try {
    await sheets.spreadsheets.values.update({
      spreadsheetId: spreadsheetId,
      range: `${sheetName}!M${rowNum}`,
      valueInputOption: 'USER_ENTERED',
      resource: { values: [[currentTime]] },
    });

    await sheets.spreadsheets.values.update({
      spreadsheetId: spreadsheetId,
      range: `${sheetName}!N${rowNum}`,
      valueInputOption: 'USER_ENTERED',
      resource: { values: [[newStatus]] },
    });

    if (orderNumber) {
      await sheets.spreadsheets.values.update({
        spreadsheetId: spreadsheetId,
        range: `${sheetName}!O${rowNum}`,
        valueInputOption: 'USER_ENTERED',
        resource: { values: [[orderNumber]] },
      });
    }

    mainWindow.webContents.send('log', { sessionId: 'main', message: `[Sheets] Row ${rowNum} marked as Done` });
  } catch (error) {
    console.error('Error marking row as done:', error);
    mainWindow.webContents.send('log', { sessionId: 'main', message: `[Sheets] Error: ${error.message}` });
  }
}

async function unLockToSpreadsheet(rowNum, spreadsheetId, sheetName) {
  const auth = new GoogleAuth({ 
    keyFile:keyPath, 
    scopes: ['https://www.googleapis.com/auth/spreadsheets'] 
  });
  const client = await auth.getClient();
  google.options({ auth: client });
  console.log(spreadsheetId, 'best')
  try {
    await sheets.spreadsheets.values.update({
      spreadsheetId: spreadsheetId,
      range: `${sheetName}!M${rowNum}:N${rowNum}`,
      valueInputOption: 'USER_ENTERED',
      resource: { values: [['', '']] },
    });
    mainWindow.webContents.send('log', { sessionId: 'main', message: `[Sheets] Row ${rowNum} unlocked` });
  } catch (error) {
    console.error('Error unlocking row:', error);
    mainWindow.webContents.send('log', { sessionId: 'main', message: `[Sheets] Unlock error: ${error.message}` });
  }
}

// ======================= Browser =======================
async function createBrowserSession(queueItem) {
  const sessionId = ++sessionCounter;
  const browser = await pie.connect(app, puppeteer);
  
  const browserWindow = new BrowserWindow({
    width: 900,
    height: 700,
    title: `セッション ${sessionId} - ${queueItem.rowData.firstName || 'Unknown'}`,
    icon: './title.png',
   
    webPreferences: {
      partition: `セッション-${sessionId}-${Date.now()}`,
      nodeIntegration: true,
      contextIsolation: false
    }
  });

  browserWindow.webContents.setUserAgent(userAgent);
  browserWindow.webContents.on('dom-ready', () => {
    browserWindow.webContents.executeJavaScript(`
      document.documentElement.classList.add('セッション-${sessionId}');
    `);
  });
  
  const sessionPage = await pie.getPage(browser, browserWindow);
  console.log(sessionId,queueItem, 'log')
  activeSessions.set(sessionId, {
    window: browserWindow,
    page: sessionPage,
    sessionId,
    spreadsheetId: queueItem.spreadsheetId,
    sheetName: queueItem.sheetName,
    rowData: queueItem.rowData,
    rowNum: queueItem.rowNum
  });

  // セッション状態初期化（モバイル監視用）
  updateSessionState(sessionId, {
    id: sessionId,
    rowNum: queueItem.rowNum,
    status: 'running',
    progress: 0,
    step: 'navigation_started',
    messageJa: 'Apple公式サイトにアクセス中',
  });

  mainWindow.webContents.send('log', { sessionId: 'main', message: `Created session ${sessionId} for row ${queueItem.rowNum}` });

  return { sessionId, browserWindow };
}

// ======================= Sheets =======================
function getCandidateTargets() {
  return [{ id: spreadsheetKey1, tab: 'list' }];
}

async function getRowsFromSpreadsheet(id, tab) {
  const auth = new GoogleAuth({ keyFile: keyPath, scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
  const client = await auth.getClient();
  google.options({ auth: client });

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: id,
    range: `${tab}!A:R`
  });

  const rows = res.data.values || [];
  const headers = rows.shift();

  return rows.map((r, i) => {
    let obj = {};
    headers.forEach((h, idx) => obj[h] = r[idx]);
    obj.rowNum = i + 2;
    return obj;
  });
}

function safeStr(v){ return v ? String(v) : '' }
function findRowUrl(rowData) {
  if (!rowData) return '';
  const candidates = new Set(['itemUrl', 'url', 'pageUrl', 'itemPageUrl'].map(s => s.toLowerCase()));
  for (const [k, v] of Object.entries(rowData)) {
    if (candidates.has(String(k).toLowerCase())) return v;
  }
  return '';
}

function findRowIphoneId(rowData) {
  if (!rowData) return '';
  const candidates = new Set(['iphone_id', 'iphoneId', 'iphoneID'].map(s => s.toLowerCase()));
  for (const [k, v] of Object.entries(rowData)) {
    if (candidates.has(String(k).toLowerCase())) return v;
  }
  return '';
}

function joinUrl(base, tail) {
  if (!base) return tail || '';
  if (!tail) return base;
  const b = String(base);
  const t = String(tail);
  const bEndsWithSlash = b.endsWith('/');
  const tStartsWithSlash = t.startsWith('/');
  if (bEndsWithSlash && tStartsWithSlash) return b + t.slice(1);
  if (!bEndsWithSlash && !tStartsWithSlash) return b + '/' + t;
  return b + t;
}

async function lockRow(id, tab, row) {
  const auth = new GoogleAuth({ keyFile: keyPath, scopes: ['https://www.googleapis.com/auth/spreadsheets'] });
  const client = await auth.getClient();
  google.options({ auth: client });

  await sheets.spreadsheets.values.update({
    spreadsheetId: id,
    range: `${tab}!N${row}`,
    valueInputOption: 'USER_ENTERED',
    resource: { values: [['locked']] }
  });
}

// ======================= Scraper =======================
async function scrapeWebsite(sessionId, targetWindow, data) {
   
  const sessionInfo= activeSessions.get(sessionId);
  // 入力用の安全な文字列化ヘルパー
 
  if (!sessionInfo) {
    if (targetWindow && !targetWindow.isDestroyed()) {
      targetWindow.close();
    }
    return;
  }
  
  const currentSpreadsheetId = sessionInfo.spreadsheetId;
  const currentSheetName = sessionInfo.sheetName;
  const spreadsheetInfo = sessionInfo.rowData;
  const rowNum = sessionInfo.rowNum;
  const sessionUrl =
    joinUrl(
      safeStr(findRowUrl(data)) || url,
      safeStr(findRowIphoneId(data)),
    );
   console.log(sessionUrl, 'SessionUrl')
  

  try {
    mainWindow.webContents.send('status', { sessionId, status: '実行中...' });
    mainWindow.webContents.send('log', { sessionId, message: `-------- 実行開始 (Row ${rowNum}) --------` });

    // 進捗: サイト遷移開始
    updateSessionState(sessionId, {
      status: 'running',
      progress: 5,
      step: 'navigation_started',
      messageJa: 'Apple公式サイトにアクセス中',
    });

    // Puppeteer + Electron セッション開始
    const page = activeSessions.get(sessionId).page;
    page.setDefaultTimeout(120000);
    page.setDefaultNavigationTimeout(120000);
    
    const client = await page.target().createCDPSession();
    await client.send('Runtime.enable');
    await client.send('Runtime.setAsyncCallStackDepth', { maxDepth: 32 });

    // Apple公式サイトに遷移
    await page.goto(sessionUrl , { waitUntil: 'domcontentloaded' });
    mainWindow.webContents.send('log', { sessionId, message: `Navigated to ${sessionUrl}` });

    updateSessionState(sessionId, {
      progress: 10,
      step: 'select_model',
      messageJa: 'モデルを選択中',
    });
    await page.waitForTimeout(500);
   
    // ======================= 商品選択処理 =======================
    const maxRetries = 3;
    for (let i = 0; i < maxRetries; i++) {
      // モデル選択
      const colorSelector = `.rf-bfe-dimension-dimensioncolor > fieldset > ul > li:nth-child(${1 + Number(colorOption)}) > label`;
      if (modelOption != 'skip') {
        const is_model = await page.evaluate(() => {
          const all = document.querySelectorAll('[name="dimensionScreensize"]');
          return all && all.length > 0 ? 'has_model' : 'skip';
        });

        if (is_model != 'skip') {
          const screenSelector = '[name="dimensionScreensize"]';
          await waitForClickableElement(screenSelector, sessionId);
          await page.waitForTimeout(300);
          await directClick(screenSelector, Number(modelOption), sessionId, page);
        }
      }

      // カラー選択
      await waitForClickableElement(colorSelector, sessionId);
      await page.waitForTimeout(300);
      await directClick(colorSelector, 0, sessionId, page);

      updateSessionState(sessionId, {
        progress: 25,
        step: 'select_color',
        messageJa: 'カラーを選択中',
      });

      const capacitySelector = await page.evaluate(() => {
        const all = document.querySelectorAll(
          '.rc-dimension-selector-row [name="dimensionCapacity"]',
        );
        return all && all.length > 1
          ? '[name="dimensionCapacity"]'
          : '#noTradeIn';
      });

      await waitForClickableElement(capacitySelector, sessionId);


      // Apple Store在庫監視モード
      if (deliveryOption == 'appleStore' && i == 0) {
        await page.setRequestInterception(true);
        let targetUrl;
        let itemNumber;

        const interceptedRequestHandler = (interceptedRequest) => {
          if (interceptedRequest.url().includes('fulfillment-messages')) {
            let parsedUrl = new URL(interceptedRequest.url());
            let params = new URLSearchParams(parsedUrl.search);
            params.set('store', storeOption);
            itemNumber = params.get('parts.0');
            targetUrl = `${parsedUrl.origin}${parsedUrl.pathname}?${params.toString()}`;
            interceptedRequest.continue({ url: targetUrl });
          } else {
            interceptedRequest.continue();
          }
        };
        page.on('request', interceptedRequestHandler);
        await directClick(capacitySelector, Number(storageOption), sessionId, page);

        while (runningStatus) {
          mainWindow.webContents.send('log', { sessionId, message: `Checking ${itemNumber} stock...` });
          const spreadsheetStock = await checkStockBySpreadsheet(itemNumber, storeOption);
          if (!spreadsheetStock) {
            const response = await axios.get(targetUrl);
            const data = response.data;
            const firstKey = Object.keys(data.body.content.pickupMessage.stores[0].partsAvailability)[0];
            const pickupDisplay = data.body.content.pickupMessage.stores[0].partsAvailability[firstKey].pickupDisplay;

            if (pickupDisplay == 'unavailable') {
              await page.waitForTimeout(storeMonitoringInterval * 1000);
            } else {
              mainWindow.webContents.send('log', { sessionId, message: `${firstKey} is in stock!!` });
              break;
            }
          } else if (spreadsheetStock == 'available') {
            mainWindow.webContents.send('log', { sessionId, message: `${itemNumber} is in stock!!` });
            break;
          } else {
            await page.waitForTimeout(storeMonitoringInterval * 1000);
          }
        }

        page.off('request', interceptedRequestHandler);
        await page.setRequestInterception(false);
        await page.goto(sessionUrl, { waitUntil: 'networkidle0'});
      } else {
        await page.waitForSelector('#noTradeIn');
        await directClick(capacitySelector, Number(storageOption), sessionId, page);
      }

      updateSessionState(sessionId, {
        progress: 35,
        step: 'select_storage',
        messageJa: 'ストレージ容量を選択中',
      });
      const noTradeInSelector = '#noTradeIn';
      await waitForClickableElement(noTradeInSelector, sessionId);
      await page.waitForTimeout(300);
      await page.waitForSelector('[value="UNLOCKED_JP"]', { visible: true});
      await directClick(noTradeInSelector, 0, sessionId, page);
      await page.waitForTimeout(1000);
      const carrierModelSelector = '[value="UNLOCKED_JP"]';
      await waitForClickableElement(carrierModelSelector, sessionId);
      await page.waitForTimeout(300);
      await page.waitForSelector('[value="fullprice"]', { visible: true});
      await directClick(carrierModelSelector, 0, sessionId, page);

      const purchaseInSelector = '[value="fullprice"]';
      await waitForClickableElement(purchaseInSelector, sessionId);
      await page.waitForTimeout(300);
      await page.waitForSelector('[data-autom="noapplecare"]', { visible: true});
      await directClick(purchaseInSelector, 0, sessionId, page);
      await page.waitForTimeout(1000);

      await page.waitForTimeout(300);
      const applecareSelector = '[data-autom="noapplecare"]';
      await waitForClickableElement(applecareSelector, sessionId);
      await page.waitForTimeout(300);
      await page.waitForSelector('[name="add-to-cart"]', { visible: true});
      await directClick(applecareSelector, 0, sessionId, page);

      const addToCartSelector = '[name="add-to-cart"]';
      await waitForClickableElement(addToCartSelector, sessionId);
      await page.waitForTimeout(300);
      await directClick(addToCartSelector, 0, sessionId, page);

      updateSessionState(sessionId, {
        progress: 50,
        step: 'add_to_cart',
        messageJa: 'カートに追加中',
      });

      const goToBagSelector = '.rc-summaryheader-button button';
      await waitForClickableElement(goToBagSelector, sessionId);
      await page.waitForTimeout(300);
      await directClick(goToBagSelector, 0, sessionId, page);

      const quantitySelector = '.rs-quantity-dropdown';
      try {
        await page.waitForSelector(quantitySelector);
        page.waitForTimeout(300);
        await page.select(quantitySelector, quantityOption);
      } catch (error) {
        console.log(`数量変更失敗: ${error}`);
      }

      const goToCheckoutSelector = '[id="shoppingCart.actions.navCheckoutOtherPayments"]';
      await waitForClickableElement(goToCheckoutSelector, sessionId);
      await page.waitForTimeout(300);
      await directClick(goToCheckoutSelector, 0, sessionId, page);

      updateSessionState(sessionId, {
        progress: 60,
        step: 'go_to_checkout',
        messageJa: 'チェックアウト画面へ移動中',
      });

      // ログイン処理
      const guestLoginSelector = '[id="signIn.guestLogin.guestLogin"]';
      await waitForClickableElement(guestLoginSelector, sessionId);
      await page.waitForTimeout(300);
      await directClick(guestLoginSelector, 0, sessionId, page);

      await page.waitForTimeout(2000);
      
      // ======================= 配送 or 店舗受取 =======================
      if (deliveryOption == 'delivery') {
        const locationEditSelector = '.rs-edit-location-button';
        await waitForClickableElement(locationEditSelector, sessionId);
        await page.waitForTimeout(300);
        await directClick(locationEditSelector, 0, sessionId, page);
        const postalSelector='[id="checkout.fulfillment.deliveryTab.delivery.deliveryLocation.address.deliveryWarmingSubLayout.postalCode"]';
        await page.waitForSelector(postalSelector, { timeout: 10000 });
        await page.$eval(postalSelector, el => el.value = '');
        await page.type(postalSelector, safeStr(spreadsheetInfo?.postalCode));
        
        await page.$eval('[id="checkout.fulfillment.deliveryTab.delivery.deliveryLocation.address.deliveryWarmingSubLayout.postalCode"]', el => el.value = '');
        await page.type('[id="checkout.fulfillment.deliveryTab.delivery.deliveryLocation.address.deliveryWarmingSubLayout.postalCode"]', safeStr(spreadsheetInfo?.postalCode));
        await page.select('[id="checkout.fulfillment.deliveryTab.delivery.deliveryLocation.address.deliveryWarmingSubLayout.state"]', safeStr(spreadsheetInfo?.state));
        
        await page.waitForTimeout(2000);
        const deliverySelector = '#rs-checkout-continue-button-bottom';
        await waitForClickableElement(deliverySelector, sessionId);
        await page.waitForTimeout(300);
        await directClick(deliverySelector, 0, sessionId, page);
        await page.waitForTimeout(1000);

        // 配送住所入力
        await page.waitForSelector('[id="checkout.shipping.addressSelector.newAddress.address.lastName"]');
        await page.type('[id="checkout.shipping.addressSelector.newAddress.address.lastName"]', safeStr(spreadsheetInfo?.lastName));
        await page.type('[id="checkout.shipping.addressSelector.newAddress.address.firstName"]', safeStr(spreadsheetInfo?.firstName));
        await page.$eval('[id="checkout.shipping.addressSelector.newAddress.address.postalCode"]', el => el.value = '');
        await page.type('[id="checkout.shipping.addressSelector.newAddress.address.postalCode"]', safeStr(spreadsheetInfo?.postalCode));
        await page.select('[id="checkout.shipping.addressSelector.newAddress.address.state"]', safeStr(spreadsheetInfo?.state));
        await page.type('[id="checkout.shipping.addressSelector.newAddress.address.city"]', safeStr(spreadsheetInfo?.city));
        await page.type('[id="checkout.shipping.addressSelector.newAddress.address.street"]', safeStr(spreadsheetInfo?.street));
        await page.type('[id="checkout.shipping.addressSelector.newAddress.address.street2"]', safeStr(spreadsheetInfo?.street2));
        await page.evaluate(() => {
          const field = document.querySelector('[id="checkout.shipping.addressContactEmail.address.emailAddress');
          if (field) field.value = '';
        });
        await page.type('[id="checkout.shipping.addressContactEmail.address.emailAddress"]', safeStr(spreadsheetInfo?.emailAddress));
        await page.type('[id="checkout.shipping.addressContactPhone.address.mobilePhone"]', safeStr(spreadsheetInfo?.mobilePhone));

      } else if (deliveryOption == 'convenienceStore' || deliveryOption == 'appleStore') {
        const convenienceStoreSelector = '[for="fulfillmentOptionButtonGroup1"]';
        await waitForClickableElement(convenienceStoreSelector, sessionId);
        await directClick(convenienceStoreSelector, 0, sessionId, page);

        try {
          const locationEditSelector = '.rs-edit-location-button';
          await page.waitForSelector(locationEditSelector, { timeout: 3000 });
          await directClick(locationEditSelector, 0, sessionId, page);
        } catch { }

        const storeLocatorSearchInputSelector = '[id="checkout.fulfillment.pickupTab.pickup.storeLocator.searchInput"]';
        await page.waitForSelector(storeLocatorSearchInputSelector);
        await page.$eval(storeLocatorSearchInputSelector, el => el.value = '');
        await page.type(storeLocatorSearchInputSelector, safeStr(spreadsheetInfo?.postalCode));

        const locationEditButtonSelector = '[id="checkout.fulfillment.pickupTab.pickup.storeLocator.search"]';
        await waitForClickableElement(locationEditButtonSelector, sessionId);
        await directClick(locationEditButtonSelector, 0, sessionId, page);

        if (deliveryOption == 'convenienceStore') {
          const availableConvenienceStoreSelector = '.rt-storelocator-store-marker-pup';
          await page.waitForSelector(availableConvenienceStoreSelector);
          await directClick(availableConvenienceStoreSelector, 0, sessionId, page);
        } else if (deliveryOption == 'appleStore') {
          const availableAppleStoreSelector = `input[value="${storeOption}"]`;
          await page.waitForSelector(availableAppleStoreSelector);
          await directClick(availableAppleStoreSelector, 0, sessionId, page);
        }

        const deliverySelector = '#rs-checkout-continue-button-bottom';
        await waitForClickableElement(deliverySelector, sessionId);
        await directClick(deliverySelector, 0, sessionId, page);

        // 受取人情報
        await page.waitForSelector('[id="checkout.pickupContact.selfPickupContact.selfContact.address.lastName"]');
        await page.type('[id="checkout.pickupContact.selfPickupContact.selfContact.address.lastName"]', safeStr(spreadsheetInfo?.lastName));
        await page.type('[id="checkout.pickupContact.selfPickupContact.selfContact.address.firstName"]', safeStr(spreadsheetInfo?.firstName));
        await page.type('[id="checkout.pickupContact.selfPickupContact.selfContact.address.emailAddress"]', safeStr(spreadsheetInfo?.emailAddress));
        await page.type('[id="checkout.pickupContact.selfPickupContact.selfContact.address.mobilePhone"]', safeStr(spreadsheetInfo?.mobilePhone));
      }

      updateSessionState(sessionId, {
        progress: 75,
        step: deliveryOption === 'delivery' ? 'input_shipping' : 'input_pickup',
        messageJa:
          deliveryOption === 'delivery'
            ? '配送先住所を入力中'
            : '受取情報を入力中',
      });

      // ======================= 支払い選択 =======================
      const checkoutContinueSelector = '#rs-checkout-continue-button-bottom';
      await waitForClickableElement(checkoutContinueSelector, sessionId);
      await page.waitForTimeout(300);
      await directClick(checkoutContinueSelector, 0, sessionId, page);

      if (payOption == 'creditcard') {
        const creditSelector = '[id="checkout.billing.billingoptions.credit"]';
        const cardNumber=spreadsheetInfo?.cardNumber;
        await waitForClickableElement(creditSelector, sessionId);
        await page.waitForTimeout(300);
        await directClick(creditSelector, 0, sessionId, page);
        await page.waitForTimeout(500);

        const cardNumberSelector = '[id="checkout.billing.billingOptions.selectedBillingOptions.creditCard.cardInputs.cardInput-0.cardNumber"]';
        await waitForClickableElement(cardNumberSelector, sessionId);
        console.log(cardNumber, 'cardNumber');
        await page.type(cardNumberSelector, cardNumber, { delay: 100 });
        await page.waitForTimeout(500);
        await page.type('[id="checkout.billing.billingOptions.selectedBillingOptions.creditCard.cardInputs.cardInput-0.expiration"]', safeStr(spreadsheetInfo?.expiration));
        await page.type('[id="checkout.billing.billingOptions.selectedBillingOptions.creditCard.cardInputs.cardInput-0.securityCode"]', safeStr(spreadsheetInfo?.securityCode));

      } else if (payOption == 'bank') {
        const bankSelector = '[id="checkout.billing.billingoptions.apple_pay"]';
        await waitForClickableElement(bankSelector, sessionId);
        await page.waitForTimeout(300);
        await directClick(bankSelector, 0, sessionId, page);
      }

      updateSessionState(sessionId, {
        progress: 85,
        step: 'input_payment',
        messageJa: '支払い情報を入力中',
      });

      // ======================= 請求先住所（受取の場合のみ） =======================
      if (deliveryOption == 'convenienceStore' || deliveryOption == 'appleStore') {
        if (payOption == 'creditcard') {
          await page.waitForSelector('[id="checkout.billing.billingOptions.selectedBillingOptions.creditCard.billingAddress.address.lastName"]');
          await page.type('[id="checkout.billing.billingOptions.selectedBillingOptions.creditCard.billingAddress.address.lastName"]', safeStr(spreadsheetInfo?.lastName));
          await page.type('[id="checkout.billing.billingOptions.selectedBillingOptions.creditCard.billingAddress.address.firstName"]', safeStr(spreadsheetInfo?.firstName));
          await page.$eval('[id="checkout.billing.billingOptions.selectedBillingOptions.creditCard.billingAddress.address.postalCode"]', el => el.value = '');
          await page.type('[id="checkout.billing.billingOptions.selectedBillingOptions.creditCard.billingAddress.address.postalCode"]', safeStr(spreadsheetInfo?.postalCode));
          await page.select('[id="checkout.billing.billingOptions.selectedBillingOptions.creditCard.billingAddress.address.state"]', safeStr(spreadsheetInfo?.state));
          await page.type('[id="checkout.billing.billingOptions.selectedBillingOptions.creditCard.billingAddress.address.city"]', safeStr(spreadsheetInfo?.city));
          await page.type('[id="checkout.billing.billingOptions.selectedBillingOptions.creditCard.billingAddress.address.street"]', safeStr(spreadsheetInfo?.street));
          await page.type('[id="checkout.billing.billingOptions.selectedBillingOptions.creditCard.billingAddress.address.street2"]', safeStr(spreadsheetInfo?.street2));
        } else if (payOption == 'bank') {
          await page.waitForSelector('[id="checkout.billing.billingOptions.selectedBillingOptions.wireTransfer.billingAddress.address.lastName"]');
          await page.type('[id="checkout.billing.billingOptions.selectedBillingOptions.wireTransfer.billingAddress.address.lastName"]', safeStr(spreadsheetInfo?.lastName));
          await page.type('[id="checkout.billing.billingOptions.selectedBillingOptions.wireTransfer.billingAddress.address.firstName"]', safeStr(spreadsheetInfo?.firstName));
          await page.type('[id="checkout.billing.billingOptions.selectedBillingOptions.wireTransfer.billingAddress.address.street"]', safeStr(spreadsheetInfo?.street));
          await page.type('[id="checkout.billing.billingOptions.selectedBillingOptions.wireTransfer.billingAddress.address.street2"]', safeStr(spreadsheetInfo?.street2));
          await page.$eval('[id="checkout.billing.billingOptions.selectedBillingOptions.wireTransfer.billingAddress.address.postalCode"]', el => el.value = '');
          await page.type('[id="checkout.billing.billingOptions.selectedBillingOptions.wireTransfer.billingAddress.address.postalCode"]', safeStr(spreadsheetInfo?.postalCode));
          await page.select('[id="checkout.billing.billingOptions.selectedBillingOptions.wireTransfer.billingAddress.address.state"]', safeStr(spreadsheetInfo?.state));
          await page.type('[id="checkout.billing.billingOptions.selectedBillingOptions.wireTransfer.billingAddress.address.city"]', safeStr(spreadsheetInfo?.city));
        }
      }

      // ======================= 注文確認 =======================
      const checkoutContinue2Selector = '#rs-checkout-continue-button-bottom';
      await waitForClickableElement(checkoutContinue2Selector, sessionId);
      await page.waitForTimeout(300);
      await directClick(checkoutContinue2Selector, 0, sessionId, page);
      page.waitForTimeout(10000);
      if (confirmOption == 'true') {
        await page.waitForSelector('.rs-review-summary');
        const confirmSelector = '#rs-checkout-continue-button-bottom';
        await waitForClickableElement(confirmSelector, sessionId);
        await page.waitForTimeout(300);
        await directClick(confirmSelector, 0, sessionId, page);
      }

      // 購入完了処理
      await lockAndMarkDoneInSpreadsheet(rowNum, currentSpreadsheetId, currentSheetName);
      break; // 成功したらループを抜ける
    }

    mainWindow.webContents.send('log', { sessionId, message: `注文完了 - ログアウト処理開始` });
    mainWindow.webContents.send('status', { sessionId, status: '完了' });
    mainWindow.webContents.send('log', { sessionId, message: `-------- 正常終了 --------` });
    page.waitForTimeout(500000);

    updateSessionState(sessionId, {
      status: 'completed',
      progress: 100,
      step: 'complete',
      messageJa: '注文完了',
    });
  } catch (error) {
    mainWindow.webContents.send('status', { sessionId, status: '失敗' });
    mainWindow.webContents.send('log', { sessionId, message: `Error: ${error.message}` });
    
    // エラー時は行をアンロック
    try {
      await unLockToSpreadsheet(rowNum, currentSpreadsheetId, currentSheetName);
    } catch (unlockError) {
      console.error('Unlock error:', unlockError);
      mainWindow.webContents.send('log', { sessionId, message: `Unlock error: ${unlockError.message}` });
    }

    updateSessionState(sessionId, {
      status: 'error',
      step: 'error',
      messageJa: `エラー発生: ${error.message}`,
    });
  } finally {
    // セッションクリーンアップ
    await new Promise(resolve => setTimeout(resolve, 30000));
    activeSessions.delete(sessionId);
    if (targetWindow && !targetWindow.isDestroyed()) {
      targetWindow.close();
    }
    
    // 次の処理を開始
    setTimeout(processQueue, 2000);
  }
}



// ======================= ユーティリティ関数 =======================

async function directClick(selector, index = 0, sessionId, page) {

  try {
    await page.waitForSelector(selector, { visible: true, timeout: 16000 });
    await safePageEvaluate((sel, idx) => {
      const el = document.querySelectorAll(sel)[idx];
      if (!el) throw new Error(`Element not found: ${sel} at index ${idx}`);
      el.scrollIntoView({ behavior: 'smooth', block: 'center' });
      el.click();
    }, selector, index, sessionId, page);
    await page.waitForTimeout(150);
  } catch (err) {
    console.log(`directClick Error: ${err.message}`);
    mainWindow.webContents.send('log', { sessionId, message: `Click failed for selector: ${selector} - ${err.message}` });
    throw err;
  }
}


async function waitForClickableElement(selector, sessionId, pageOrFrame = null, timeout = 60000) {
  const page = pageOrFrame || activeSessions.get(sessionId).page;
  console.log(`[${sessionId}] waiting for ${selector} to be clickable`);

  await page.waitForFunction(
    sel => {
      const el = document.querySelector(sel);
      if (!el) return false;
      const style = window.getComputedStyle(el);
      if (style.visibility === 'hidden' || style.display === 'none') return false;
      if (el.disabled) return false;
      const rect = el.getBoundingClientRect();
      return rect.width > 0 && rect.height > 0;
    },
    { timeout },
    selector
  );

  // ensure it’s scrolled into view
  await page.$eval(selector, el => el.scrollIntoView({ behavior: 'auto', block: 'center' }));
  await page.waitForTimeout(300); // give JS time to attach listeners
}
async function safePageEvaluate(fn, ...args) {
  const page =activeSessions.get(args[2]).page

  const maxRetries = 3;
  for (let i = 0; i < maxRetries; i++) {
    try {
      return await page.evaluate(fn, ...args);
    } catch (error) {
      if (error.message.includes('Runtime.callFunctionOn timed out')) {
        if (i === maxRetries - 1) throw error;
        await page.waitForTimeout(3000);
      } else {
        throw error;
      }
    }
  }
}
// ======================= Window =======================
function createMainWindow() {
  mainWindow = new BrowserWindow({
    width: 600,
    height: 600,
    show: false,
    webPreferences: { nodeIntegration: true, contextIsolation: false }
  });
  mainWindow.loadFile('index.html');
}

// ======================= App =======================
pie.initialize(app);

app.whenReady().then(() => {
  createMainWindow();
  startApiServer(); // ✅ IMPORTANT
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});