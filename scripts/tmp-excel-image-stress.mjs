import fs from 'node:fs'
import path from 'node:path'
import { chromium } from 'playwright-core'

const EDGE_PATHS = [
  'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
  'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe'
]

const BASE_URL = process.env.BASE_URL || 'http://127.0.0.1'
const API_URL = process.env.API_URL || 'http://127.0.0.1:8080'
const LOGIN_USERNAME = process.env.LOGIN_USERNAME || 'admin'
const LOGIN_PASSWORD = process.env.LOGIN_PASSWORD || 'admin123'
const LOAD_ROUNDS = Number.parseInt(process.env.LOAD_ROUNDS || '3', 10)
const SAVE_ROUNDS = Number.parseInt(process.env.SAVE_ROUNDS || '3', 10)

function pickEdgePath() {
  for (const candidate of EDGE_PATHS) {
    if (fs.existsSync(candidate)) {
      return candidate
    }
  }
  throw new Error('Microsoft Edge executable not found')
}

function nowIso() {
  return new Date().toISOString()
}

async function loginAndGetToken() {
  const response = await fetch(`${API_URL}/login`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({
      username: LOGIN_USERNAME,
      password: LOGIN_PASSWORD,
      code: '',
      uuid: ''
    })
  })
  const data = await response.json()
  if (!response.ok || Number(data?.code) !== 200 || !data?.token) {
    throw new Error(`login failed: ${JSON.stringify(data)}`)
  }
  return data.token
}

async function uploadWorkbookViaApi(token, workbookPath) {
  const uploadStart = Date.now()
  const bytes = fs.readFileSync(workbookPath)
  const formData = new FormData()
  const fileBlob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
  formData.append('file', fileBlob, path.basename(workbookPath))
  const response = await fetch(`${API_URL}/tool/excel-editor/upload`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`
    },
    body: formData
  })
  const uploadJson = await response.json()
  const uploadEnd = Date.now()

  return {
    fileName: uploadJson?.data?.currentFileName || path.basename(workbookPath),
    httpStatus: response.status,
    responseCode: uploadJson?.code ?? null,
    responseMsg: uploadJson?.msg ?? '',
    uploadMs: uploadEnd - uploadStart
  }
}

async function runLoadRound(page, fileName, round) {
  const editorUrl = `${BASE_URL}/excel-editor/index?fileName=${encodeURIComponent(fileName)}`

  await page.goto('about:blank', { waitUntil: 'load', timeout: 120000 })

  const navStart = Date.now()
  const viewResponsePromise = page.waitForResponse(
    (response) => response.url().includes('/tool/excel-editor/view') && response.request().method() === 'GET',
    { timeout: 240000 }
  ).catch(() => null)

  await page.goto(editorUrl, {
    waitUntil: 'domcontentloaded',
    timeout: 240000
  })

  const viewResponse = await viewResponsePromise
  const viewResponseEnd = Date.now()
  const viewStatus = viewResponse ? viewResponse.status() : null

  await page.waitForSelector('.sheet-grid td.sheet-cell textarea, .sheet-grid td.sheet-cell select', {
    timeout: 240000
  })
  const firstCellReady = Date.now()

  await page.waitForFunction(
    () => {
      const loadingMasks = document.querySelectorAll('.el-loading-mask')
      return loadingMasks.length === 0
    },
    { timeout: 240000 }
  ).catch(() => null)

  const uiReady = Date.now()

  const pageStats = await page.evaluate(() => {
    const imageCount = document.querySelectorAll('.sheet-image').length
    const cellCount = document.querySelectorAll('td.sheet-cell').length
    const wrap = document.querySelector('.grid-wrap')
    const scrollMax = wrap ? Math.max(wrap.scrollHeight - wrap.clientHeight, 0) : 0
    const heapMB = typeof performance !== 'undefined' && performance.memory
      ? Math.round((performance.memory.usedJSHeapSize / 1024 / 1024) * 100) / 100
      : null
    return {
      imageCount,
      cellCount,
      scrollMax,
      heapMB
    }
  })

  const scrollProbe = await page.evaluate(async () => {
    const wrap = document.querySelector('.grid-wrap')
    if (!wrap) {
      return null
    }
    const max = Math.max(wrap.scrollHeight - wrap.clientHeight, 0)
    const rounds = 5
    const durations = []
    for (let i = 0; i < rounds; i += 1) {
      const start = performance.now()
      wrap.scrollTop = max
      await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)))
      wrap.scrollTop = 0
      await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)))
      durations.push(performance.now() - start)
    }
    const total = durations.reduce((sum, value) => sum + value, 0)
    const avg = total / durations.length
    const maxDuration = Math.max(...durations)
    return {
      rounds,
      avgMs: Math.round(avg * 100) / 100,
      maxMs: Math.round(maxDuration * 100) / 100
    }
  })

  return {
    round,
    viewStatus,
    navToViewMs: viewResponseEnd - navStart,
    navToFirstCellMs: firstCellReady - navStart,
    navToUiReadyMs: uiReady - navStart,
    pageStats,
    scrollProbe
  }
}

async function runSaveRound(page, round) {
  const selectorA1 = 'td.sheet-cell[data-row-index="0"][data-col-index="0"] textarea'
  const selectorB1 = 'td.sheet-cell[data-row-index="0"][data-col-index="1"] textarea'

  await page.waitForSelector(selectorA1, { timeout: 120000 })
  await page.click(selectorA1)

  const editStart = Date.now()
  await page.fill(selectorA1, `STRESS_${round}_${Date.now()}`)
  await page.click(selectorB1).catch(async () => {
    await page.keyboard.press('Tab')
  })
  const editEnd = Date.now()

  const saveButton = page.locator('button:has-text("在线保存"), button:has-text("保存")').first()
  await saveButton.waitFor({ timeout: 120000 })

  const saveStart = Date.now()
  const saveResponsePromise = page.waitForResponse(
    (response) => response.url().includes('/tool/excel-editor/apply') && response.request().method() === 'POST',
    { timeout: 240000 }
  ).catch(() => null)
  await saveButton.click()
  const saveResponse = await saveResponsePromise
  const saveJson = saveResponse ? await saveResponse.json().catch(() => ({})) : {}
  const saveEnd = Date.now()

  return {
    round,
    editMs: editEnd - editStart,
    saveMs: saveEnd - saveStart,
    responseStatus: saveResponse ? saveResponse.status() : null,
    responseCode: saveJson?.code ?? null
  }
}

function summarizeNumbers(list, field) {
  const values = list
    .map((item) => Number(item?.[field]))
    .filter((value) => Number.isFinite(value))
  if (!values.length) {
    return null
  }
  const total = values.reduce((sum, value) => sum + value, 0)
  return {
    avg: Math.round((total / values.length) * 100) / 100,
    min: Math.min(...values),
    max: Math.max(...values)
  }
}

async function run() {
  const workbookPath = process.argv[2]
  if (!workbookPath) {
    throw new Error('Usage: node scripts/tmp-excel-image-stress.mjs <xlsx_path>')
  }
  if (!fs.existsSync(workbookPath)) {
    throw new Error(`xlsx not found: ${workbookPath}`)
  }

  const workbookStat = fs.statSync(workbookPath)
  const workbookSizeMB = Math.round((workbookStat.size / 1024 / 1024) * 100) / 100

  const token = await loginAndGetToken()
  const browser = await chromium.launch({
    executablePath: pickEdgePath(),
    headless: true,
    args: ['--disable-dev-shm-usage']
  })

  const report = {
    timestamp: nowIso(),
    baseUrl: BASE_URL,
    apiUrl: API_URL,
    workbookPath,
    workbookSizeMB,
    upload: null,
    loadRounds: [],
    saveRounds: [],
    summary: {}
  }

  try {
    const context = await browser.newContext({
      viewport: { width: 1600, height: 960 }
    })
    await context.addCookies([
      {
        name: 'Admin-Token',
        value: token,
        url: BASE_URL
      }
    ])

    report.upload = await uploadWorkbookViaApi(token, workbookPath)
    if (Number(report.upload.responseCode) !== 200) {
      throw new Error(`upload failed: ${JSON.stringify(report.upload)}`)
    }

    const page = await context.newPage()
    console.log(`[stress] uploaded as ${report.upload.fileName}, start load rounds=${LOAD_ROUNDS}`)

    for (let round = 1; round <= LOAD_ROUNDS; round += 1) {
      console.log(`[stress] load round ${round}`)
      const loadResult = await runLoadRound(page, report.upload.fileName, round)
      report.loadRounds.push(loadResult)
    }

    console.log(`[stress] start save rounds=${SAVE_ROUNDS}`)
    for (let round = 1; round <= SAVE_ROUNDS; round += 1) {
      console.log(`[stress] save round ${round}`)
      const saveResult = await runSaveRound(page, round)
      report.saveRounds.push(saveResult)
    }

    await page.close()
    await context.close()
  } finally {
    await browser.close()
  }

  report.summary = {
    uploadMs: report.upload?.uploadMs ?? null,
    navToViewMs: summarizeNumbers(report.loadRounds, 'navToViewMs'),
    navToFirstCellMs: summarizeNumbers(report.loadRounds, 'navToFirstCellMs'),
    navToUiReadyMs: summarizeNumbers(report.loadRounds, 'navToUiReadyMs'),
    saveMs: summarizeNumbers(report.saveRounds, 'saveMs')
  }

  const outFile = path.resolve(process.cwd(), `scripts/excel-image-stress-report-${Date.now()}.json`)
  fs.writeFileSync(outFile, `${JSON.stringify(report, null, 2)}\n`, 'utf8')

  console.log(JSON.stringify({
    ok: true,
    reportFile: outFile,
    summary: report.summary,
    workbookSizeMB,
    uploadedFileName: report.upload?.fileName || ''
  }, null, 2))
}

run().catch((error) => {
  console.error(JSON.stringify({
    ok: false,
    message: error?.message || String(error)
  }, null, 2))
  process.exitCode = 1
})
