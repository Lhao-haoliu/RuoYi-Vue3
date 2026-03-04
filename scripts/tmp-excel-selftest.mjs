import fs from 'node:fs'
import { createServer } from 'vite'
import { chromium } from 'playwright-core'

const EDGE_PATHS = [
  'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
  'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe'
]

function pickEdgePath() {
  for (const candidate of EDGE_PATHS) {
    if (fs.existsSync(candidate)) {
      return candidate
    }
  }
  throw new Error('Microsoft Edge executable not found')
}

function buildWorkbookViewData() {
  const rowCount = 20
  const maxColumnCount = 8
  const valueMap = new Map([
    ['0:0', 'Cat'],
    ['1:0', 'catalog'],
    ['2:0', 'cat'],
    ['3:0', 'dog'],
    ['0:1', '10'],
    ['1:1', '0.25'],
    ['2:1', '2024-01-02'],
    ['3:1', '2'],
    ['0:3', '1'],
    ['1:3', '2']
  ])

  const rows = []
  for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
    const cells = []
    for (let colIndex = 0; colIndex < maxColumnCount; colIndex += 1) {
      const key = `${rowIndex}:${colIndex}`
      const value = valueMap.get(key) || ''
      cells.push({
        rowIndex,
        colIndex,
        rowSpan: 1,
        colSpan: 1,
        value,
        displayValue: value,
        style: {}
      })
    }
    rows.push({
      rowIndex,
      heightPx: 32,
      cells
    })
  }

  return {
    currentFileName: 'selftest.xlsx',
    fileName: 'selftest.xlsx',
    version: '1',
    sheets: [
      {
        name: 'Sheet1',
        rowCount,
        maxColumnCount,
        rowHeights: Array.from({ length: rowCount }, () => 32),
        columnWidths: Array.from({ length: maxColumnCount }, () => 108),
        originalRowHeights: Array.from({ length: rowCount }, () => 32),
        originalColumnWidths: Array.from({ length: maxColumnCount }, () => 108),
        rows,
        validations: [],
        images: [],
        freezePane: {
          xSplit: 0,
          ySplit: 0,
          leftColumn: 0,
          topRow: 0
        },
        filterConfig: {
          enabled: false,
          hasHeader: true,
          headerRow: 0,
          startCol: 0,
          endCol: 1,
          dataStartRow: 1,
          dataEndRow: 3,
          criteria: {}
        }
      }
    ]
  }
}

async function mockApi(page) {
  const workbookView = buildWorkbookViewData()
  await page.route('**/dev-api/**', async (route) => {
    const url = new URL(route.request().url())
    const path = url.pathname
    const method = route.request().method().toUpperCase()
    const json = (payload) => route.fulfill({
      status: 200,
      contentType: 'application/json',
      body: JSON.stringify(payload)
    })

    if (path.endsWith('/getInfo')) {
      return json({
        code: 200,
        user: {
          userId: 1,
          userName: 'selftest',
          nickName: 'Self Test',
          avatar: ''
        },
        roles: ['admin'],
        permissions: ['*:*:*'],
        isDefaultModifyPwd: false,
        isPasswordExpired: false
      })
    }

    if (path.endsWith('/getRouters')) {
      return json({ code: 200, data: [] })
    }

    if (path.endsWith('/tool/excel-editor/info')) {
      return json({
        code: 200,
        data: {
          currentFileName: 'selftest.xlsx'
        }
      })
    }

    if (path.endsWith('/tool/excel-editor/lock')) {
      if (method === 'GET') {
        return json({
          code: 200,
          data: {
            fileName: 'selftest.xlsx',
            ownerUsername: 'selftest',
            expiresAt: Date.now() + 60000,
            self: true
          }
        })
      }
      return json({
        code: 200,
        data: {
          fileName: 'selftest.xlsx',
          ownerUsername: 'selftest',
          expiresAt: Date.now() + 60000,
          self: true
        }
      })
    }

    if (path.endsWith('/tool/excel-editor/lock/heartbeat')) {
      return json({
        code: 200,
        data: {
          fileName: 'selftest.xlsx',
          ownerUsername: 'selftest',
          expiresAt: Date.now() + 60000,
          self: true
        }
      })
    }

    if (path.endsWith('/tool/excel-editor/lock/release')) {
      return json({ code: 200, msg: 'ok', data: {} })
    }

    if (path.endsWith('/tool/excel-editor/view')) {
      return json({
        code: 200,
        data: workbookView
      })
    }

    if (path.endsWith('/tool/excel-editor/apply')) {
      return json({
        code: 200,
        msg: '保存成功',
        data: {
          currentFileName: 'selftest.xlsx',
          fileName: 'selftest.xlsx',
          version: '2'
        }
      })
    }

    if (path.endsWith('/tool/excel-editor/content')) {
      return route.fulfill({
        status: 200,
        contentType: 'application/octet-stream',
        body: 'selftest'
      })
    }

    return json({ code: 200, data: {} })
  })
}

async function run() {
  const results = []
  const server = await createServer({
    root: process.cwd(),
    logLevel: 'error',
    server: {
      host: '127.0.0.1',
      port: 5173
    }
  })
  await server.listen()

  let browser
  try {
    browser = await chromium.launch({
      executablePath: pickEdgePath(),
      headless: true
    })
    const context = await browser.newContext({
      viewport: { width: 1440, height: 900 }
    })
    await context.addCookies([
      {
        name: 'Admin-Token',
        value: 'selftest-token',
        domain: '127.0.0.1',
        path: '/'
      }
    ])
    const page = await context.newPage()
    await mockApi(page)

    await page.goto('http://127.0.0.1:5173/excel-editor/index', { waitUntil: 'domcontentloaded' })
    await page.waitForSelector('.excel-editor-page', { timeout: 120000 })
    await page.waitForSelector('td.sheet-cell[data-row-index="0"][data-col-index="0"] .cell-input', { timeout: 120000 })

    const cell = (rowIndex, colIndex) => page.locator(`td.sheet-cell[data-row-index="${rowIndex}"][data-col-index="${colIndex}"]`)
    const input = (rowIndex, colIndex) => cell(rowIndex, colIndex).locator('.cell-input')
    const getCellValue = async (rowIndex, colIndex) => (await input(rowIndex, colIndex).inputValue()).trim()

    const initialValue = await getCellValue(0, 0)
    results.push({
      step: 'load-workbook',
      ok: initialValue === 'Cat',
      detail: `A1=${initialValue}`
    })

    await cell(0, 0).click()
    const textColorTrigger = page.locator('.style-toolbar .style-color-picker').first().locator('.el-color-picker__trigger')
    await textColorTrigger.click()
    const dropdown = page.locator('.el-color-dropdown:visible').last()
    const colorInput = dropdown.locator('input').last()
    await colorInput.fill('#ff0000')
    await dropdown.getByRole('button', { name: '确定', exact: true }).click()
    await page.waitForTimeout(120)
    const textColorStyle = await input(0, 0).evaluate((el) => getComputedStyle(el).color)
    results.push({
      step: 'text-color',
      ok: /rgb\(\s*255,\s*0,\s*0\s*\)/i.test(textColorStyle),
      detail: textColorStyle
    })

    await page.getByRole('button', { name: '格式刷' }).click()
    await cell(0, 1).click()
    const b1Color = await input(0, 1).evaluate((el) => getComputedStyle(el).color)
    results.push({
      step: 'format-painter',
      ok: /rgb\(\s*255,\s*0,\s*0\s*\)/i.test(b1Color),
      detail: b1Color
    })

    await cell(0, 0).click()
    await page.keyboard.down('Shift')
    await cell(1, 1).click()
    await page.keyboard.up('Shift')
    await page.getByRole('button', { name: '全边框' }).click()
    const borderTopWidth = await cell(0, 0).evaluate((el) => getComputedStyle(el).borderTopWidth)
    results.push({
      step: 'border-tool',
      ok: borderTopWidth !== '0px',
      detail: `borderTopWidth=${borderTopWidth}`
    })

    await cell(1, 1).click()
    await page.getByRole('button', { name: '百分比' }).click()
    const b2Percent = await getCellValue(1, 1)
    await page.getByRole('button', { name: '增小数' }).click()
    const b2Inc = await getCellValue(1, 1)
    await page.getByRole('button', { name: '减小数' }).click()
    const b2Dec = await getCellValue(1, 1)
    await cell(2, 1).click()
    await page.getByRole('button', { name: '日期' }).click()
    const b3Date = await getCellValue(2, 1)
    results.push({
      step: 'number-format',
      ok: b2Percent.includes('%') && b2Inc !== b2Dec && /^\d{4}-\d{2}-\d{2}$/.test(b3Date),
      detail: JSON.stringify({ b2Percent, b2Inc, b2Dec, b3Date })
    })

    await page.getByRole('button', { name: '冻结首行' }).click()
    const stickyAfterFreeze = await cell(0, 0).evaluate((el) => getComputedStyle(el).position)
    await page.getByRole('button', { name: '取消冻结' }).click()
    const stickyAfterClear = await cell(0, 0).evaluate((el) => getComputedStyle(el).position)
    results.push({
      step: 'freeze-pane',
      ok: stickyAfterFreeze === 'sticky' && stickyAfterClear !== 'sticky',
      detail: `freeze=${stickyAfterFreeze}, clear=${stickyAfterClear}`
    })

    await cell(0, 0).click()
    await page.keyboard.down('Shift')
    await cell(1, 0).click()
    await page.keyboard.up('Shift')
    await page.getByRole('button', { name: '查找替换' }).click()
    const findDialog = page.locator('.el-dialog:visible').filter({ hasText: '查找与替换' }).first()
    await findDialog.locator('input[placeholder="输入要查找的内容"]').fill('cat')
    await findDialog.locator('input[placeholder="输入替换内容"]').fill('doggo')
    const previewTextBefore = await findDialog.locator('.el-tag').innerText()
    await findDialog.locator('.el-checkbox').filter({ hasText: '整词匹配' }).click()
    const previewTextWholeWord = await findDialog.locator('.el-tag').innerText()
    await findDialog.locator('.el-checkbox').filter({ hasText: '区分大小写' }).click()
    const previewTextCase = await findDialog.locator('.el-tag').innerText()
    await findDialog.locator('.el-checkbox').filter({ hasText: '仅当前选区' }).click()
    const previewTextSelection = await findDialog.locator('.el-tag').innerText()
    const replaceButton = findDialog.getByRole('button', { name: '全部替换' })
    if (await replaceButton.isDisabled()) {
      const matchCaseBox = findDialog.locator('.el-checkbox').filter({ hasText: '区分大小写' }).first()
      const checked = await matchCaseBox.evaluate((el) => el.classList.contains('is-checked'))
      if (checked) {
        await matchCaseBox.click()
      }
    }
    if (await replaceButton.isDisabled()) {
      const wholeWordBox = findDialog.locator('.el-checkbox').filter({ hasText: '整词匹配' }).first()
      const checked = await wholeWordBox.evaluate((el) => el.classList.contains('is-checked'))
      if (checked) {
        await wholeWordBox.click()
      }
    }
    await replaceButton.click()
    await findDialog.getByRole('button', { name: '关闭', exact: true }).last().click()
    const a1AfterReplace = await getCellValue(0, 0)
    const a2AfterReplace = await getCellValue(1, 0)
    const a3AfterReplace = await getCellValue(2, 0)
    results.push({
      step: 'find-replace',
      ok: a1AfterReplace === 'doggo' && a2AfterReplace === 'catalog' && a3AfterReplace === 'cat',
      detail: JSON.stringify({
        previewTextBefore,
        previewTextWholeWord,
        previewTextCase,
        previewTextSelection,
        a1AfterReplace,
        a2AfterReplace,
        a3AfterReplace
      })
    })

    await cell(0, 0).click()
    await page.keyboard.down('Shift')
    await cell(3, 1).click()
    await page.keyboard.up('Shift')
    const headerSwitch = page.locator('.style-toolbar .el-switch').first()
    const switchChecked = await headerSwitch.evaluate((el) => el.classList.contains('is-checked'))
    if (switchChecked) {
      await headerSwitch.click()
    }
    await page.getByRole('button', { name: '启用筛选' }).click()
    await page.getByRole('button', { name: '列筛选' }).click()
    const filterDialog = page.locator('.el-dialog:visible').filter({ hasText: '列筛选' }).first()
    const candidateOptions = filterDialog.locator('.filter-value-box .el-checkbox')
    const candidateCount = await candidateOptions.count()
    await filterDialog.getByRole('button', { name: '清空' }).click()
    if (candidateCount > 0) {
      await candidateOptions.first().click()
    }
    await filterDialog.getByRole('button', { name: '应用筛选' }).click()
    await page.waitForTimeout(150)
    const visibleRows = await page.locator('td.row-index').allInnerTexts()
    const firstFourVisible = visibleRows.filter((text) => ['1', '2', '3', '4'].includes(text.trim())).length
    results.push({
      step: 'filter-no-header-first-column',
      ok: candidateCount > 0 && firstFourVisible < 4,
      detail: JSON.stringify({ candidateCount, firstFourVisible, visibleRows: visibleRows.slice(0, 8) })
    })

    await page.getByRole('button', { name: '清除筛选' }).click()
    await cell(0, 3).click()
    await page.keyboard.down('Shift')
    await cell(1, 3).click()
    await page.keyboard.up('Shift')
    const fillHandle = cell(1, 3).locator('.fill-handle')
    const targetCell = cell(3, 3)
    const handleBox = await fillHandle.boundingBox()
    const targetBox = await targetCell.boundingBox()
    if (handleBox && targetBox) {
      await page.mouse.move(handleBox.x + handleBox.width / 2, handleBox.y + handleBox.height / 2)
      await page.mouse.down()
      await page.mouse.move(targetBox.x + targetBox.width / 2, targetBox.y + targetBox.height / 2)
      await page.mouse.up()
    }
    await page.waitForTimeout(180)
    const d3Value = await getCellValue(2, 3)
    const d4Value = await getCellValue(3, 3)
    results.push({
      step: 'fill-handle',
      ok: d3Value === '3' && d4Value === '4',
      detail: JSON.stringify({ d3Value, d4Value })
    })

    const gridMetrics = await page.locator('.grid-wrap').evaluate((el) => {
      const style = getComputedStyle(el)
      return {
        minHeight: style.minHeight,
        maxHeight: style.maxHeight
      }
    })
    const gridMinHeight = Number.parseFloat(gridMetrics.minHeight || '0')
    results.push({
      step: 'grid-size',
      ok: Number.isFinite(gridMinHeight) && gridMinHeight >= 620,
      detail: JSON.stringify(gridMetrics)
    })

    console.log(JSON.stringify({ ok: results.every((item) => item.ok), results }, null, 2))
  } finally {
    if (browser) {
      await browser.close()
    }
    await server.close()
  }
}

run().catch((error) => {
  console.error(error)
  process.exit(1)
})
