import { useState, useRef } from 'react'
import * as XLSX from 'xlsx'
import { HyperFormula } from 'hyperformula'
import './App.css'

function App() {
  const [files, setFiles] = useState([])
  const [sheet1Name, setSheet1Name] = useState('Sheet1_見積情報処理用')
  const [sheet2Name, setSheet2Name] = useState('アクセスへコピペ用　Sheet2(案件情報）')
  const [output1Name, setOutput1Name] = useState('sales_data_test_meisai2.csv')
  const [output2Name, setOutput2Name] = useState('sales_data_test_mm2.csv')
  const [results, setResults] = useState(null)
  const [warnings, setWarnings] = useState([])
  const [isDragging, setIsDragging] = useState(false)
  const fileInputRef = useRef(null)

  const handleDragOver = (e) => {
    e.preventDefault()
    setIsDragging(true)
  }

  const handleDragLeave = (e) => {
    e.preventDefault()
    setIsDragging(false)
  }

  const handleDrop = (e) => {
    e.preventDefault()
    setIsDragging(false)
    const droppedFiles = Array.from(e.dataTransfer.files)
    processFiles(droppedFiles)
  }

  const handleFileSelect = (e) => {
    const selectedFiles = Array.from(e.target.files)
    processFiles(selectedFiles)
  }

  const processFiles = (newFiles) => {
    const errors = []

    newFiles.forEach((file) => {
      const ext = file.name.toLowerCase().split('.').pop()
      if (ext === 'xlsx' || ext === 'xls') {
        const reader = new FileReader()
        reader.onload = (e) => {
          try {
            const data = new Uint8Array(e.target.result)
            const workbook = XLSX.read(data, {
              type: 'array',
              cellFormula: true,
              cellHTML: false,
              cellText: true,
              cellDates: true,
              cellNF: false,
              cellStyles: false
            })
            const fileData = {
              file,
              name: file.name,
              size: file.size,
              sheetCount: workbook.SheetNames.length,
              workbook,
              error: null
            }
            setFiles((prev) => [...prev, fileData])
          } catch (error) {
            const fileData = {
              file,
              name: file.name,
              size: file.size,
              sheetCount: 0,
              workbook: null,
              error: '読取エラー'
            }
            setFiles((prev) => [...prev, fileData])
          }
        }
        reader.readAsArrayBuffer(file)
      } else {
        errors.push(`${file.name}: Excel以外のファイル形式は対応していません`)
      }
    })

    if (errors.length > 0) {
      alert(errors.join('\n'))
    }
  }

  const removeFile = (index) => {
    setFiles((prev) => prev.filter((_, i) => i !== index))
  }

  const formatFileSize = (bytes) => {
    if (bytes < 1024) return bytes + ' B'
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB'
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB'
  }

  // シートからフォーマット済みテキストを取得する関数
  // isSheet1: sheet1固有の加工処理を適用するかどうか
  const sheetToFormattedArray = (sheet, fileName = '', isSheet1 = false) => {
    const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1')
    // sheet1はA〜Q列（17列、インデックス0〜16）のみ使用
    const maxCol = isSheet1 ? Math.min(range.e.c, 16) : range.e.c
    const result = []

    for (let row = range.s.r; row <= range.e.r; row++) {
      const rowData = []
      for (let col = range.s.c; col <= maxCol; col++) {
        const cellAddress = XLSX.utils.encode_cell({ r: row, c: col })
        const cell = sheet[cellAddress]

        if (cell) {
          const formattedValue = XLSX.utils.format_cell(cell)
          let cellValue = formattedValue || ''

          // <<と>>を含むセルは空白に変換
          if (cellValue.includes('<<') && cellValue.includes('>>')) {
            cellValue = ''
          }

          // sheet1固有の加工
          if (isSheet1) {
            // H列（商品組番号、インデックス7）が0の場合は空白に変換
            if (col === 7 && (cellValue === '0' || cellValue === 0)) {
              cellValue = ''
            }

            // V列（インデックス21）の値は削除
            if (col === 21) {
              cellValue = ''
            }
          }

          rowData.push(cellValue)
        } else {
          rowData.push('')
        }
      }
      result.push(rowData)
    }

    // sheet1固有の後処理：J列（インデックス9）が空白の行は、B列（インデックス1）も空白にする
    if (isSheet1) {
      result.forEach((row, index) => {
        if (index === 0) return

        const jValue = row[9]
        const jIsEmpty = !jValue || String(jValue).trim() === '' || String(jValue).trim() === '0'

        if (jIsEmpty && row[1]) {
          row[1] = ''
        }
      })
    }

    return result
  }

  // セルから数値を取得する（数式セルの場合は再帰的に解決）
  const getCellNumericValue = (refCell, workbook, depth = 0, sourceSheet = '') => {
    if (depth > 10) return 0 // 無限再帰防止

    // 数値型でキャッシュがある場合
    if (refCell.t === 'n' && refCell.v != null) {
      return refCell.v
    }
    // キャッシュ値をformat_cellで取得して数値変換を試みる
    const formatted = XLSX.utils.format_cell(refCell)
    if (formatted && formatted.trim() !== '') {
      const num = Number(formatted.replace(/,/g, ''))
      if (!isNaN(num)) return num
    }
    // 生の値がある場合
    if (refCell.v !== undefined && refCell.v !== null) {
      const num = Number(refCell.v)
      if (!isNaN(num)) return num
    }
    // 数式セルの場合、数式を手動で解決してから数値変換
    if (refCell.f) {
      const resolved = resolveFormulaManually({ f: refCell.f, v: refCell.v, t: refCell.t }, workbook, depth + 1, sourceSheet)
      if (resolved && resolved.trim() !== '') {
        const num = Number(resolved.replace(/,/g, ''))
        if (!isNaN(num)) return num
      }
    }
    return 0
  }

  // シート内の指定範囲の数値を直接合計する（数式セルは再帰的に解決）
  const sumRange = (workbook, sheetName, startRef, endRef, depth = 0) => {
    if (depth > 10) return 0
    const refSheet = workbook.Sheets[sheetName]
    if (!refSheet) return 0
    const start = XLSX.utils.decode_cell(startRef)
    const end = XLSX.utils.decode_cell(endRef)
    let total = 0
    for (let r = start.r; r <= end.r; r++) {
      for (let c = start.c; c <= end.c; c++) {
        const addr = XLSX.utils.encode_cell({ r, c })
        const cell = refSheet[addr]
        if (cell) {
          total += getCellNumericValue(cell, workbook, depth, sheetName)
        }
      }
    }
    return total
  }

  // HyperFormulaで計算できなかった数式を手動で解決する
  // sourceSheet: この数式が属するシート名（同一シート内参照の解決に使用）
  const resolveFormulaManually = (cell, workbook, depth = 0, sourceSheet = '') => {
    if (!cell || !cell.f) return ''
    if (depth > 10) return '' // 無限再帰防止

    const formula = cell.f

    // TEXT(SheetName!CellRef, "format") パターン
    const textMatch = formula.match(/^TEXT\((?:'([^']+)'|([^!]+))!([A-Z]+\d+)\s*,\s*"([^"]+)"\)$/i)
    if (textMatch) {
      const sheetName = textMatch[1] || textMatch[2]
      const cellRef = textMatch[3]
      const formatStr = textMatch[4]
      const refSheet = workbook.Sheets[sheetName]
      if (refSheet && refSheet[cellRef]) {
        const refCell = refSheet[cellRef]
        // 日付シリアル値をJSのDateに変換してフォーマット
        if (refCell.t === 'n' && formatStr.toLowerCase().includes('yy')) {
          const date = XLSX.SSF.parse_date_code(refCell.v)
          if (date) {
            const y = String(date.y)
            const m = String(date.m).padStart(2, '0')
            const d = String(date.d).padStart(2, '0')
            return formatStr
              .replace(/yyyy/i, y)
              .replace(/yy/i, y.slice(-2))
              .replace(/mm/, m)
              .replace(/dd/i, d)
          }
        }
        // 日付型セルの場合
        if (refCell.t === 'd' && refCell.v instanceof Date) {
          const dt = refCell.v
          const y = String(dt.getFullYear())
          const m = String(dt.getMonth() + 1).padStart(2, '0')
          const d = String(dt.getDate()).padStart(2, '0')
          return formatStr
            .replace(/yyyy/i, y)
            .replace(/yy/i, y.slice(-2))
            .replace(/mm/, m)
            .replace(/dd/i, d)
        }
        // それ以外はformat_cellで取得
        return XLSX.utils.format_cell(refCell) || ''
      }
    }

    // IF(OR(SheetName!Cell="",SheetName!Cell=0),"",SheetName!Cell) パターン
    const ifOrMatch = formula.match(/^IF\(OR\((?:'([^']+)'|([^!]+))!([A-Z]+\d+)=""\s*,\s*(?:'([^']+)'|([^!]+))!([A-Z]+\d+)=0\)\s*,\s*""\s*,\s*(?:'([^']+)'|([^!]+))!([A-Z]+\d+)\)$/i)
    if (ifOrMatch) {
      const sheetName = ifOrMatch[1] || ifOrMatch[2]
      const cellRef = ifOrMatch[9]
      const refSheet = workbook.Sheets[sheetName]
      if (refSheet && refSheet[cellRef]) {
        const refCell = refSheet[cellRef]
        const val = getCellNumericValue(refCell, workbook, depth + 1, sheetName)
        if (val === 0 || val === '') return ''
        return String(val)
      }
      return ''
    }

    // IFERROR(INDEX(SheetName!Range, SMALL(IF(SheetName!Range<>"", ROW(SheetName!Range)), ROW(R?))-offset), "")
    // → 空白でないセルを順番に取り出す配列数式。直接元の範囲を合計するSUM用途では、
    //   元の範囲のデータをそのまま返せばよい。
    //   ここでは単一セルの値として解決するため、数式内のINDEXの参照先範囲とROW(Rn)から
    //   該当する値を取得する
    const iferrorIndexMatch = formula.match(/^IFERROR\(INDEX\((?:'([^']+)'|([^!]+))!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)\s*,\s*SMALL\(IF\(.*?\)\s*,\s*ROW\(R(\d+)\)\)-(\d+)\)\s*,\s*""\)$/i)
    if (iferrorIndexMatch) {
      const sheetName = iferrorIndexMatch[1] || iferrorIndexMatch[2]
      const startCol = iferrorIndexMatch[3]
      const startRow = parseInt(iferrorIndexMatch[4])
      const endCol = iferrorIndexMatch[5]
      const endRow = parseInt(iferrorIndexMatch[6])
      const rowNum = parseInt(iferrorIndexMatch[7]) // ROW(Rn)のn
      const offset = parseInt(iferrorIndexMatch[8]) // 減算するオフセット
      const refSheet = workbook.Sheets[sheetName]
      if (refSheet) {
        // 空白でないセルを順番に集める
        const nonEmptyValues = []
        const startC = XLSX.utils.decode_col(startCol)
        const endC = XLSX.utils.decode_col(endCol)
        for (let r = startRow - 1; r <= endRow - 1; r++) {
          for (let c = startC; c <= endC; c++) {
            const addr = XLSX.utils.encode_cell({ r, c })
            const refCell = refSheet[addr]
            if (refCell) {
              const val = getCellNumericValue(refCell, workbook, depth + 1, sheetName)
              if (val !== 0 && val !== '') {
                nonEmptyValues.push(val)
              }
            }
          }
        }
        // SMALL + ROW(Rn) - offset でn番目の非空白値を取得
        const idx = rowNum - offset
        if (idx >= 0 && idx < nonEmptyValues.length) {
          return String(nonEmptyValues[idx])
        }
        return '' // IFERROR → 空文字
      }
      return ''
    }

    // SUM(SheetName!Range) パターン
    const sumMatch = formula.match(/^SUM\((?:'([^']+)'|([^!]+))!([A-Z]+\d+):([A-Z]+\d+)\)$/i)
    if (sumMatch) {
      const sheetName = sumMatch[1] || sumMatch[2]
      const startRef = sumMatch[3]
      const endRef = sumMatch[4]
      return String(sumRange(workbook, sheetName, startRef, endRef, depth + 1))
    }

    // 汎用 IF(CellRef=0,"",CellRef) パターン — 同一シート内の参照
    // 例: IF(V2=0,"",V2)
    const ifSimpleMatch = formula.match(/^IF\(([A-Z]+\d+)=0\s*,\s*""\s*,\s*([A-Z]+\d+)\)$/i)
    if (ifSimpleMatch) {
      const cellRef = ifSimpleMatch[2]
      if (sourceSheet) {
        const ws = workbook.Sheets[sourceSheet]
        if (ws && ws[cellRef]) {
          const val = getCellNumericValue(ws[cellRef], workbook, depth + 1)
          if (val === 0) return ''
          return String(val)
        }
      }
      return ''
    }

    // 汎用 IF(SheetName!CellRef=0,"",SheetName!CellRef) パターン — 他シート参照
    const ifSimpleCrossMatch = formula.match(/^IF\((?:'([^']+)'|([^!]+))!([A-Z]+\d+)=0\s*,\s*""\s*,\s*(?:'[^']+'|[^!]+)!([A-Z]+\d+)\)$/i)
    if (ifSimpleCrossMatch) {
      const sheetName = ifSimpleCrossMatch[1] || ifSimpleCrossMatch[2]
      const cellRef = ifSimpleCrossMatch[4]
      const refSheet = workbook.Sheets[sheetName]
      if (refSheet && refSheet[cellRef]) {
        const val = getCellNumericValue(refSheet[cellRef], workbook, depth + 1)
        if (val === 0) return ''
        return String(val)
      }
      return ''
    }

    // フォールバック: SheetJSのキャッシュ値を使う
    if (cell.v !== undefined && cell.v !== null) {
      return XLSX.utils.format_cell(cell) || String(cell.v)
    }
    return ''
  }

  // sheet2専用: HyperFormulaで数式を計算して2次元配列を返す
  const sheetToFormattedArrayWithFormulas = (workbook, targetSheetName, fileName = '') => {
    // シート名を英数字のみのエイリアスにマッピング（HyperFormulaの日本語パース問題を回避）
    const originalNames = workbook.SheetNames
    const aliasMap = {} // originalName -> alias
    const reverseMap = {} // alias -> originalName
    originalNames.forEach((name, i) => {
      const alias = 'S' + i
      aliasMap[name] = alias
      reverseMap[alias] = name
    })

    // 数式内のシート名参照をエイリアスに置換する
    const rewriteFormula = (formula) => {
      let result = formula
      // 長い名前から先に置換（部分一致を防ぐ）
      const sortedNames = [...originalNames].sort((a, b) => b.length - a.length)
      for (const name of sortedNames) {
        const alias = aliasMap[name]
        // 'シート名'! 形式（シングルクォート付き）
        const quotedPattern = "'" + name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + "'"
        result = result.replace(new RegExp(quotedPattern, 'g'), "'" + alias + "'")
        // シート名! 形式（クォートなし）
        const unquotedPattern = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '!'
        result = result.replace(new RegExp(unquotedPattern, 'g'), alias + '!')
      }
      return result
    }

    // 全シートのデータをHyperFormulaに登録
    const hfSheets = {}

    originalNames.forEach((name) => {
      const ws = workbook.Sheets[name]
      const ref = ws['!ref']
      if (!ref) {
        hfSheets[aliasMap[name]] = [[]]
        return
      }
      const range = XLSX.utils.decode_range(ref)
      const rows = []
      for (let r = range.s.r; r <= range.e.r; r++) {
        const row = []
        for (let c = range.s.c; c <= range.e.c; c++) {
          const addr = XLSX.utils.encode_cell({ r, c: c })
          const cell = ws[addr]
          if (cell && cell.f) {
            // 数式セル: エイリアスに書き換えてからHyperFormulaに渡す
            row.push('=' + rewriteFormula(cell.f))
          } else if (cell) {
            row.push(cell.v !== undefined && cell.v !== null ? cell.v : '')
          } else {
            row.push('')
          }
        }
        rows.push(row)
      }
      hfSheets[aliasMap[name]] = rows
    })

    const targetAlias = aliasMap[targetSheetName]

    const hf = HyperFormula.buildFromSheets(hfSheets, { licenseKey: 'gpl-v3' })

    // 対象シートから計算済みの値を取得
    const sheetAliases = Object.keys(hfSheets)
    const targetIndex = sheetAliases.indexOf(targetAlias)
    if (targetIndex === -1) return []

    const ws = workbook.Sheets[targetSheetName]
    const ref = ws['!ref']
    if (!ref) return []
    const range = XLSX.utils.decode_range(ref)
    const result = []

    for (let r = range.s.r; r <= range.e.r; r++) {
      const rowData = []
      for (let c = range.s.c; c <= range.e.c; c++) {
        const addr = XLSX.utils.encode_cell({ r, c: c })
        const cell = ws[addr]
        let cellValue = ''

        if (cell) {
          if (cell.f) {
            // 数式セル: HyperFormula → 手動解決 → キャッシュ値の順で試行
            // 1. HyperFormulaの計算結果を試す
            // まず手動解決を試す（参照先セルの生データを直接読むため信頼性が高い）
            const manual = resolveFormulaManually(cell, workbook)
            if (manual && manual.trim() !== '') {
              cellValue = manual
            } else {
              // 手動解決できない場合はHyperFormulaを試す
              const hfValue = hf.getCellValue({ sheet: targetIndex, row: r, col: c })
              if (hfValue != null && typeof hfValue !== 'object') {
                cellValue = String(hfValue)
              } else {
                // キャッシュ値にフォールバック
                const cached = XLSX.utils.format_cell(cell)
                if (cached && cached.trim() !== '') {
                  cellValue = cached
                } else if (cell.v !== undefined && cell.v !== null) {
                  cellValue = String(cell.v)
                }
              }
            }
          } else {
            // 非数式セル: キャッシュ値をそのまま使用
            const cached = XLSX.utils.format_cell(cell)
            if (cached && cached.trim() !== '') {
              cellValue = cached
            } else if (cell.v !== undefined && cell.v !== null && String(cell.v).trim() !== '') {
              cellValue = String(cell.v)
            }
          }
        }

        // <<と>>を含むセルは空白に変換
        if (cellValue.includes('<<') && cellValue.includes('>>')) {
          cellValue = ''
        }

        rowData.push(cellValue)
      }
      result.push(rowData)
    }

    if (result.length > 1) console.log('[HF Debug] 結果データ行1:', result[1])

    hf.destroy()
    return result
  }

  // 行が実質的に空白かどうかを判定する関数
  // sheet1: 品名（F列、インデックス5）で判定
  // sheet2: 全セルが空白かどうかで判定
  const isRowEmpty = (row, isSheet1 = false) => {
    if (!row || row.length === 0) return true

    if (isSheet1) {
      // sheet1: 品名（インデックス5）をチェック
      const productName = row[5]
      if (productName === undefined || productName === null) return true
      const productNameStr = String(productName).trim()
      return productNameStr === '' || productNameStr === '0' || productName === 0
    } else {
      // sheet2: 全セルが空白かどうかで判定
      return row.every((cell) => {
        if (cell === undefined || cell === null) return true
        return String(cell).trim() === ''
      })
    }
  }

  const mergeSheets = () => {
    if (files.length === 0) {
      alert('ファイルをアップロードしてください')
      return
    }

    const newWarnings = []
    const sheet1Data = []
    const sheet2Data = []
    let sheet1Headers = []
    let sheet2Headers = []

    files.forEach((fileData) => {
      if (fileData.error) {
        newWarnings.push(`${fileData.name}: ファイルの読み取りに失敗しました`)
        return
      }

      const workbook = fileData.workbook

      // Sheet1の処理
      if (workbook.SheetNames.includes(sheet1Name)) {
        const sheet = workbook.Sheets[sheet1Name]
        const jsonData = sheetToFormattedArray(sheet, fileData.name, true)

        if (jsonData.length > 0) {
          if (sheet1Headers.length === 0) {
            sheet1Headers = jsonData[0] || []
          }

          for (let i = 1; i < jsonData.length; i++) {
            if (!isRowEmpty(jsonData[i], true)) {
              sheet1Data.push({
                fileName: fileData.name,
                rowData: jsonData[i]
              })
            }
          }
        }
      } else {
        newWarnings.push(`${fileData.name}: シート「${sheet1Name}」が見つかりません`)
      }

      // Sheet2の処理
      if (workbook.SheetNames.includes(sheet2Name)) {
        const jsonData = sheetToFormattedArrayWithFormulas(workbook, sheet2Name, fileData.name)

        if (jsonData.length > 0) {
          if (sheet2Headers.length === 0) {
            sheet2Headers = jsonData[0] || []
          }

          for (let i = 1; i < jsonData.length; i++) {
            if (!isRowEmpty(jsonData[i], false)) {
              sheet2Data.push({
                fileName: fileData.name,
                rowData: jsonData[i]
              })
            }
          }
        }
      } else {
        newWarnings.push(`${fileData.name}: シート「${sheet2Name}」が見つかりません`)
      }
    })

    setWarnings(newWarnings)

    // 結果を保存
    setResults({
      sheet1: {
        headers: sheet1Headers,
        data: sheet1Data,
        fileName: output1Name
      },
      sheet2: {
        headers: sheet2Headers,
        data: sheet2Data,
        fileName: output2Name
      }
    })
  }

  const buildRows = (sheetKey) => {
    const { headers, data } = results[sheetKey]
    const rows = [[...headers]]
    const isSheet1 = sheetKey === 'sheet1'
    let lastQuoteNumber = ''
    let lastCreateDate = ''

    data.forEach((item) => {
      const row = []
      for (let i = 0; i < headers.length; i++) {
        const cellValue = item.rowData[i] || ''

        if (isSheet1) {
          if (i === 0) {
            if (cellValue && cellValue.trim() !== '') {
              lastQuoteNumber = cellValue
              row.push(cellValue)
            } else {
              row.push(lastQuoteNumber)
            }
          } else if (i === 2) {
            if (cellValue && cellValue.trim() !== '') {
              lastCreateDate = cellValue
              row.push(cellValue)
            } else {
              row.push(lastCreateDate)
            }
          } else {
            row.push(cellValue)
          }
        } else {
          row.push(cellValue)
        }
      }
      rows.push(row)
    })

    return rows
  }

  const rowsToCsv = (rows) => {
    return rows.map((row) =>
      row.map((cell) => {
        const str = String(cell ?? '')
        if (str.includes(',') || str.includes('"') || str.includes('\n')) {
          return '"' + str.replace(/"/g, '""') + '"'
        }
        return str
      }).join(',')
    ).join('\n')
  }

  const downloadCsvFile = (csvString, fileName) => {
    const bom = '\uFEFF'
    const blob = new Blob([bom + csvString], { type: 'text/csv;charset=utf-8;' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = fileName
    document.body.appendChild(a)
    a.click()
    document.body.removeChild(a)
    URL.revokeObjectURL(url)
  }

  const downloadAll = () => {
    if (!results) return

    if (results.sheet1.data.length > 0) {
      const rows = buildRows('sheet1')
      const csv = rowsToCsv(rows)
      downloadCsvFile(csv, results.sheet1.fileName)
    }

    if (results.sheet2.data.length > 0) {
      setTimeout(() => {
        const rows = buildRows('sheet2')
        const csv = rowsToCsv(rows)
        downloadCsvFile(csv, results.sheet2.fileName)
      }, 100)
    }
  }

  return (
    <div className="app">
      <header>
        <h1>Excel シート結合ツール</h1>
        <p>複数のExcelファイルから指定したシートを抽出し、1つのファイルに結合します</p>
      </header>

      <div
        className={`dropzone ${isDragging ? 'dragging' : ''}`}
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onDrop={handleDrop}
        onClick={() => fileInputRef.current?.click()}
      >
        <p>ここにExcelファイルをドラッグ＆ドロップ</p>
        <p>または クリックしてファイルを選択</p>
        <input
          ref={fileInputRef}
          type="file"
          multiple
          accept=".xlsx,.xls"
          onChange={handleFileSelect}
          style={{ display: 'none' }}
        />
      </div>

      {files.length > 0 && (
        <div className="file-list">
          <h2>アップロード済みファイル ({files.length}件)</h2>
          {files.map((fileData, index) => (
            <div key={index} className="file-item">
              <div className="file-info">
                <span className="file-name">{fileData.name}</span>
                <span className="file-meta">
                  {formatFileSize(fileData.size)} / {fileData.sheetCount}シート
                </span>
                {fileData.error && <span className="badge error">{fileData.error}</span>}
              </div>
              <button className="btn-remove" onClick={() => removeFile(index)}>
                削除
              </button>
            </div>
          ))}
        </div>
      )}

      <div className="sheet-config">
        <h2>結合するシート名</h2>
        <div className="input-group">
          <label>
            シート①（見積情報）
            <input
              type="text"
              value={sheet1Name}
              onChange={(e) => setSheet1Name(e.target.value)}
            />
          </label>
        </div>
        <div className="input-group">
          <label>
            シート②（案件情報）
            <input
              type="text"
              value={sheet2Name}
              onChange={(e) => setSheet2Name(e.target.value)}
            />
          </label>
        </div>
      </div>

      <div className="sheet-config">
        <h2>出力ファイル名</h2>
        <div className="input-group">
          <label>
            見積情報ファイル名
            <input
              type="text"
              value={output1Name}
              onChange={(e) => setOutput1Name(e.target.value)}
            />
          </label>
        </div>
        <div className="input-group">
          <label>
            案件情報ファイル名
            <input
              type="text"
              value={output2Name}
              onChange={(e) => setOutput2Name(e.target.value)}
            />
          </label>
        </div>
      </div>

      <button className="btn-merge" onClick={mergeSheets} disabled={files.length === 0}>
        結合する
      </button>

      {warnings.length > 0 && (
        <div className="warnings">
          <h3>警告</h3>
          {warnings.map((warning, index) => (
            <p key={index} className="warning-message">{warning}</p>
          ))}
        </div>
      )}

      {results && (
        <div className="results">
          <h2>結合結果</h2>

          {results.sheet1.data.length > 0 ? (
            <p className="result-info">見積情報: {results.sheet1.fileName} ({results.sheet1.data.length}行)</p>
          ) : (
            <p className="no-data">見積情報: データが見つかりません</p>
          )}

          {results.sheet2.data.length > 0 ? (
            <p className="result-info">案件情報: {results.sheet2.fileName} ({results.sheet2.data.length}行)</p>
          ) : (
            <p className="no-data">案件情報: データが見つかりません</p>
          )}

          {(results.sheet1.data.length > 0 || results.sheet2.data.length > 0) && (
            <button className="btn-download" onClick={downloadAll}>
              ダウンロード
            </button>
          )}
        </div>
      )}
    </div>
  )
}

export default App
