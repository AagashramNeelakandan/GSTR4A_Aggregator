/* eslint-disable prefer-const */
/* eslint-disable eqeqeq */

// App Name: GSTR4A_Aggregater
// Author: Aagashram Neelakandan


function fileValidation(fileInput) {
  var filePath = fileInput.value;

  // Allowing file type
  const allowedExtensions = /(\.xlsx)$/i // (\.doc|\.docx|\.odt|\.pdf|\.tex|\.txt|\.rtf|\.wps|\.wks|\.wpd)$/i;
  if (!allowedExtensions.exec(filePath)) {
    alert('Invalid file type. Please upload .xlsx files')
    fileInput.value = ''
    return false
  }
  return true
}

async function performAggregation (inputFiles) {
  if (fileValidation(inputFiles)) {
    await parseExcelFile2(inputFiles)
  }
}

// Excel Columns Index
const _GET_TRADE_NAME_ROW_ = 3
const _GET_GSTIN_ROW_ = 2
const _GET_RATE_ROW_ = 10
const _GET_TAX_VALUE_ROW_ = 11

async function parseExcelFile2 (inputElement) {
  let modifiedRows = []
  let files = inputElement.files || []
  if (!files.length) return
  let index = 0

  Object.keys(files).forEach(i => {
    const file = files[i]
    const reader = new FileReader()
    reader.onload = (e) => {
      // server call for uploading or reading the files one-by-one
      // by using 'reader.result' or 'file'
      let arrayBuffer = reader.result
      let workbook = new ExcelJS.Workbook()

      workbook.xlsx.load(arrayBuffer).then(function (workbook) {
        console.timeEnd()

        const worksheet = workbook.getWorksheet('B2B')

        if (typeof worksheet == 'undefined') {
          alert('Unable to parse the File. Please choose the correct one.')
          // Deleting the Previous Tables
          document.getElementById('table_div').style.display = 'none'
          document.getElementById('tableBody').innerHTML = ''
          return
        }
        worksheet.eachRow(function (row, rowNumber) {
 
          if (rowNumber >= 7) {
            modifiedRows = loadNParseExcel(row, modifiedRows)
          }
        })


        if (index == files.length - 1) {
          // Sorting Trade Names by Alphabatically
          modifiedRows.sort((a, b) => ((a['Trade/Legal name'] == b['Trade/Legal name']) ? 0 : ((a['Trade/Legal name'] < b['Trade/Legal name']) ? -1 : 1)))
          // Sorting Rates
          modifiedRows.forEach((element) => element['Rate (%)'].sort((a, b) => ((parseInt(Object.keys(a)) < parseInt(Object.keys(b))) ? -1 : 1)))
          return generateTable(modifiedRows)
        }
        index++
      })
    }
    reader.readAsBinaryString(file)
  })
}



function loadNParseExcel (row, dataRows) {
  // console.log('DataRows Before: ' + JSON.stringify(dataRows))

  // Clean up Process
  const cleanUp = (strVal) => (strVal.toString().trim().replace(',', ''))

  // Getting Row values
  const gstinRowVal = cleanUp(row.values[_GET_GSTIN_ROW_])
  const rateRowVal = parseInt(cleanUp(row.values[_GET_RATE_ROW_]))
  const taxRowVal = parseFloat(cleanUp(row.values[_GET_TAX_VALUE_ROW_]))
  const tradeRowVal = cleanUp(row.values[_GET_TRADE_NAME_ROW_])

  const dataRowIndex = dataRows.findIndex(elements => elements['GSTIN of supplier'] == gstinRowVal)
  if (dataRowIndex !== -1) {
    // console.log()
    const rateKeyIndex = dataRows[dataRowIndex]['Rate (%)'].findIndex(index => Object.keys(index)[0] == rateRowVal)
    if (rateKeyIndex !== -1) {
      dataRows[dataRowIndex]['Rate (%)'][rateKeyIndex][rateRowVal] += Math.round(taxRowVal * 10000) / 10000
    } else {
      const pushRates = {}
      pushRates[rateRowVal] = taxRowVal
      dataRows[dataRowIndex]['Rate (%)'].push(pushRates)
    }
  } else {
    const rateObj = {}
    rateObj[rateRowVal] = taxRowVal
    const pushElement = {
      'Trade/Legal name': tradeRowVal,
      'GSTIN of supplier': gstinRowVal,
      'Rate (%)': [rateObj]
    }
    dataRows.push(pushElement)
  }
  return dataRows
}

// For Generating Tables
function generateTable (modifiedRows) {
  // console.log("Modified Rows in GenerateTable: " + JSON.stringify(modifiedRows))

  // If output is none, then raise an alert
  // Unhide Table
  document.getElementById('table_div').style.display = 'block'

  let tableBody = document.getElementById('tableBody')
  // Making the Table empty
  tableBody.innerHTML = ''

  if (modifiedRows == [] || modifiedRows == null) {
    alert('Unable to parse the File. Please choose the correct one.')
    return
  }

  // Final Total
  let finalTotal = 0

  for (let index = 0; index < modifiedRows.length; index++) {
    let rates_count_col = modifiedRows[index]['Rate (%)'].length + 1
    var row = `<tr>
          <td class="tg-pnt8" rowspan="${rates_count_col}">${modifiedRows[index]['Trade/Legal name']}</td>
          <td class="tg-pnt8" rowspan="${rates_count_col}">${modifiedRows[index]['GSTIN of supplier']}</td>`

    let totalRate = 0
    let rateKey = Object.keys(modifiedRows[index]['Rate (%)'][0])[0]
    let val = modifiedRows[index]['Rate (%)'][0][rateKey]

    totalRate += parseFloat(val)

    row += `<td class="tg-pnt8">${rateKey}</td>
                <td class="tg-dg7a">${Math.round(val * 10000) / 10000}</td>`

    for (let rateIndex = 1; rateIndex < modifiedRows[index]['Rate (%)'].length; rateIndex++) {
      let rateKey = Object.keys(modifiedRows[index]['Rate (%)'][rateIndex])[0]
      let val = modifiedRows[index]['Rate (%)'][rateIndex][rateKey]
      totalRate += parseFloat(val)

      if (rateIndex % 2 == 1) {
        row += `<tr>
                      <td class="tg-nrix">${rateKey}</td>
                      <td class="tg-0lax">${Math.round(val * 10000) / 10000}</td>
                    </tr>`
      } else {
        row += `<tr>
                      <td class="tg-pnt8">${rateKey}</td>
                      <td class="tg-dg7a">${Math.round(val * 10000) / 10000}</td>
                    </tr>`
      }
    }
    // Total Value
    row += ` <tr>
                    <td class="tg-wa1i">Total</td>
                    <td class="tg-wa1i">${Math.round(totalRate * 10000) / 10000}</td>
                </tr>`

    tableBody.innerHTML += row

    // Adding Individual Totals
    finalTotal += totalRate
  }

  // Adding Final total row
  var row = `<tr>
                  <td class="tg-wa1i" rowspan="2" colspan="3">TOTAL</td>
                  <td class="tg-wa1i" rowspan="2">${Math.round(finalTotal * 10000) / 10000}</td>
              <tr>`

  tableBody.innerHTML += row
}
