// 转base64编码
function base64(str) {
  return window.btoa(unescape(encodeURIComponent(str)))
}
// 生成xls
function createXls({ headers, data, worksheet }) {
  let str = '<tr>'
  for(let item in headers) {
    str += `<td>${headers[item]}</td>`
  }
  str += '</tr>'
  for(let i = 0; i < data.length; i++) { 
    str += '<tr>'
    for(let item in data[i]) {
      str += `<td style="mso-number-format:'\@';">${ data[i][item] + '\t'}</td>`
    }
    str += '</tr>'
  }
  const template = 
  `<html
    xmlns:o="urn:schemas-microsoft-com:office:office" 
    xmlns:x="urn:schemas-microsoft-com:office:excel" 
    xmlns="http://www.w3.org/TR/REC-html40"
  >
    <head>
      <meta charset="utf-8">
      <!--[if gte mso 9]>
        <xml>
          <x:ExcelWorkbook>
            <x:ExcelWorksheets>
              <x:ExcelWorksheet>
                <x:Name>${worksheet}</x:Name>
                <x:WorksheetOptions>
                  <x:DisplayGridlines/>
                </x:WorksheetOptions>
              </x:ExcelWorksheet>
            </x:ExcelWorksheets>
          </x:ExcelWorkbook>
        </xml>
      <![endif]-->
    </head>
    <body>
      <table>${str}</table>
    </body>
  </html>`
  const uri = 'data:application/vnd.ms-excel;base64,'
  window.location.href = uri + base64(template)
}

// 生成csv
function createCsv({ headers, data, worksheet }) {
  let str = ''
  for(let item in headers) {
    str += `${headers[item]},`
  }
  str += '\n'
  for(let i = 0; i < data.length; i++ ) {
    for(let item in data[i]) {
      str += `${data[i][item]}\t,`
    }
    str += '\n'
  }
  let uri = 'data:text/csv;charset=utf-8,\ufeff' + encodeURIComponent(str)
  let link = document.createElement('a')
  link.href = uri
  link.download = `${worksheet}.csv`
  document.body.appendChild(link)
  link.click()
  document.body.removeChild(link)
}

// 下载excel
function download(type) {
  const headers = ['SKU', '商品详情', '价格'] // 表头
  const worksheet = '商品列表'  // 表名
  const data = [
    {
      sku: 2801010031508,
      title: 'HUAWEI MateBook X Pro',
      price: 9999
    },
    {
      sku: 2801010028201,
      title: 'HUAWEI MateBook 14',
      price: 5888
    },
    {
      sku: 2101010000413,
      title: 'HUAWEI MateBook D 15',
      price: 4099
    }
  ] // 模拟数据
  switch (type) {
    case 'xls':
      createXls({ headers, data, worksheet })
      break
    case 'csv':
      createCsv({ headers, data, worksheet })
      break
  }
}
