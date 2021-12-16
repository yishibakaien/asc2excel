const fs = require('fs')
const xlsx = require('node-xlsx')
const path = require('path')

const sourceFilesDir = path.join(__dirname, './source')
const outputFileDir = path.join(__dirname, './output/')

const files = fs.readdirSync(sourceFilesDir)

files.forEach((fileName) => {
  const pwd = path.resolve(sourceFilesDir, fileName)
  const fileText = fs.readFileSync(pwd, 'utf8').toString()

  let lines = fileText.split('\n')

  lines = lines.slice(3)

  lines = lines.map((line, index) => {
    let x = line.split(' ')
    let prefix = x.slice(0, 6)
    let suffix = x.slice(6, 13).join(' ')
    prefix.push(suffix)
    prefix.unshift(index + 1)
    return prefix
  })

  const data = [
    {
      data: [
        [
          '序号',
          '时间标识',
          '源通道',
          '帧ID',
          'CAN类型',
          '方向',
          '长度',
          '数据',
        ],
        ...lines,
      ],
    },
  ]

  const buffer = xlsx.build(data)

  if (!fs.existsSync(outputFileDir)) {
    fs.mkdirSync(outputFileDir)
  }

  fs.writeFileSync(`${outputFileDir}${fileName}.xls`, buffer)
})
