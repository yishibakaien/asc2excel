const fs = require('fs')
const xlsx = require('node-xlsx')
const path = require('path')

const sourceFilesDir = path.join(__dirname, './source/')
const outputFileDir = path.join(__dirname, './output/')

const files = fs.readdirSync(sourceFilesDir)

files.forEach((fileName) => {
  const pwd = path.resolve(sourceFilesDir, fileName)
  const fileString = fs.readFileSync(pwd, 'utf8').toString()

  const lines = fileString.split('\n')

  // 前三行为文件描述
  const fileDescLines = lines.slice(0, 3).map((line) => [line])

  const contentLines = lines.slice(3).map((line, index) => {
    const frontPart = line.split(' ').slice(0, 6)
    const endPart = line.slice(6)

    // 这里序号转为 String 可以让序号内容在单元格中左对齐
    return [String(index + 1), ...frontPart, endPart]
  })

  const data = [
    {
      data: [
        ...fileDescLines,
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
        ...contentLines,
      ],
    },
  ]

  const buffer = xlsx.build(data)

  if (!fs.existsSync(outputFileDir)) {
    fs.mkdirSync(outputFileDir)
  }

  fs.writeFileSync(`${outputFileDir}${fileName}.xls`, buffer)
})
