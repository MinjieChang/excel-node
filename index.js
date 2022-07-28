// import xlsx from 'node-xlsx'
const fs = require('fs')
const xlsx = require('node-xlsx').default
const NodeMonkey = require('node-monkey')
const lodash = require('lodash')
const monkey = NodeMonkey(
  {
    server: {
      disableLocalOutput: true,
    },
  },
  'ninja'
)

// Parse a buffer
const template = xlsx.parse(`./files/template.xlsx`)
// Parse a file
const collect = xlsx.parse(`./files/collect.xlsx`)

const temp = template[0]
const coll = collect[0]

const templateData = temp.data
const collectData = coll.data

console.log('template1')
console.dir(templateData, { depth: null })
console.log('collect')
console.dir(collectData, { depth: null })

const hanzhenIdx = [1, 3]
const qiyeNameIdx = [2, 0]
const contentIdx = [3, 0]
const deadTimeIdx = [9, 0]
const qiankuanIdx = [9, 1]
const shoukuanIdx = [9, 2]
const beizhuIdx = [9, 3]
const gongsiIdx = [19, 1]

// tempData.forEach((data) => {
//   const name = data[nameIdx[0]][nameIdx[0]]
// })

const allSheets = []
console.log(collectData, 'collData2')
collectData.forEach((data, idx) => {
  if (idx >= 2) {
    const tempData = lodash.cloneDeep(templateData)
    tempData[qiankuanIdx[0]][qiankuanIdx[1]] = ''
    tempData[shoukuanIdx[0]][shoukuanIdx[1]] = ''

    const hanzhenBianhao = data[1]
    const qiyeName = data[4]
    const deadTime = data[28]
    const beizhu = data[5]
    const qiankuan = data[34]
    const tempDataName = tempData[qiyeNameIdx[0]][qiyeNameIdx[1]].split('：')[0]
    // console.log(tempDataName, 'tempDataName2')
    tempData[hanzhenIdx[0]][hanzhenIdx[1]] = hanzhenBianhao
    tempData[qiyeNameIdx[0]][qiyeNameIdx[1]] = qiyeName
    tempData[contentIdx[0]][contentIdx[1]] = tempData[contentIdx[0]][contentIdx[1]].replace(
      tempDataName,
      qiyeName
    )
    tempData[deadTimeIdx[0]][deadTimeIdx[1]] = deadTime
    tempData[gongsiIdx[0]][gongsiIdx[1]] = qiyeName
    tempData[beizhuIdx[0]][beizhuIdx[1]] = beizhu
    if (['应付账款', '预付款项', '其他应付款'].includes(beizhu)) {
      tempData[qiankuanIdx[0]][qiankuanIdx[1]] = qiankuan
    } else if (['应收账款', '其他应收款', '预收账款'].includes(beizhu)) {
      tempData[shoukuanIdx[0]][shoukuanIdx[1]] = qiankuan
    }
    allSheets.push({ name: '郭老师' + (idx - 1), data: tempData })
  }
})

const data = [
  [1, 2, 3],
  [true, false, null, 'sheetjs'],
  ['foo', 'bar', new Date('2014-02-19T14:30Z'), '0.3'],
  ['baz', null, 'qux'],
]
const sheetOptions = {
  '!cols': [{ wch: 60 }, { wch: 20 }, { wch: 20 }, { wch: 30 }],
}

const buffer = xlsx.build(
  // [
  //   { name: 'mySheetName', data: tempData },
  //   { name: 'mySecondSheet', data: tempData },
  // ],
  allSheets,
  { sheetOptions }
) // Returns a buffer

// fs.writeFile('a.xlsx', buffer, function (err) {
//   if (err) {
//     console.log('Write failed: ' + err)
//     return
//   }
//   console.log('Write completed.')
// })

fs.writeFileSync('a.xlsx', buffer)

const data2 = [
  ['111'],
  [1, 2, 3],
  [true, false, null, 'sheetjs'],
  ['foo', 'bar', new Date('2014-02-19T14:30Z'), '0.3'],
  ['baz', null, 'qux'],
]
const range = { s: { c: 0, r: 0 }, e: { c: 3, r: 0 } } // A1:A4
const sheetOptions3 = {
  '!merges': [range],
  '!cols': [{ wch: 60 }, { wch: 20 }, { wch: 20 }, { wch: 20 }],
  '!rows': [
    { hpt: 1, level: 2 },
    { hpt: 1, level: 2 },
    { hpt: 1, level: 2 },
    { hpt: 14, level: 2 },
    { hpt: 2, level: 2 },
  ],
}
var buffer2 = xlsx.build([{ name: 'mySheetName', data: data2 }], { sheetOptions: sheetOptions3 }) // Returns a buffer
fs.writeFileSync('b.xlsx', buffer2)
