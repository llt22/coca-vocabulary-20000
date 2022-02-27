const XLSX = require("xlsx");
const fs = require('fs')

let count = 0
let part = 0
let fileName
let writeStream

fs.readdir('excel', (err, files) => {
    files.forEach(file => {
        const workbook = XLSX.readFile(`excel/${file}`)
        const sheetName = workbook.SheetNames[0]
        const worksheet = workbook.Sheets[sheetName]
        const items = XLSX.utils.sheet_to_json(worksheet, { header: 1 })

        items.forEach((item) => {
            const word = item[0]
            const phonetic = item[1]
            const paraphrase = item[2]
            if (count % 200 === 0) {
                part += 1
                fileName = `vocabulary/part${part}_${count + 1}-${count + 200}.md`
                writeStream = fs.createWriteStream(fileName, {
                    flags: 'a' // 'a' means appending (old data will be preserved)
                })

            }

            const r_word = `${count + 1} ${word} \r\n`
            const r_phonetic = `- ${phonetic} \r\n`
            const r_paraphrase = `- ${paraphrase} \r\n`
            const r_d_ren = `- [人人词典](https://www.91dict.com/words?w=${word}) `
            const r_d_ke = `[柯林斯](https://www.collinsdictionary.com/zh/dictionary/english/${word}) `
            const r_d_lan = `[朗文](https://www.ldoceonline.com/dictionary/${word}) `
            const end = '\r\n\r\n'


            writeStream.write(r_word + r_phonetic + r_paraphrase + r_d_ren + r_d_ke + r_d_lan + end)
            count++
        })

    });
});
