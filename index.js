const fs = require("fs")
const PDFParser = require("pdf2json")
const express = require('express')
const multer = require('multer')
const path = require('path')
const XLSX = require('xlsx')

const upload = multer({ dest: 'uploads/' })

const app = express()

const EXCEL_FILE = 'ijazah.xlsx'

const appendToExcel = (ijazahData) => {
    let workbook, worksheet

    if (fs.existsSync(EXCEL_FILE)) {
        workbook = XLSX.readFile(EXCEL_FILE);
        worksheet = workbook.Sheets[workbook.SheetNames[0]];
        console.log('ada');
    } else {
        workbook = XLSX.utils.book_new()
        worksheet = XLSX.utils.aoa_to_sheet([['No', 'No ijazah', 'NISN', 'NPSN']])
        XLSX.utils.book_append_sheet(workbook, worksheet, "Data")
        console.log('gak ada');
    }

    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 })

    const rowNumber = data.length

    const newRow = [
        rowNumber,
        ijazahData[0].split('%')[0],
        ijazahData[1],
        ijazahData[2]
    ]

    data.push(newRow)

    const newWorksheet = XLSX.utils.aoa_to_sheet(data)
    workbook.Sheets[workbook.SheetNames[0]] = newWorksheet
    XLSX.writeFile(workbook, EXCEL_FILE)
}

app.use(express.static('public'))
app.listen(3000, () => console.log(`listening on ${3000}`))

app.post('/upload', upload.array('file', 10), (req, res) => {
    let ijazahData = []
    req.files.forEach(file => {
        const filePath = file.path
        const pdfParser = new PDFParser()
        let ijazahRow = []

        pdfParser.on("pdfParser_dataError", (errData) =>
            console.error(errData.parserError)

        )

        pdfParser.on("pdfParser_dataReady", (pdfData) => {
            const texts = pdfData.Pages[0].Texts

            for (let i = 0; i < texts.length; i++) {
                const content = decodeURIComponent(texts[i].R[0].T)
                if (content == "Ijazah: " || content == "Nasional:") {
                    // console.log(`|${content}|(${i}) => ${texts[i + 1].R[0].T}`);
                    ijazahRow.push(texts[i + 1].R[0].T)
                    // appendToExcel('sss')
                }
            }
            appendToExcel(ijazahRow)
            // console.log(ijazahRow);
            ijazahData.push(ijazahRow)
        })

        pdfParser.loadPDF(filePath)
    })
    res.send("Berhasil upload " + req.files.length + " file!");
})


// pdfParser.on("pdfParser_dataError", (errData) =>
//     console.error(errData.parserError)

// )

// pdfParser.on("pdfParser_dataReady", (pdfData) => {
//     const texts = pdfData.Pages[0].Texts

//     for (let i = 0; i < texts.length; i++) {
//         const content = decodeURIComponent(texts[i].R[0].T)
//         if (content == "Ijazah: " || content == "Nasional:") {
//             console.log(`|${content}|(${i}) => ${texts[i + 1].R[0].T}`);
//         }
//     }
//     // texts.forEach(text => {
//     //     const x = text.x
//     //     const y = text.y
//     //     const content = decodeURIComponent(text.R[0].T)
//     //     if (content == "Ijazah: " || content == "Nasional:") {
//     //         const idx = texts.findIndex(item => item.R[0].T === content)
//     //         console.log(`|${idx}|`);
//     //     }
//     //     // console.log(typeof texts);
//     // });
//     // fs.writeFile(
//     //     "./output/coba.json",
//     //     JSON.stringify(pdfData),
//     //     (data) => console.log(data)
//     // )
// })

// pdfParser.loadPDF("./file/tes2.pdf")
// const data = 'Hello, world!';

// fs.writeFile("./output/coba.json", data, 'utf8', (err) => {
//     if (err) {
//         console.error('Error writing file:', err);
//         return;
//     }
//     console.log('File written successfully!');
// });