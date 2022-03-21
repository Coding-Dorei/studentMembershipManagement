const xlsx = require('xlsx')

const book = xlsx.utils.book_new()

function readExcel(fileName){
    excelFile = xlsx.readFile(fileName)
    sheetName = excelFile.SheetNames[0]
    firstSheet = excelFile.Sheets[sheetName]
    return xlsx.utils.sheet_to_json(firstSheet,{defval:""})
}

const studentList = readExcel('Membership.xlsx')

const BankingHistory =  readExcel('BankingHistory.xlsx')

const dup = require("./sameName")

let sheetJson = []

let refund = ["김성지"]

let except = ["류태우","윤동국","이근희"]

for(const raw of BankingHistory){
    if(raw.맡기신금액 == 34000 || raw.맡기신금액 == 24000 || raw.맡기신금액 == 184000){
        let name = raw.기재내용
        if(refund.indexOf(name) != -1) continue
        if(dup.indexOf(name) != -1){
            console.log(`Check:${name}`)
            continue
        }
        for(const raw2 of studentList){
            if(raw2.name == raw.기재내용){
                if(except.indexOf(name) != -1){
                    raw2.level = "tmp"
                }
                sheetJson.push({
                    name:name,
                    납부금액:raw.맡기신금액 % 160000,
                    학생회비납부:raw2.level || 0,
                    학번:raw2.ID.toString().slice(0,2)
                })
                break
            }
        }
    }
}

console.log(sheetJson)

xlsx.utils.book_append_sheet(book,xlsx.utils.json_to_sheet(sheetJson),"시트1")

xlsx.writeFile(book,"과잠입금자목록.xlsx")