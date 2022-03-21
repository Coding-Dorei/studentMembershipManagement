const xlsx = require('xlsx')
/**
 * 
 * @param {String} fileName 
 * @returns jsonData
 */
function readExcel(fileName){
    excelFile = xlsx.readFile(fileName)
    sheetName = excelFile.SheetNames[0]
    firstSheet = excelFile.Sheets[sheetName]
    return xlsx.utils.sheet_to_json(firstSheet,{defval:""})
}

const 납부자명단 = readExcel("Membership.xlsx")
const 통장거래내역 = readExcel("BankingHistory.xlsx")
const dup = require('./sameName')

for(i = 0;i < 통장거래내역.length;i++){
    if(통장거래내역[i].맡기신금액 == 160000){
        let name = 통장거래내역[i].기재내용
        for(j = 납부자명단.length-1;j>=0;j--){
            if(납부자명단[j].name == name){
                if(dup.indexOf(name) != -1){
                    console.log(`Check:${name}`)
                }else{
                    //console.log(name)
                    납부자명단[j].level = 1
                }
                break
            }
        }
    }
}

const book = xlsx.utils.book_new()

const students = xlsx.utils.json_to_sheet(납부자명단)

students['!cols'] = [
    {wpx:120},
    {wpx:60},
    {wpx:40},
    {wpx:120}
]

students.E1 = ""
students.F1 = ""
students.G1 = ""
students.H1 = ""
students.I1 = ""
students.J1 = ""

xlsx.utils.book_append_sheet(book,students,"시트1")

xlsx.writeFile(book,"Membership.xlsx")