const xlsx = require('xlsx')

function readExcel(fileName){
    excelFile = xlsx.readFile(fileName)
    sheetName = excelFile.SheetNames[0]
    firstSheet = excelFile.Sheets[sheetName]
    return xlsx.utils.sheet_to_json(firstSheet,{defval:""})
}

const studentList = readExcel('과잠입금자목록.xlsx')
const bankingHistory = readExcel('BankingHistory.xlsx')
const res = readExcel('response.xlsx')
const students = readExcel('MemberShip.xlsx')

const book = xlsx.utils.book_new()

for(const raw of res){
    raw.납부금액 = 0
}

for(const raw of res){
    for(const student of studentList){
        if(raw.name == student.name) {
            raw.납부금액 = student.납부금액
            break
        }
    }
}

for(const raw of res){
    if(raw.납부금액 == 0){
        for(const raw2 of bankingHistory){
            if(raw2.기재내용 == raw.name){
                if(raw2.맡기신금액 >= 24000 && raw2.맡기신금액 < 160000){
                    raw.납부금액 = raw2.맡기신금액
                    if(raw.name == "박진희") raw.납부금액 = 34000
                    console.log(`${raw.name}님이 ${raw.납부금액}원 납부!`)
                    
                }
            }
        }
    }
}

// for(const res_raw of res){
//     let flag = 0
//     for(const student of students){
//         if(student.ID == res_raw.FULL_ID){
//             //console.log(`${student.ID} ${student.name}`)
//             flag = 1
//             break
//         }
//     }
//     if(!flag){
//         console.log(`No Such Student!!\n${res_raw.FULL_ID} ${res_raw.name}`)
//     }
// }

xlsx.utils.book_append_sheet(book,xlsx.utils.json_to_sheet(res),"설문지 응답 시트1")

xlsx.writeFile(book,"response.xlsx")