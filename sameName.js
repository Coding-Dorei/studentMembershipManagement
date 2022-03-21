const xlsx = require('xlsx')

const excelFile = xlsx.readFile('Membership.xlsx')

let sheetName = excelFile.SheetNames[0]
let mainSheet = excelFile.Sheets[sheetName]

const jsonData = xlsx.utils.sheet_to_json(mainSheet,{defval:""})

let uniqueNames = []
let duplicates = []
for(i  = 0; i < jsonData.length;i++){
    if(uniqueNames.indexOf(jsonData[i].name) != -1){
        //console.log(jsonData[i].name)
        if(duplicates.indexOf(jsonData[i].name) == -1){
            duplicates.push(jsonData[i].name)
        }
    }else {
        uniqueNames.push(jsonData[i].name)
    }
}

module.exports = duplicates