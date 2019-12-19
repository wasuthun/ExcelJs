const ExcelJS = require('exceljs');
var workbook = new ExcelJS.Workbook();
var arrayId = []
var arrayItem = []
var file = workbook.xlsx.readFile('AAAAS.xlsx')
  .then(function(e) {
    const sheet = e.getWorksheet()
    sheet.eachRow(row => {
        arrayId.push(row.getCell(5).value)
        arrayItem.push({id:row.getCell(5).value,item:row.getCell(6).value})
    })
    var newSheet = e.addWorksheet('newsSheet')
    var count = 4
    arrayId=arrayId.filter(function(item, index){
                return arrayId.indexOf(item) >= index;
            });
    arrayItem.forEach( (item,i) => {
            newSheet.getRow(1).getCell(4).value=arrayItem[1].id
            newSheet.getRow(1).getCell(5).value=arrayItem[1].item
        for (let index = 2; index <= arrayId.length; index++) {
            if(arrayId[index]===item.id){
                if(arrayItem[i].id===arrayItem[i-1].id){
                count=count+1
                newSheet.getRow(index).getCell(4).value=item.id
                newSheet.getRow(index).getCell(count).value=item.item
                }
                else{
                    count=5
                    newSheet.getRow(index).getCell(4).value=item.id
                    newSheet.getRow(index).getCell(count).value=item.item
                }
            }
        }
    })
    e.xlsx.writeFile('as.xlsx')
  });