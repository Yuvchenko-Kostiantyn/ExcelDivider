const xlsx = require('xlsx');
const fs = require('fs');

const file = xlsx.readFile(`input/${process.argv[2]}`);
const sheets = file.Sheets
const sheet = sheets[Object.keys(sheets)[0]];
const data = xlsx.utils.sheet_to_json(sheet);
let fileIndex = 1;
const fileName = process.argv[2].split('.')[0]

function saveIntoFile(arr){
    // console.log(arr.length)
    let newWB = xlsx.utils.book_new();
    let newWS = xlsx.utils.json_to_sheet(arr);
    xlsx.utils.book_append_sheet(newWB, newWS, 'Worksheet');
    xlsx.writeFile(newWB, `output/${fileName}-${fileIndex}.xls`);
    fileIndex++
}

const sortedByGroup = data.sort((a,b) => {
    const fa = a['Группа'].toLowerCase();
    const fb = b['Группа'].toLowerCase();

    if(fa < fb){
        return -1;
    } else if(fa > fb){
        return 1;
    } else {
        return 0;
    }
})

let container = []
let count = 1;
for(let i = 0; i<=sortedByGroup.length-1; i++){
    container.push(sortedByGroup[i]);
    if(!sortedByGroup[i-1]){
        continue
    } else if(sortedByGroup[i]['Группа'].toLowerCase() !== sortedByGroup[i-1]['Группа'].toLowerCase()){
        count++;
    }

    if(count === 345){
        saveIntoFile(container);
        container = [];
        count = 0
    };;
}


if(container.length>0){
    saveIntoFile(container);
}