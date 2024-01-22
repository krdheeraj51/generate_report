const ExcelJS = require('exceljs');
const _ = require("lodash");
const moment = require("moment");
const sampleData = require("./assets/sample.json"); 

async function taskExecution(){
    // 1. read json file
    // console.log(sampleData); 
    await createWorkBookwithSampleData(sampleData)

    // 2. Create a Excel file and write data into it
    // 3. pickup minimum and maximum date and put inside n number of columns 
    // 4. Print those date on that n number of colums
    // 5. Highlight columns and type a text in that column
    // 6. make that text Centre and color background



}

async function createWorkBookwithSampleData(data){
    // 2. Create a Excel file and write data into it
    const workbook = new ExcelJS.Workbook();
    const workbookSheet = workbook.addWorksheet('View',{
        views: [{ state: "frozen", ySplit: 1, xSplit:9 }]
    });
    // get weekdays 
    const weekDaysArray= await getWeekDaysArray(data);
    // create Column names
    let columnNames= await getColumnNames(data[0],weekDaysArray);
    workbookSheet.columns = columnNames;
    let rowsData= await addRowsToSheet(data);
    await workbookSheet.addRows(rowsData);
    await completeReportData(workbookSheet,data,weekDaysArray)
    // save under export.xlsx
    await workbook.xlsx.writeFile('./outPut/generateReport.xlsx');
}

async function completeReportData(workbookSheet,data,wekDaysArray){
    let columnCount =_.values(data[0]).length + 1;
    const format='DD-MMM';
    const plStArray=[];
    const plEndArray=[];
    console.log("wekDaysArray ::",wekDaysArray);
    for(let rowIndex=0;rowIndex<data.length;rowIndex++){
        let plStartD= moment(data[rowIndex]['Plan Start'],format);
        let plEndD= moment(data[rowIndex]['Plan End'],format);
        let kaStartD= moment(data[rowIndex]['KA Start'],format);
        let kaEndD= moment(data[rowIndex]['KA End'],format);
        let shadowStartD= moment(data[rowIndex]['Shadow Start'],format);
        let shadowEndD= moment(data[rowIndex]['Shadow End'],format);
        let revStartD= moment(data[rowIndex]['Rev Start'],format);
        let revEndD= moment(data[rowIndex]['Rev End'],format);
        let goLiveD= moment(data[rowIndex]['Go Live'],format);
        
        const plIndexes= await findIndexRange(plStartD,plEndD,wekDaysArray,columnCount,rowIndex);
        const kaIndexes= await findIndexRange(kaStartD,kaEndD,wekDaysArray,columnCount,rowIndex);
        const shadowIndexes= await findIndexRange(shadowStartD,shadowEndD,wekDaysArray,columnCount,rowIndex);
        const revIndexes= await findIndexRange(revStartD,revEndD,wekDaysArray,columnCount,rowIndex);
        workbookSheet.mergeCells(plIndexes);
        const plMergeCells= workbookSheet.getCell(plIndexes.split(':')[0]);
        plMergeCells.value='Planning';
        plMergeCells.alignment={horizontal:'center',vertical:'middle'};
        plMergeCells.fill={type:'pattern',pattern:'solid',fgColor:{argb:'ff90eebf'}};
        workbookSheet.mergeCells(kaIndexes);
        const kaMergeCells= workbookSheet.getCell(kaIndexes.split(':')[0]);
        kaMergeCells.value='Knowledge Acquisition';
        kaMergeCells.alignment={horizontal:'center',vertical:'middle'};
        kaMergeCells.fill={type:'pattern',pattern:'solid',fgColor:{argb:'ff00ffff'}};
        workbookSheet.mergeCells(shadowIndexes);
        const shadowMergeCells= workbookSheet.getCell(shadowIndexes.split(':')[0]);
        shadowMergeCells.value='Shadow';
        shadowMergeCells.alignment={horizontal:'center',vertical:'middle'};
        shadowMergeCells.fill={type:'pattern',pattern:'solid',fgColor:{argb:'ff0d98ba'}};
        workbookSheet.mergeCells(revIndexes);
        const revMergeCells= workbookSheet.getCell(revIndexes.split(':')[0]);
        revMergeCells.value='Rev Shadow';
        revMergeCells.alignment={horizontal:'center',vertical:'middle'};
        revMergeCells.fill={type:'pattern',pattern:'solid',fgColor:{argb:'ff3eb489'}};
        const findDateIndex = (date) =>_.findIndex(wekDaysArray,(d)=>moment(d,'DD-MMM').isSame(moment(date,'DD-MMM')));
        const goLiveindex= findDateIndex(goLiveD);
        const goLiveColindex=goLiveindex + columnCount;
        const goLiveIndexes= await convertNumberExcelColumn(goLiveColindex);
        workbookSheet.mergeCells(`${goLiveIndexes}${rowIndex + 2}`);
        const goLiveMergeCells= workbookSheet.getCell(`${goLiveIndexes}${rowIndex + 2}`);
        goLiveMergeCells.value='Live';
        goLiveMergeCells.alignment={horizontal:'center',vertical:'middle'};
        goLiveMergeCells.fill={type:'pattern',pattern:'solid',fgColor:{argb:'ff3e69b4'}};
    }
}
async function findIndexRange(startDate,endDate,weekDaysList,colCount,index){
    const startDateIndex = _.findIndex(weekDaysList,(date)=>moment(date,'DD-MMM').isSame(moment(startDate,'DD-MMM')));;
    const endDateIndex = _.findLastIndex(weekDaysList,(d)=>moment(d,'DD-MMM').isSameOrBefore(moment(endDate,'DD-MMM')));
    console.log("startDateIndex ::",startDateIndex);
    console.log("endDateIndex ::",endDateIndex);
    if(startDateIndex !==-1 && endDateIndex!== -1){
        let startCol= await convertNumberExcelColumn(startDateIndex + colCount);
        let endCol= await convertNumberExcelColumn(endDateIndex + colCount);
        console.log(`${startCol}${index+2}:${endCol}${index+2}`);
        if(startCol === endCol){
            return  `${startCol}${index+2}`
        }
        return  `${startCol}${index+2}:${endCol}${index+2}`;
    
    }
    return;
    
}


async function getWeekDaysArray(data){
    /**
    * Calculate weeks days 
    */
    const PlaningDateArray=[];
    const shadowDateArray=[];
    const kADateArray=[];
    const revDateArray=[];
    const goLiveDateArray=[];
    for(let da of data){
        const format='DD-MMM-YY';
        let planStartDate= moment(da["Plan Start"],format);
        let planEndDate = moment(da["Plan End"],format);
        let planCurrentDate = planStartDate.clone();
        while (planCurrentDate.isSameOrBefore(planEndDate)){
            if(planCurrentDate.day()===1){
                PlaningDateArray.push(planCurrentDate.clone().format('DD-MMM-YY'));  
            }
            planCurrentDate.add(1,'day');
        }

        let shadowStartDate= moment(da["Shadow Start"],format);
        let shadowEndDate = moment(da["Shadow End"],format);
        let shadowCurrentDate = shadowStartDate.clone();
        while (shadowCurrentDate.isSameOrBefore(shadowEndDate)){
            if(shadowCurrentDate.day()===1){
                shadowDateArray.push(shadowCurrentDate.clone().format('DD-MMM-YY'));  
            }
            shadowCurrentDate.add(1,'day');
        }
        let kAStartDate= moment(da["KA Start"],format);
        let kAEndDate = moment(da["KA End"],format);
        let kACurrentDate = kAStartDate.clone();
        while (kACurrentDate.isSameOrBefore(kAEndDate)){
            if(kACurrentDate.day()===1){
                kADateArray.push(kACurrentDate.clone().format('DD-MMM-YY'));  
            }
            kACurrentDate.add(1,'day');
        }
        let revStartDate= moment(da["Rev Start"],format);
        let revEndDate = moment(da["Rev End"],format);
        let revCurrentDate = revStartDate.clone();
        while (revCurrentDate.isSameOrBefore(revEndDate)){
            if(revCurrentDate.day()===1){
                revDateArray.push(revCurrentDate.clone().format('DD-MMM-YY'));  
            }
            revCurrentDate.add(1,'day');
        }
       let goLiveDate= moment(da["Go Live"],format);    
       goLiveDateArray.push(goLiveDate.clone().startOf('isoWeek').format('DD-MMM-YY'));


    }
    let combinedArray= [...PlaningDateArray,...shadowDateArray,...kADateArray,...revDateArray,...goLiveDateArray];
    let combinedSortUniqueList = _.chain(combinedArray).uniq().sortBy((date)=>moment(date,'DD-MMM-YY')).value();
    return combinedSortUniqueList;
}


async function convertNumberExcelColumn(number){
    console.log("number ::",number);
    let result='';
    while(number >0){
        const remainder= (number -1) % 26;
        result = String.fromCharCode(65 + remainder) + result;
        number = Math.floor((number -1)/26);

    }
    return result;
}

async function addRowsToSheet(data){
    let newArray=[];
    for(let i=0;i<data.length;i++){
       let rowsArray= _.values(data[i]);      
       newArray.push(rowsArray);
    }
    return newArray;
}

async function getColumnNames(data,weekDays){
    let columnArray= _.keys(data);
    let modifyWeekDaysFormat= _.map(weekDays,(date)=>moment(date,'DD-MMM-YY').format('DD-MMM'));
    let combinedArray= columnArray.concat(...modifyWeekDaysFormat);
    const columnObjArray= _.map(combinedArray,(item,index)=>({
        header: item,
        id: index + 1
    }));
    return columnObjArray;

}

taskExecution();