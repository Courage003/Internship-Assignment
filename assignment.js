const fs = require('fs')  //importing file system path
const excel = require('xlsx')   //importing excel library

const readJSONfile=(filePath)=>{
    try{
        const data=fs.readFileSync(filePath,'utf8');    //synchronously read contents of file with encoding utf8
        return JSON.parse(data);   //parsed (JSON string into Javascript object) object is returned
    }
    catch(error){
        console.error('Error while reading JSON file:', error.message);
        process.exit(1);
    }
}

//purpose of the function is to process the nested data within a json object and append it to excel workbook 
const NestedData = (jsonData, workBook) => {
    Object.keys(jsonData).forEach((key) => {
        const sheetData = jsonData[key];

        if (Array.isArray(sheetData)) {
            // Handling arrays (sections)
            sheetData.forEach((section) => {
                const sectionName = section.sectionName;
                const books = section.books || [];

                if (books.length > 0) {
                    const worksheet = excel.utils.json_to_sheet(books);   //creation of worksheet
                    excel.utils.book_append_sheet(workBook, worksheet, sectionName);   //appends it into excel workbook
                }
            });
        } else if (typeof sheetData === 'object' && sheetData !== null && sheetData !== undefined) {
            // Handling nested objects (exclude null and undefined)
            const worksheet = excel.utils.json_to_sheet([sheetData]);
            excel.utils.book_append_sheet(workBook, worksheet, key);
        } else {
            // Handling other data types
            const worksheet = excel.utils.json_to_sheet([{ [key]: sheetData }]);
            excel.utils.book_append_sheet(workBook, worksheet, key);
        }
    });
};

const convertToExcel= (jsonFilePath, outputfilePath)=>{
    const jsonData = readJSONfile(jsonFilePath);
    const workBook=excel.utils.book_new();

    NestedData(jsonData,workBook);


    try{
        excel.writeFile(workBook,outputfilePath);
        console.log('Execution is suucessful:', outputfilePath);
    }

    catch(error){
        console.error("error occurred", error.message);
        process.exit(1);
    }
}


const jsonFilePath='data.json';
const outputfilePath='SampleOutput 1.xlsx';
convertToExcel(jsonFilePath,outputfilePath);
