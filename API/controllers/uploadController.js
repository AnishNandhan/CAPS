const path = require("path")
const fs = require("fs")
const xlsx = require("xlsx")

const upload_excel = (req, res) => {
    if(path.extname(req.file.originalname) != ".xlsx") {
        res.send("Upload excel file!")
        return
    }

    const hours_sheet = xlsx.readFile(`sheets/uploads/${req.file.originalname}`)
    let hours_list = hours_sheet.SheetNames

    const details_sheet = xlsx.readFile("sheets/EmpDetails.xlsx", {cellDates: true})
    let details_list = details_sheet.SheetNames
    
    let hours_response = xlsx.utils.sheet_to_json(
        hours_sheet.Sheets[hours_list[0]]
    )

    let details_response = xlsx.utils.sheet_to_json(
        details_sheet.Sheets[details_list[0]],
        {dateNF:"mm/dd/yy"}
    )

    console.log(details_response)

    // details_response -> employee details
    // hours_response -> jira excel
    
    // get year and month for jira
    const year = hours_response[0]["Year"]
    const monthName = Object.keys(hours_response[0])[3].split('\n')[0]
    const month = "JanFebMarAprMayJunJulAugSepOctNovDec".indexOf(monthName) / 3 + 1
    
    // what to do if report for month already exits
    if(fs.existsSync(`sheets/reports/${month}-${year}${path.extname(req.file.originalname)}`)) {
        fs.unlink(`sheets/reports/${month}-${year}${path.extname(req.file.originalname)}`, (err) => {
            if (err) {
                console.log(err)
                return
            }
        })
        // res.send("Report for month already exists")
        // next()
    }
    
    let newData = []

    hours_response.forEach((record) => {
        let emp_name = record["Time entry: User"]   
        
        let emp_record = details_response.find(emp => emp["Emp Name"] == emp_name && emp["Employment status"] == "Active")

        if(!emp_record) {
            console.log(`Employee ${emp_name} does not exist`)
            return
            // res.send(`Employee ${emp_name} does not exist`)
            // next()
        }

        let newRecord = {
            ...emp_record, 
            ...omit(record, "Time entry: User", "Year")
        }

        newData.push(newRecord)
    })

    // write into new excel file
    let newWB = xlsx.utils.book_new()
    let newWS = xlsx.utils.json_to_sheet(newData, {dateNF:"mm/dd/yy"})
    xlsx.utils.book_append_sheet(newWB, newWS, "Page_1")
    xlsx.writeFile(newWB, `sheets/reports/${month}-${year}.xlsx`)

    // delete uploaded excel after generating report
    fs.unlink(`sheets/uploads/${req.file.originalname}`, (err) => {
        if (err) {
            console.log(err)
            return
        }
    })
    res.send("done")
}

const omit = (obj, ...props) => {
    const result = { ...obj };
    props.forEach(function(prop) {
      delete result[prop];
    });
    return result;
}

module.exports = {
    upload_excel
}