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

    let n = hours_response.length

    let details_response = xlsx.utils.sheet_to_json(
        details_sheet.Sheets[details_list[0]],
        {dateNF:"mm/dd/yy"}
    )

    // details_response -> employee details
    // hours_response -> jira excel
    
    // get year and month for jira
    const year = hours_response[0]["Year"]
    const monthName = Object.keys(hours_response[0])[3].split('\n')[0]
    const month = "JanFebMarAprMayJunJulAugSepOctNovDec".indexOf(monthName) / 3 + 1
    const lastDay = new Date(year, month + 1, 0).getDate()
    
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

    // Sheet 1 data
    hours_response.forEach((record) => {
        let emp_name = record["Time entry: User"]   
        
        let emp_record = details_response.find(emp => emp["Emp Name"] == emp_name && emp["Employment status"] == "Active")

        if(!emp_record) {
            // console.log(`Employee ${emp_name} does not exist`)
            return
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


    // Sheet 2 data
    let added = []
    for(let i = 0; i < n; i++) {
        let record = hours_response[i]
        let name = record["Time entry: User"]
        let hours1 = 0
        let hours2 = 0
        let emp_record = details_response.find(emp => emp["Emp Name"] == name && emp["Employment status"] == "Active")

        if(emp_record) {
            if(!added.find(emp => emp["Emp Name"] == name)) {
                for(let j = i; j < n; j++) {
                    if(hours_response[j]["Time entry: User"] === name) {
                        // console.log(Object.values(hours_response[j]).slice(3, 18), Object.values(hours_response[j]).slice(18, lastDay + 3))
                        hours1 += hoursSum(Object.values(hours_response[j]).slice(3, 18))
                        hours2 += hoursSum(Object.values(hours_response[j]).slice(18, lastDay + 3))
                    }
                }

                if(parseInt(emp_record["Pay cycle"]) == 1) {
                    added.push({
                        "Emp Name": name,
                        "Payment Amount": (hours1 + hours2) * emp_record["Pay rate"],
                        "Pay period": "30 days",
                        "Bill hours":  hours1 + hours2,
                        "Hours validated": 'N',
                        "Pay due": new Date(year, month - 1, 30),
                        "Payment done": 'N',
                        "Payment date": ""
                    })
                }
                else if(parseInt(emp_record["Pay cycle"]) == 2) {
                    added.push({
                        "Emp Name": name,
                        "Payment Amount": hours1 * emp_record["Pay rate"],
                        "Pay period": "First half",
                        "Bill hours":  hours1,
                        "Hours validated": 'N',
                        "Pay due": new Date(year, month - 1, 15),
                        "Payment done": 'N',
                        "Payment date": ""
                    })
    
                    added.push({
                        "Emp Name": name,
                        "Payment Amount": hours2 * emp_record["Pay rate"],
                        "Pay period": "Second half",
                        "Bill hours": hours2,
                        "Hours validated": 'N',
                        "Pay due": new Date(year, month - 1, 30),
                        "Payment done": 'N',
                        "Payment date": ""
                    })
                }
            }
        }        
        
    }

    let newWS2 = xlsx.utils.json_to_sheet(added, {dateNF:"mm/dd/yy"})
    xlsx.utils.book_append_sheet(newWB, newWS2, "Page_2")
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

const hoursSum = (obj) => {
    let hours = 0
    for(let i = 0; i < obj.length; i++) {
        hours += parseFloat(obj[i])
    }
    return hours
}

module.exports = {
    upload_excel
}