const xlsx = require("xlsx")
const fs = require("fs")

const generateReport = async (req, res) => {
    const paid = req.body.paid
    const due_date = new Date(req.body.date)

    const month = due_date.getMonth() + 1
    const year = due_date.getFullYear()

    if(!fs.existsSync(`sheets/reports/${month}-${year}.xlsx`)) {
        res.send("Report for month does not exist")
        return
    }

    const WB = xlsx.readFile(`sheets/reports/${month}-${year}.xlsx`, {
        cellDates: true,
        cellNF: false,
        cellText: false
    })
    const sheet_list = WB.SheetNames

    let records = xlsx.utils.sheet_to_json(WB.Sheets[sheet_list[1]], { 
        dateNF: "yyyy-mm-dd",
        raw: false
     })

    let ret = []

    records.forEach(record => {
        record["Pay due"] = new Date(record["Pay due"] + 'T00:00:00Z')
        if((paid =="No" && record["Payment done"] == "N") && due_date.valueOf() == record["Pay due"].valueOf()) {
            console.log(record)
            ret.push(record)
        }
        else if((paid == "Yes" && record["Payment done"] == "Y") && due_date.valueOf() == record["Pay due"].valueOf()) {
            ret.push(record)
        }
    })

    res.json(ret)
}

const generateReport2 = async (req, res) => {
    const empName = req.query.emp_name
    let all = false

    if (empName == "undefined" || empName == "") {
        all = true
    }

    const startDate = new Date(req.query.start_date)
    const endDate = new Date(req.query.end_date)

    if (startDate > endDate) {
        res.json([])
        return
    }

    const year = startDate.getFullYear()
    const startMonth = startDate.getMonth() + 1
    const endMonth = endDate.getMonth() + 1
    let first = startDate
    let last

    let ret = []

    for (let i = startMonth; i <= endMonth; i++) {
        if (fs.existsSync(`sheets/reports/${i}-${year}.xlsx`)) {
            let WB = xlsx.readFile(`sheets/reports/${i}-${year}.xlsx`, { cellDates: true })
            let sheet_list = WB.SheetNames

            let records = xlsx.utils.sheet_to_json(WB.Sheets[sheet_list[0]], { dateNF: "mm/dd/yy" })

            let lastDayOfMonth = new Date(year, i, 0)
            last = endDate < lastDayOfMonth ? endDate : lastDayOfMonth

            if (all) {
                records.forEach(record => {
                    let temp1 = Object.keys(record).slice(0, 13).reduce((result, key) => {
                        result[key] = record[key];
                        return result;
                    }, {})

                    temp1["Employment start date"].setSeconds(temp1["Employment start date"].getSeconds() + 10)
                    temp1["Employment end date"].setSeconds(temp1["Employment end date"].getSeconds() + 10)

                    let temp2 = generateDateIndices(first, last).map(key => record[key]).reduce((a, b) => a + b, 0)

                    ret.push({
                        ...temp1,
                        "Hours worked": temp2,
                        "Payment Amount": temp1["Pay rate"] * temp2
                    })
                })
            }
            else {
                records.forEach(record => {
                    if (record["Emp Name"] == empName) {
                        let temp1 = Object.keys(record).slice(0, 13).reduce((result, key) => {
                            result[key] = record[key];
                            return result;
                        }, {})

                        temp1["Employment start date"].setSeconds(temp1["Employment start date"].getSeconds() + 10)
                        temp1["Employment end date"].setSeconds(temp1["Employment end date"].getSeconds() + 10)

                        let temp2 = generateDateIndices(first, last).map(key => record[key]).reduce((a, b) => a + b, 0)

                        ret.push({
                            ...temp1,
                            "Hours worked": temp2,
                            "Payment Amount": temp1["Pay rate"] * temp2
                        })
                    }
                })
            }


        }
        first.setDate(last.getDate() + 1)
    }


    // console.log(ret)

    res.json(ret)
}

const cutObject = (obj, max) => Object.keys(obj)
    .filter((key, index) => index < max)
    .map(key => ({ [key]: obj[key] }));

const generateDateIndices = (start_, end_) => {
    let start = new Date(start_)
    let end = new Date(end_)

    let ret = []

    for (let i = start; i <= end; i.setDate(i.getDate() + 1)) {
        let monthName = i.toLocaleString('default', { month: 'short' })

        ret.push(`${monthName}\n${i.getDate().toLocaleString('en-US', { minimumIntegerDigits: 2, useGrouping: false })}`)
    }
    return ret
}

module.exports = {
    generateReport,
    generateReport2
}