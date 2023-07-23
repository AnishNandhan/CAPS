const xlsx = require("xlsx")

const generateReport = async (req, res) => {
    const empName = req.query.emp_name
    let all = false

    if (empName == "undefined" || empName == "") {        
        all = true
    }

    const startDate = new Date(req.query.start_date)
    const endDate = new Date(req.query.end_date)

    if(startDate > endDate) {
        res.json([])
        return
    }

    const year = startDate.getFullYear()
    const startMonth = startDate.getMonth() + 1
    const endMonth = endDate.getMonth() + 1
    const startDay = startDate.getDate()
    const endDay = endDate.getDate()

    let ret = []

    for(let i = startMonth; i <= startMonth; i++) {
        let WB = xlsx.readFile(`sheets/reports/${i}-${year}.xlsx`, {cellDates: true})
        let sheet_list = WB.SheetNames

        let records = xlsx.utils.sheet_to_json(WB.Sheets[sheet_list[0]], {dateNF:"mm/dd/yy"})

        if(all) {
            records.forEach(record => {
                let temp1 = Object.keys(record).slice(0, 13).reduce((result, key) => {
                    result[key] = record[key];    
                    return result;
                }, {})

                temp1["Employment start date"].setSeconds(temp1["Employment start date"].getSeconds() + 10)
                temp1["Employment end date"].setSeconds(temp1["Employment end date"].getSeconds() + 10)
    
                let temp2 = generateDateIndices(startDate, endDate).map(key => record[key]).reduce((a, b) => a + b, 0)
    
                ret.push({
                    ...temp1,
                    "Hours worked" : temp2
                })         
            })
        } 
        else {
            records.forEach(record => {
                if(record["Emp Name"] == empName) {
                    let temp1 = Object.keys(record).slice(0, 13).reduce((result, key) => {
                        result[key] = record[key];        
                        return result;
                    }, {})

                    temp1["Employment start date"].setSeconds(temp1["Employment start date"].getSeconds() + 10)
                    temp1["Employment end date"].setSeconds(temp1["Employment end date"].getSeconds() + 10)
        
                    let temp2 = generateDateIndices(startDate, endDate).map(key => record[key]).reduce((a, b) => a + b, 0)
        
                    ret.push({
                        ...temp1,
                        "Hours worked" : temp2
                    })
                }            
            })
        }
        
        
    }

    // console.log(ret)
    
    res.json(ret)
}

const cutObject = (obj, max) => Object.keys(obj)
  .filter((key, index) => index < max)
  .map(key => ({[key]: obj[key]}));

const generateDateIndices =  (start_, end_) => {
    let start = new Date(start_)
    let end = new Date(end_)
    let monthName = start.toLocaleString('default', { month: 'short' })
    let month = start.getMonth() + 1
    let year = start.getFullYear()
    let startDay = start.getDate()
    let endDay = end.getDate()

    let n = new Date(year, month, 0)
    let ret = []

    for(let i = startDay; i <= endDay; i++) {
        ret.push(`${monthName}\n${i.toLocaleString('en-US', {minimumIntegerDigits: 2, useGrouping: false})}`)
    }
    return ret
}

module.exports = {
    generateReport
}