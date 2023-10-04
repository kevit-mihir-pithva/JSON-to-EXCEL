const XLSX = require("xlsx")
const fs = require("fs")

//sample data
customers = [
  {
    "name": { "first": "mihir", "last": "pithva" },
    "email": "mihir@example.com",
    "dateOfBirth": "2003-03-17",
    "customerId": 1,
  },
  {
    "name": { "first": "nik", "last": "bhadeshiya" },
    "email": "nik@example.com",
    "dateOfBirth": "2003-09-12",
    "customerId": 2,
  },
  {
    "name": { "first": "amay", "last": "lunagaria" },
    "email": "amay@example.com",
    "dateOfBirth": "2003-11-18",
    "customerId": 2,
  },
]

//extracting useful data fields and calculate new field age
const excelData = customers.map((customer) => {
    const { first , last } = customer.name
    const email = customer.email
    const dob = new Date(customer.dateOfBirth)
    const age = Math.floor((new Date() - dob) / (365.25 * 24 * 60 * 60 * 1000))
    return {
        "First Name":first,
        "Last Name":last || "",
        "Email":email,
        "Age":age
    }
})

//creating new excel work-book and data(exceldata) 
const work_book = XLSX.utils.book_new()

//converts an array of JS objects(excelData) to a worksheet.
const work_sheet = XLSX.utils.json_to_sheet(excelData)

//addind a worksheet to a workbook
XLSX.utils.book_append_sheet(work_book,work_sheet,"CUSTOMERS")

//saving the excel file
const data = XLSX.write(work_book,{ bookType: "xlsx" , type:"buffer"})
fs.writeFileSync("customer_data.xlsx",data,"binary")

console.log("Excel file created...")