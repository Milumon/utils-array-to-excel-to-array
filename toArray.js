

var XLSX = require("xlsx");

const members = []

const finalArray = []



members.forEach((x) => {
    finalArray.push({ "group_name": x[0], "user_mail": x[1], "user_name": x[2], "hasOrderPending": x[3] })
})

const exportToExcel = (filename, array) => {
    const fileName = filename + ".xlsx";
    const worksheet = XLSX.utils.json_to_sheet(array);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'test');

    XLSX.writeFile(workbook, fileName);
}

console.log(finalArray)

exportToExcel("report_orders_improbable", finalArray)