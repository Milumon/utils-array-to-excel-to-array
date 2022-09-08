var XLSX = require("xlsx");

const path = "Libro1.xlsx"

const getArrayFromExcelFile = (path) => {

    var workbook = XLSX.readFile(path);
    let data = []

    for (const sheet in workbook.Sheets) {
        if (workbook.Sheets[workbook.SheetNames[0]]) {
            data = data.concat(XLSX.utils.sheet_to_json(workbook.Sheets[sheet], { defval: '' }))
        }
    }

    var whitelistArray = data.map(x => {
        if (x != '' && x != undefined && x.Whitelist.includes("@improbable.io")) {
            return { whitelist: x["Whitelist"].toLowerCase() }
        }
    }).filter(x => x != undefined)

    var currentArray = data.map(x => {
        if (x != '' && x != undefined && x.Current.includes("@improbable.io")) {
            return { current: x["Current"].toLowerCase() }
        }
    }).filter(x => x != undefined)

    var deleteEmailArray = []
    var emailsIncludedArray = []

    currentArray.forEach(element => {

        const hasEmail = whitelistArray.some(function (email) {
            return email.whitelist === element.current;
        })

        if (!hasEmail) {
            deleteEmailArray.push({ deleteEmail: element.current })
        }
        else {
            emailsIncludedArray.push({ emailsIncluded: element.current })
        }
    })


    console.log("MIEMBROS ACTUALES EN SUNLIGHT (GO1) CON CORREO @improbable.io: " + currentArray.length)
    console.log("CANTIDAD DE USUARIOS EN LA LISTA CON CORREO @improbable.io: " + whitelistArray.length)
    console.log("MIEMBROS ACTUALES EN SUNLIGHT (GO1) CON CORREO @improbable.io EN LA LISTA: " + emailsIncludedArray.length)
    console.log("MIEMBROS ACTUALES EN SUNLIGHT (GO1) CON CORREO @improbable.io FUERA DE LA lista: " + deleteEmailArray.length)

    return { whitelistArray, currentArray, deleteEmailArray, emailsIncludedArray }
}


const { currentArray, whitelistArray, deleteEmailArray, emailsIncludedArray } = getArrayFromExcelFile(path)


const exportToExcel = (filename, array) => {
    const fileName = filename + ".xlsx";
    const worksheet = XLSX.utils.json_to_sheet(array);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'test');

    XLSX.writeFile(workbook, fileName);
}


exportToExcel("whitelist", whitelistArray)
exportToExcel("current", currentArray)
exportToExcel("delete", deleteEmailArray)
exportToExcel("emailsIncludedArray", emailsIncludedArray)
