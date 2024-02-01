let table = document.getElementsByClassName("sheet-body")[0];
let rows = document.getElementsByClassName("rows")[0];
let columns = document.getElementsByClassName("columns")[0];
let tableExists = false;

const generateTable = () => {
    let rowsNumber = parseInt(rows.value);
    let columnsNumber = parseInt(columns.value);
    /* ========================================== Alert if fields have invlid inputs =============================================== */
   
    if (isNaN(rowsNumber) || isNaN(columnsNumber) || rowsNumber <= 0 || columnsNumber <= 0) {
        
        Swal.fire({
            icon: 'error',
            title: 'Oops...',
            text: 'Please enter valid numbers for rows and columns!',
        });
        return;
    }
    
    table.innerHTML = "";
    
    for(let i = 0; i < rowsNumber; i++) {
        let tableRow = "";
        
        for(let j = 0; j < columnsNumber; j++) {
            tableRow += `<td contenteditable></td>`;
        }
        
        table.innerHTML += tableRow;
    }
    
    tableExists = true;
};


const ExportToExcel = (type, fn, dl) => {
    if (!tableExists) {
            /* ========================================== Alert if there is no table to be generated =============================================== */

        Swal.fire({
            icon: 'error',
            title: 'Oops...',
            text: 'There is no table to export!',
        });
        return;
    }
    var elt = table;
    var wb = XLSX.utils.table_to_book(elt, { sheet: "sheet1" });
    return dl ? XLSX.write(wb, { bookType: type, bookSST: true, type: 'base64' })
        : XLSX.writeFile(wb, fn || ('MyNewSheet.' + (type || 'xlsx')));
};
