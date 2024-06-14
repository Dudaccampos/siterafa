function exportTableToXLSX(tableID, filename = '') {
    var table = document.getElementById(tableID);
    var workbook = XLSX.utils.table_to_book(table, {sheet: "Sheet1"});
    XLSX.writeFile(workbook, filename ? filename + '.xlsx' : 'excel_data.xlsx');
}
   
    filename = filename ? filename + '.xlsx' : 'excel_data.xlsx';

    downloadLink = document.createElement("a");

    document.body.appendChild(downloadLink);

    if (navigator.msSaveOrOpenBlob) {
        var blob = new Blob(['\ufeff', tableHTML], {
            type: dataType
        });
        navigator.msSaveOrOpenBlob(blob, filename);
    } else {
        
        downloadLink.href = 'data:' + dataType + ', ' + tableHTML;

       
        downloadLink.download = filename;

        downloadLink.click();
    }

    function exportTableToXLSX(tableID, filename = '') {
        var table = document.getElementById(tableID);
        var wb = XLSX.utils.table_to_book(table, {sheet: "Sheet JS"});
        
      
        var sheet = wb.Sheets["Sheet JS"];
        var range = XLSX.utils.decode_range(sheet['!ref']);
        
        for (let R = range.s.r + 1; R <= range.e.r; ++R) {
            let cellAddress = XLSX.utils.encode_cell({r: R, c: 2}); 
            let presencaCell = table.rows[R].cells[2];
            let presenteCheckbox = presencaCell.querySelector('input[name^="presenca"]:checked');
            let ausenteCheckbox = presencaCell.querySelector('input[name^="ausencia"]:checked');
            
            let presencaStatus = "";
            if (presenteCheckbox) {
                presencaStatus = presenteCheckbox.value;
            } else if (ausenteCheckbox) {
                presencaStatus = ausenteCheckbox.value;
            }
            
            if (sheet[cellAddress]) {
                sheet[cellAddress].v = presencaStatus;
            } else {
                sheet[cellAddress] = { v: presencaStatus }; 
            }
        }
        
       
        XLSX.writeFile(wb, (filename || 'Presenca') + ".xlsx");
    }
            
    