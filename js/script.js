// SELECTORS
const result = document.querySelector('.result-container');

// EVENT HANDLER
document.getElementById('upload').addEventListener('change', handleFileSelect, false);

// FUNCTIONS
function handleFileSelect(evt) {
  var files = evt.target.files; // FileList object
  ReadExcel(files);
}

async function ReadExcel(fileList) {
  const blob = new Blob([fileList[0]], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=utf-8' });
  const buffer = await blob.arrayBuffer();
  
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  
  const sheet = workbook.worksheets[0];

  const lastRow = sheet.actualRowCount;
  const lastCol = sheet.actualColumnCount;
  console.log(lastRow);
  console.log(lastCol);
  
  const table = document.createElement('table');
  result.appendChild(table);
  const thead = document.createElement('thead');
  table.appendChild(thead);
  const tbody = document.createElement('tbody');
  table.appendChild(tbody);

  for (var i = 1; i <= lastRow; i++) {
    console.log(i);
    const tr = document.createElement('tr');
    let row = 0;

    sheet.getColumn(i).eachCell({ includeEmpty: true }, function (cell, rowNumber) {
      if (i===1) {
        var c = document.createElement('th');
        if (cell.value === 'BusinessDate') {
          c.innerHTML = 'Names / Days';
        }  else {
          var date = new Date(cell.value).getDate();
          c.innerHTML = date;
        }
        tr.appendChild(c);
      } else {
        if (rowNumber === 1) {
          var c = document.createElement('th');
        } else {
          var c = document.createElement('td');
        }
        c.innerHTML = cell.value;
        tr.appendChild(c);
      };
      row = rowNumber;
    });

    // Add to thead or tbody
    if (i === 1 && row === lastRow) {
      thead.appendChild(tr);
    } else {
      tbody.appendChild(tr);
    };
  };
};

function ExcelDateToJSDate(date) {
  return new Date(Math.round((date - 25569)*86400*1000));
}