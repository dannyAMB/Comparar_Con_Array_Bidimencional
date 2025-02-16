
alumno=[];


document.getElementById("fileInput").addEventListener("change", function (event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const excelData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        compararDatos(excelData);
    };
    reader.readAsArrayBuffer(file);
});

function to_Lower_Case(letter){
//pasa letras a minusculas

if(letter!=null){


letter = letter.toLowerCase();


}
}

export async  function listMajors() {

alumno=[];

let response_academica;



/** let response;

try {
// Fetch first 10 files
response = await gapi.client.sheets.spreadsheets.values.get({
spreadsheetId: '1Sw6_zYTqpoy31ZXBja_rog1Xets6Lqc4ZXRFuVMG6Ec',
range: 'Turnos!A:G',
});
} catch (err) {
console.error('Error de api: '+ err)   
return;
}**/




try {

response_academica = await gapi.client.sheets.spreadsheets.values.get({
spreadsheetId: '1sgTqFiPnWf91D1YtgY3yud-El9GjADCFnWMpsfcyeD0',
range: 'MATRICULADOS!A:R',

});






} catch (err) {


document.getElementById("content_table_body").innerHTML =   '<td colspan="8" style="text-align:center;"> Academica - Error de Conexión </td>'

console.error('Error de api: ' + err)

return;
}



const range_academica = response_academica.result;
if (!range_academica || !range_academica.values || range_academica.values.length == 0) {
console.warn('se encontraron errores')

return;

}




// Flatten to string to display
console.log("validar variable status: " + status_conex)
console.log(range_academica)
range_academica.values.forEach((fila) => {




const dato = {
fecha: fila[1],
programa: fila[2],
nombre_full: fila[5],
primer_nombre: fila[6],
primer_apellido: fila[8],
id: fila[10],
usuario: fila[11],
clave: fila[12],
Correo_personal: to_Lower_Case(fila[15]),
tanda: fila[17]
};
alumno.push(dato);

});

}

function compararDatos(excelData) {


    const columnIndex = 10; // Comparar la columna "Nombre"
    let coincidencias = [];
    let diferencias = [];

    alumno.slice(1).forEach((fila, i) => {
        const valorArray = fila[columnIndex];
        const valorExcel = excelData[i + 1]?.[columnIndex] || "";

        if (valorArray === valorExcel) {
            coincidencias.push(valorArray);
        } else {
            diferencias.push(`${valorArray} (Excel: ${valorExcel || 'N/A'})`);
        }
    });

    mostrarResultados(coincidencias, diferencias);
}

function mostrarResultados(coincidencias, diferencias) {
    const coincidenciasList = document.getElementById("coincidenciasList");
    const diferenciasList = document.getElementById("diferenciasList");
    coincidenciasList.innerHTML = "";
    diferenciasList.innerHTML = "";

    coincidencias.forEach(valor => {
        const li = document.createElement("li");
        li.textContent = valor;
        coincidenciasList.appendChild(li);
    });

    diferencias.forEach(valor => {
        const li = document.createElement("li");
        li.textContent = valor;
        diferenciasList.appendChild(li);
    });
}
