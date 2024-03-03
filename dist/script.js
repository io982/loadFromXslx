
const url = "/test.xlsx";

fetch(url)
    .then(resp => resp.arrayBuffer())
    .then(arrayBuffer => readXlsxFile(arrayBuffer))
    .catch(error => console.log('Ошибка:', error));

    function readXlsxFile(buffer) {
        const data = new Uint8Array(buffer);
        const workbook = XLSX.read(data, {type: 'array', cellStyles: true});
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];

        const jsa = XLSX.utils.sheet_to_json(worksheet);        
        const formulas = XLSX.utils.sheet_to_formulae(worksheet);
        const html = XLSX.utils.sheet_to_html(worksheet);

        jsa[4].age = 12;

        worksheet['B6'].v = 12;
        worksheet['B6'].w = '12';
        // Object.values(worksheet).forEach(x => {
        //     if (x.f) {
        //         x.f = x.f.replace('TEXTJOIN','ОБЪЕДИНИТЬ')
        //     }
        // })
        XLSX.writeFile(workbook, "Rslt.xlsx", {type: 'array', cellStyles: true});

        

        let root = document.querySelector('#root')
        root.innerHTML =  html;
        console.log(workbook);
        console.log(worksheet);
        console.log(jsa);
        console.log(formulas);
    }