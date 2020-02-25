const XLSX = require('xlsx');
const utils = require('@stefancfuchs/utils');


(async () => {

    // Read Student records spreadSheet
    const studentRecords = XLSX.readFile('assets/HISTÓRICO E CONTATO ALUNOS ATIVOS.ods');
    const sheetNames = studentRecords.SheetNames;
    const sheets = studentRecords.Sheets;
    const students = [];

    console.log(sheetNames.length + ' sheets in students records spreadsheet');

    for (let sheet of sheetNames) {
        const acs = sheets[sheet];

        for (let i = 3; i < 50; i++) {

            const nameCell = acs['B' + i];
            const statusCell = acs['C' + i];

            if (nameCell && statusCell && statusCell.v === 'ATIVO') {

                const name = utils.accentFold(nameCell.v).toLowerCase().trim();
                const register = acs['D' + i] ? acs['D' + i].v : null;

                let birth = acs['F' + i] ? acs['F' + i].w : null;
                if (acs['F' + i] && !birth.includes('/')) {
                    birth = new Date(acs['F' + i].w);
                }

                const docCell = acs['H' + i];
                const doc = docCell ? docCell.v : null;

                students.push({ name, register, birth, doc, sheet });
            }
        }
    }

    console.log(students.length, ' student records found at records spreadsheet');

    const studenRecordOf = (name, students) => {
        const st = students.filter(s => s.name === name);
        if (st.length > 1) {
            console.log('Warning: Two students have the name of ' + name);
            return null;
        }
        //if(st && st.length) console.log('rec found', name)
        return (st && st.length) ? st[0] : null;
    }


    // Read student id cards spreadsheet
    const idCardsAll = XLSX.readFile('assets/Controle fotos carteiras estudantis _ carteirinhas de estudante  2020.ods');
    let idCardsSheet = idCardsAll.Sheets['Falta imprimir'];
    let studentsIdsToPrintCount = 0;
    let valuesFound = 0;

    for (let i = 4; i < 600; i++) {

        const nameCell = idCardsSheet['B' + i];

        if (nameCell && nameCell.v && nameCell.v.includes('-')) {

            const name = utils.accentFold(nameCell.v.toLowerCase().split('-')[0]).trim();
            const rec = studenRecordOf(name, students);
            studentsIdsToPrintCount++;

            if (rec) {

                const docCell = idCardsSheet['E' + i];
                const birthCell = idCardsSheet['F' + i];
                const registerCell = idCardsSheet['G' + i];

                if (rec.doc && (!docCell || !docCell.v || docCell.v.length < 5)) {
                    idCardsSheet['E' + i] = { v: rec.doc, t: 's', w: rec.doc }
                    valuesFound++;
                    //console.log(name, 'doc', rec.doc, rec.sheet)
                }
                if (rec.birth && (!birthCell || !birthCell.v)) {
                    idCardsSheet['F' + i] = { v: rec.birth, t: 's', w: rec.birth }
                    valuesFound++;
                    //console.log(name, 'birth', rec.birth, rec.sheet)
                }
                if (rec.register && (!registerCell || !registerCell.v || registerCell.v.length < 6)) {
                    idCardsSheet['G' + i] = { v: rec.register, t: 's', w: rec.register }
                    //console.log(name, 'reg', rec.register, rec.sheet)
                    valuesFound++;
                }

            }

        }

    }

    console.log(studentsIdsToPrintCount, ' student records to print, ', valuesFound, ' values found and completed');
    XLSX.writeFile(idCardsAll, 'assets/Controle fotos carteiras estudantis _ carteirinhas de estudante __ MOD.ods')

})();



/* Student cards spreedsheet
Names: B4 - B?
Doc: E4 - ?
Birth F4 - ?
Register: G4 - ?
*/



/* Student records spreadsheet:

Names: B3 - B50
Status C3 - B50  - expect ATIVO

Register: D3 - D50
Birth: F3 - F50
Doc: H3 - H50
*/