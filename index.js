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
    let valuesFound = [];
    const valuesToFind = [];
    let notFoundCount = 0;

    for (let i = 4; i < 600; i++) {

        const nameCell = idCardsSheet['B' + i];

        if (nameCell && nameCell.v && nameCell.v.includes('-')) {

            const name = utils.accentFold(nameCell.v.toLowerCase().split('-')[0]).trim();
            const rec = studenRecordOf(name, students);
            studentsIdsToPrintCount++;

            const docCell = idCardsSheet['E' + i];
            const birthCell = idCardsSheet['F' + i];
            const registerCell = idCardsSheet['G' + i];

            if ((!docCell || !docCell.v || docCell.v.length < 5)) {
                valuesToFind.push('doc');
                if (rec && rec.doc) {
                    idCardsSheet['E' + i] = { v: rec.doc, t: 's', w: rec.doc }
                    valuesFound.push('doc');
                    //console.log(name, 'doc', rec.doc, rec.sheet)
                }
            }
            if ((!birthCell || !birthCell.v)) {
                valuesToFind.push('birth');
                if (rec && rec.birth) {
                    idCardsSheet['F' + i] = { v: rec.birth, t: 's', w: rec.birth }
                    valuesFound.push('birth');
                    //console.log(name, 'birth', rec.birth, rec.sheet)
                }
            }
            if (!registerCell || !registerCell.v || registerCell.v.length < 6) {
                valuesToFind.push('register');
                if (rec && rec.register) {
                    idCardsSheet['G' + i] = { v: rec.register, t: 's', w: rec.register }
                    //console.log(name, 'reg', rec.register, rec.sheet)
                    valuesFound.push('register');
                }
            }

            if (!rec) {
                notFoundCount++;
                console.log('Warning: "' + nameCell.v + '" not found! ');
            }

        }

    }

    const valuesToFindCount = valuesToFind.length;
    const docToFindCount = valuesToFind.filter(cur => cur === 'doc').length;
    const regToFindCount = valuesToFind.filter(cur => cur === 'register').length;
    const birthToFindCount = valuesToFind.filter(cur => cur === 'birth').length;

    const valuesFoundCount = valuesFound.length;
    const docFoundCount = valuesFound.filter(cur => cur === 'doc').length;
    const regFoundCount = valuesFound.filter(cur => cur === 'register').length;
    const birthFoundCount = valuesFound.filter(cur => cur === 'birth').length;

    const foundCount = studentsIdsToPrintCount - notFoundCount;

    console.log('\n' + studentsIdsToPrintCount + ' student records to print id cards, ' + foundCount + ' students found (' +
        (100 * foundCount / studentsIdsToPrintCount).toFixed(1) + ' %), ' + notFoundCount + ' not found');
    console.log(valuesToFindCount + ' values to be find. ' + valuesFoundCount +
        ' (' + (100 * valuesFoundCount / valuesToFindCount).toFixed(1) + ' %) values found and completed\n');

    console.log('Missing values count: \nDoc: ' + docToFindCount + '\nRegister: ' +
        regToFindCount + '\nBirth: ' + birthToFindCount + '\n');
    console.log('Found values count: \nDoc: ' + docFoundCount + '\nRegister: ' +
        regFoundCount + '\nBirth: ' + birthFoundCount + '\n');

    XLSX.writeFile(idCardsAll, 'assets/Controle fotos carteiras estudantis _ carteirinhas de estudante __ MOD.ods');

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