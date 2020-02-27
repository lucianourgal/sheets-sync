const XLSX = require('xlsx');
const utils = require('@stefancfuchs/utils');
const natural = require('natural');

const stringSimilarToAny = (input, arr) => {

    const minScore = 0.87;
    let highestScore = 0;

    const highestMatchString = arr
        .reduce((prev, cur) => {
            const curNorm = utils.accentFold(cur.toLowerCase().trim())
            const score = natural.JaroWinklerDistance(utils.accentFold(input).toLowerCase().trim(), curNorm)
            if (score > highestScore) {
                highestScore = score
                return cur;
            }
            return prev;
        });

    if (highestScore >= minScore) {
        return highestMatchString;
    }
    return null;
}



(async () => {

    // Read Student records spreadSheet
    const studentRecords = XLSX.readFile('assets/HISTÓRICO E CONTATO ALUNOS ATIVOS.ods');
    const sheetNames = studentRecords.SheetNames;
    const sheets = studentRecords.Sheets;
    const students = [];

    const findCollumn = (acs, name, defaultColummn) => {
        const letters = ['D', 'E', 'F', 'G', 'H', 'I', 'J'];
        let c = defaultColummn;
        for (let l = 0; l < letters.length; l++) {
            const check = letters[l];
            if (acs[check + '2'] && acs[check + '2'].v === name) {
                c = check;
            }
        }
        return c;
    }

    console.log(sheetNames.length + ' sheets in students records spreadsheet');

    for (let sheet of sheetNames) {
        const acs = sheets[sheet];

        // find out which collumns has doc/RG values
        let docCollumn = findCollumn(acs, 'RG', 'H');
        let birthCollumn = findCollumn(acs, 'NASCIMENTO', 'F');
        let regCollumn = findCollumn(acs, 'MATRÍCULA', 'D');
        //console.log(sheet, docCollumn, birthCollumn, regCollumn);

        for (let i = 3; i < 50; i++) {

            const nameCell = acs['B' + i];
            const statusCell = acs['C' + i];

            if (nameCell && statusCell && statusCell.v === 'ATIVO') {

                const name = utils.accentFold(nameCell.v).toLowerCase().trim();
                const register = acs[regCollumn + i] ? acs[regCollumn + i].v : null;

                let birth = acs[birthCollumn + i] ? acs[birthCollumn + i].w : null;
                if (acs[birthCollumn + i] && !birth.includes('/')) {
                    birth = new Date(acs[birthCollumn + i].w);
                }

                const docCell = acs[docCollumn + i];
                const doc = docCell ? docCell.v : null;

                students.push({ name, register, birth, doc, sheet, nameUntreated: nameCell.v });
            }
        }
    }

    console.log(students.length + ' active student records found at records spreadsheet\n');

    const studenRecordOf = (name, students) => {
        const st = students.filter(s => s.name === name);
        if (st.length > 1) {
            console.log('Warning: Two students have the name of ' + name);
            return null;
        }
        //if(st && st.length) console.log('rec found', name)
        return (st && st.length) ? st[0] : null;
    }

    const studentsAvailableNames = students.map(cur => cur.name);

    // Read student id cards spreadsheet
    const idCardsAll = XLSX.readFile('assets/Controle fotos carteiras estudantis _ carteirinhas de estudante  2020.ods');
    let idCardsSheet = idCardsAll.Sheets['Falta imprimir'];

    let studentsIdsToPrintCount = 0;
    let valuesFound = [];
    const valuesToFind = [];

    let notFoundNames = [];
    const didYouMean = [];

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
                const split = nameCell.v.split('-');
                let invertedName = split.reverse().join(' - ');

                notFoundNames.push(invertedName.trim());
                const similarName = stringSimilarToAny(name, studentsAvailableNames);
                let similarCompleted = null;
                for(let s=0;s<students.length;s++) {
                    if(students[s].name === similarName) {
                        similarCompleted = students[s].nameUntreated + ' - ' + students[s].sheet;
                    }
                }

                didYouMean.push(similarCompleted);
                //console.log('Warning: "' + nameCell.v + '" not found! ');
            }

        }

    }

    const notFoundCount = notFoundNames.length;
    notFoundNames = notFoundNames.map((cur, i) => {
        let name = cur;
        if (didYouMean[i]) {
            name = name + '. Did you mean "' + didYouMean[i] + '"?';
        }
        return name;
    });
    notFoundNames = notFoundNames.sort();

    if (notFoundCount) {
        console.log();
    }
    for (let x = 0; x < notFoundCount; x++) {
        console.log('Not found: ' + notFoundNames[x]);
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
        regFoundCount + '\nBirth: ' + birthFoundCount);

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