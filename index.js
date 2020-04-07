const XLSX = require('xlsx');
const fs = require('fs');
const utils = require('@stefancfuchs/utils');
const natural = require('natural');
const readlineSync = require('readline-sync');

/**
 * Giver a reference string and a array of strings, returns the string from arr which is more similar to the reference string
 * @param input reference string
 * @param arr array of strings you want to find the most similar to input string
 */
const stringSimilarToAny = (input, arr) => {

    const minScore = 0.80;
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
        }, null);

    if (highestScore >= minScore) {
        return highestMatchString;
    }
    return null;
}

/**
 * Maps tag files to tabs of external spreadsheet
 * @param tag file tag string
 */
const tagToTab = (tag) => {

    tag = tag.trim();
    const change = {
        'IMEC20': 'MEC INT 2020',
        'ELM19': 'EltMec2019',
        'AUT20': 'Aut 2020',
        'AUT19': 'Aut 2019',
        'CER20': 'Cer 2020',
        'PRJ20': 'Adm 2020 -PROEJA',
        'ADM20': 'Adm2020',
        'AGRO20': 'Agro 2020',
        'MEC20': 'Mec sub 2020',
        'ELT20': 'Elet 2020',
        'AGRO20': 'Agro Sup 2020',

        'MAT20': 'Mat 2020',
        'ENG20': 'Eng Elet 2020',
    };

    const tab = change[tag];
    //if (!tab) console.log('tag not found:', tag);
    return tab;
}

/**
 * Standartizes first letters of each word as UpperCase and the rest as lowerCase
 * @param str initial string
 */
const capStart = (str) => {
    const split = str.split(' ');
    for (let p = 0; p < split.length; p++) {
        split[p] = split[p].slice(0, 1).toUpperCase() + split[p].slice(1).toLowerCase();
    }
    return split.join(' ');
}


// Main function starts here
const createStudentDefaulForms = false; // changes code purpose

(async () => {

    // Read student records spreadSheet (external spredsheet)
    const studentRecords = XLSX.readFile('assets/HISTÓRICO E CONTATO ALUNOS ATIVOS.ods');
    const sheetNames = studentRecords.SheetNames;
    const sheets = studentRecords.Sheets;

    const students = [];
    const notActiveStudents = [];

    const findCollumn = (acs, name, defaultColummn) => {
        const letters = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S'];
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
        const getVal = (acs, cell) => acs[cell] ? acs[cell].v : '';

        // find out which collumns has doc/RG, birth date and registy values
        let docCollumn = findCollumn(acs, 'RG', 'H');
        let birthCollumn = findCollumn(acs, 'NASCIMENTO', 'F');
        let regCollumn = findCollumn(acs, 'MATRÍCULA', 'D');

        let mailCollumn = findCollumn(acs, 'E-MAIL', 'J');
        let phone1Collumn = findCollumn(acs, 'TELEFONE 1', 'K');
        let phone2Collumn = findCollumn(acs, 'TELEFONE 2', 'L');
        let phone3Collumn = findCollumn(acs, 'CEL. RESPONSÁVEL', 'M');
        let parent1Collumn = findCollumn(acs, 'MÃE', 'N');
        let parent2Collumn = findCollumn(acs, 'PAI', 'O');
        let entranceAtCollumn = findCollumn(acs, 'CHAMADA', 'P');
        let entranceKindCollumn = findCollumn(acs, 'COTA', 'Q');

        //console.log(sheet, docCollumn, birthCollumn, regCollumn);

        for (let i = 3; i < 50; i++) {

            const nameCell = acs['B' + i];
            const statusCell = acs['C' + i];

            if (nameCell && nameCell.v) { // if this line has text at the name collumn

                const name = utils.accentFold(nameCell.v).toLowerCase().trim();
                const register = getVal(acs, regCollumn + i);
                const doc = getVal(acs, docCollumn + i);
                const status = statusCell && statusCell.v;
                let birth = acs[birthCollumn + i] ? acs[birthCollumn + i].w : null;
                if (acs[birthCollumn + i] && !birth.includes('/')) {
                    birth = new Date(acs[birthCollumn + i].w);
                }
                const email = getVal(acs, mailCollumn + i);
                const phone1 = getVal(acs, phone1Collumn + i);
                const phone2 = getVal(acs, phone2Collumn + i);
                const phone3 = getVal(acs, phone3Collumn + i);
                const parent1 = getVal(acs, parent1Collumn + i);
                const parent2 = getVal(acs, parent2Collumn + i);
                const entranceAt = getVal(acs, entranceAtCollumn + i);
                const entranceKind = getVal(acs, entranceKindCollumn + i);

                const nameWithTab = capStart(nameCell.v) + '" - ' + sheet;
                const newStudent = {
                    name, register, birth, doc, sheet, nameUntreated: nameCell.v, status, nameWithTab,
                    email, phone1, phone2, phone3, parent1, parent2, entranceAt, entranceKind
                }; // Creates student object

                if (status === 'ATIVO' ||
                    nameCell && !status && (sheet === 'Mat 2020' || sheet === 'Eng Elet 2020')) { // Temp fix due incomplete tabs at external .ods file

                    students.push(newStudent); // Active students only
                } else {
                    notActiveStudents.push(newStudent); // It can be a student of not "ATIVO" status
                }
            }
        }
    }

    console.log(students.length + ' active student records found at records spreadsheet\n');
    // console.log(notActiveStudents.length + ' possible students records');

    /**
     * Finds register with the same name in both local and external spredsheets 
     * @param name student to look for
     * @param students students array - maybe filtered by course, maybe active only, maybe inactive only
     * @returns student object
     */
    const studenRecordOf = (name, students) => {
        const st = students.filter(s => s.name === name);
        if (st.length > 1) {
            console.log('Warning: Two students have the name of ' + name);
            return null;
        }
        //if(st && st.length) console.log('rec found', name);
        return (st && st.length) ? st[0] : null;
    }

    // PART 2
    // Read student id cards spreadsheet
    //
    if (!createStudentDefaulForms) {

        console.log('Mode A active: Starting students id card infos completion');
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
                const tab = tagToTab(nameCell.v.split('-')[1]);
                const filteredStudents = tab ? students.filter(stu => stu.sheet === tab) : students;
                //if(!filteredStudents.length) console.log(filteredStudents.length, 'tab filter', tab);

                let rec = studenRecordOf(name, filteredStudents);

                if (!rec) { // If no match was found using exact name match at expected tab, tries using all other tabs data

                    const maybeMatch = studenRecordOf(name, students); // Student match from unexpected course tab

                    if (maybeMatch) {

                        // Asks user if it should use this match data
                        const response = readlineSync.question("Use '" + maybeMatch.nameWithTab + "' info for '" + capStart(nameCell.v) + "'?  (Y/N)  __");
                        if (response === 'y' || response === 'Y') {
                            rec = maybeMatch;
                            console.log('"' + capStart(nameCell.v) + '" info completed!\n');
                        } else {
                            console.log('"' + capStart(nameCell.v) + '" skipped\n');
                        }
                    }
                }

                studentsIdsToPrintCount++;

                const docCell = idCardsSheet['E' + i]; // This sheet has defined collumns since its controlled by me
                const birthCell = idCardsSheet['F' + i];
                const registerCell = idCardsSheet['G' + i];

                // Completes data at local spreedsheet if data is missing and external spreedsheet has this data
                if ((!docCell || !docCell.v || docCell.v.length < 5)) {
                    valuesToFind.push('doc');
                    if (rec && rec.doc) {
                        idCardsSheet['E' + i] = { v: rec.doc, t: 's', w: rec.doc }
                        valuesFound.push('doc');
                    }
                }
                if ((!birthCell || !birthCell.v)) {
                    valuesToFind.push('birth');
                    if (rec && rec.birth) {
                        idCardsSheet['F' + i] = { v: rec.birth, t: 's', w: rec.birth }
                        valuesFound.push('birth');
                    }
                }
                if (!registerCell || !registerCell.v || registerCell.v.length < 6) {
                    valuesToFind.push('register');
                    if (rec && rec.register) {
                        idCardsSheet['G' + i] = { v: rec.register, t: 's', w: rec.register }
                        valuesFound.push('register');
                    }
                }

                if (!rec) { // If no match was found for this student

                    const split = nameCell.v.split('-');

                    let invertedName = split.reverse().join(' - ');
                    notFoundNames.push(invertedName.trim());

                    /**
                     * Finds register with the similar name in both local and external spredsheets (similar code of studenRecordOf)
                     * @param name student to look for
                     * @param students students array - maybe filtered by course, maybe active only, maybe inactive only
                     * @returns string to inform this code user about possible match
                     */
                    const findSimilarName = (fName, fStudents) => {
                        const fSimilarName = stringSimilarToAny(fName, fStudents.map(cur => cur.name));
                        let fSimilarCompleted = null;
                        for (let s = 0; s < fStudents.length; s++) {
                            if (fStudents[s].name === fSimilarName) {
                                fSimilarCompleted = capStart(fStudents[s].nameUntreated) + '" - ' + fStudents[s].sheet;
                            }
                        }
                        return fSimilarCompleted;
                    }

                    let similarCompleted = findSimilarName(name, filteredStudents); // Looks for similar names at the same course
                    if (!similarCompleted) {
                        similarCompleted = findSimilarName(name, students); // Looks for similar names between all courses
                    }

                    didYouMean.push(similarCompleted);
                }

            }

        }

        const notFoundCount = notFoundNames.length;
        // Gives sugestions of matches, since it could be erros like typos,
        // missing last names or students in different courses in each spreadsheet
        notFoundNames = notFoundNames.map((cur, i) => {
            let name = cur;
            if (didYouMean[i]) {
                name = name + '. Did you mean "' + didYouMean[i] + '?';
            } else { // look in inactive students
                const inactiveRecord = studenRecordOf(name, notActiveStudents);
                if (inactiveRecord) {
                    name = name + '. Status: ' + inactiveRecord.status + ' (' + inactiveRecord.sheet + ')';
                }
            }
            return name;
        });
        notFoundNames = notFoundNames.sort();

        if (notFoundCount) { // Prints not found cases to console
            console.log();
            for (let x = 0; x < notFoundCount; x++) {
                console.log('Not found: ' + notFoundNames[x]);
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
            regFoundCount + '\nBirth: ' + birthFoundCount);

        XLSX.writeFile(idCardsAll, 'assets/Controle fotos carteiras estudantis _ carteirinhas de estudante __ COMPLETED.ods');

    } else {

        // Mooded part 2: Create Student spreadsheets
        console.log('Mode B active: Generate students spreadSheets');

        const selectedSheets = ['Aut 2020', 'MEC INT 2020',
            'Adm2020', 'Cer 2020', 'Mec sub 2020',
            'Eng Elet 2020', 'Agro Sup 2020'];
        const model = XLSX.readFile('assets/Student default form.xlsx', { cellStyles: true, type: 'file', cellHTML: true, dense: true, sheetStubs: true, });
        if (!fs.existsSync('outputs')) {
            fs.mkdirSync('outputs');
        }

        for (const sheet of selectedSheets) {

            const classStudents = students.filter(std => std.sheet === sheet);
            if (!fs.existsSync('outputs/' + sheet)) {
                fs.mkdirSync('outputs/' + sheet);
            }

            for (const student of classStudents) {

                const newSheet = { ...model };
                const dataSheet = newSheet.Sheets['Dados'];
                dataSheet['A2'] = { v: student.name, t: 's', w: student.name } // name
                dataSheet['B2'] = { v: student.birth, t: 's', w: student.birth } //birth
                dataSheet['C2'] = { v: student.email, t: 's', w: student.email } //mail
                dataSheet['D2'] = { v: student.phone1, t: 's', w: student.phone1 } //phone1
                dataSheet['E2'] = { v: student.phone2, t: 's', w: student.phone2 } //phone2
                dataSheet['F2'] = { v: student.phone3, t: 's', w: student.phone3 } // resp. cell phone
                dataSheet['G2'] = { v: student.parent1, t: 's', w: student.parent1 } // parent1
                dataSheet['H2'] = { v: student.parent2, t: 's', w: student.parent2 } // parent2
                dataSheet['I2'] = { v: student.entranceAt, t: 's', w: student.entranceAt } // entry time
                dataSheet['J2'] = { v: student.entranceKind, t: 's', w: student.entranceKind } // entry type

                newSheet.Sheets['Dados'] = dataSheet;
                XLSX.writeFile(newSheet, 'outputs/' + sheet + '/' + student.nameUntreated + '.xlsx', { cellStyles: true, type: 'file', cellHTML: true, dense: true, sheetStubs: true, });
            }
            console.log('Class ' + sheet + ': ' + classStudents.length + ' spreadsheets saved');

        }

    }

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