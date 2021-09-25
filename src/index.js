const {google} = require('googleapis')

/**
 * Google API Object used for authenticating reading and writing spreadsheet functions
 * 
 * @returns {object} Important parameters to access the spreadsheet
 */
const googleApiObject = async () => {

    const auth = new google.auth.GoogleAuth({
        keyFile: "jsonKey.json", 
        scopes: "https://www.googleapis.com/auth/spreadsheets", 
    })
    
    const authClientObject = await auth.getClient()
    
    const googleSheetsInstance = google.sheets({ version: "v4", auth: authClientObject })
    
    const spreadsheetId = "1Ffw7x9431n7_NgjDufExbsBXMZ0c5QZvvyGqlw-Z0Sg";
    
    return {
        'auth' : auth,
        'googleSheetsInstance' : googleSheetsInstance,
        'spreadsheetId' : spreadsheetId,
    }
}

const readSpreadsheet = async () => {

    //Connecting to the googleAPI
    const {auth, googleSheetsInstance, spreadsheetId} = await googleApiObject()

    const readData = await googleSheetsInstance.spreadsheets.values.get({
        auth, 
        spreadsheetId, 
        range: "engenharia_de_software", 
    })

    return readData.data
}

const writeSpreadsheet = async (studentSituation) => {

    //Connecting to the googleAPI
    const {auth, googleSheetsInstance, spreadsheetId} = await googleApiObject()

    console.log("Writting the Following Information to the Spreadsheet:...")

    await googleSheetsInstance.spreadsheets.values.clear({
        spreadsheetId,
        range: "engenharia_de_software!G4:H27",
    })

    await googleSheetsInstance.spreadsheets.values.append({
        auth, 
        spreadsheetId,
        range: "engenharia_de_software!G4:H27", 
        valueInputOption: "USER_ENTERED", 
        resource: {
            values: studentSituation,
        },
    })
}

/**
 * Gets the number of classes in the semester in the spreadsheet
 * 
 * @returns {number}
 */
const getSpreadsheetTotalClasses = async () => {
    console.log("Getting Total Number of Classes in the Semester...")
    const data = await readSpreadsheet()
    let getNumberFromString = 0
    const spreadSheetTotalClasses = data.values.splice(1,1)

    const stringOfTotalClasses = spreadSheetTotalClasses[0]

    stringOfTotalClasses.forEach(classes => {
        getNumberFromString = classes.match(/\d+/)[0] 
    })

    return getNumberFromString
}

/**
 * Gets all student information, cutting the headers from the spreadsheet array 
 * 
 * @returns {Array}
 */
const getSpreadsheetStudentInformation = async () => {

    const data = await readSpreadsheet()

    const spreadSheetWithNameAndGrades = data.values.splice(3)

    return spreadSheetWithNameAndGrades

}

/**
 * Gets all grades and calculates the average in the semester
 * 
 * @returns {Array} Average grades for each student
 */
const calculateAverageGrade = async () => {

    console.log("Calculating Average Grade For Each Student...")
    //Storing all average grades
    let studentsAverageGrade = []

    const studentInformation = await getSpreadsheetStudentInformation()

    studentInformation.forEach(student => {

        //Getting the individual grades for each student
        const spreadSheetWithGrades = student.splice(3,3)

        let gradeSum = 0
        
        spreadSheetWithGrades.forEach(grade => {
            gradeSum += parseInt(grade)

        })

        const averageGrade = gradeSum / spreadSheetWithGrades.length

        studentsAverageGrade.push(Math.round(averageGrade))

    })

    return studentsAverageGrade
}

/**
 * Gets the absences row for each student
 * 
 * @returns {number} 
 */
const studentsAbsences = async () => { 

    //Storing all absences
    let absences = []
    console.log("Calculating Absences For Each Student...")
    const studentInformation = await getSpreadsheetStudentInformation()

    studentInformation.forEach(student => {

        //Getting the individual absences for each student
        const spreadSheetWithAbsences = student.splice(2,1)

        absences.push(spreadSheetWithAbsences[0])
    })

    return absences
}

/**
 * Determines wether the student Passed, Failed, or is in the Final Exam
 * Also calculates mininal passing grade on the latter situation 
 * 
 * @param {Array} averageGrades Array containg all average grades for each student
 * @param {number} totalClasses Number of classes in the semester on the spreadsheet
 */
const determineStudentSituation = async (averageGrades, totalClasses) => {

    console.log("Calculating the Student Situation in the Semester and Final Exam Minimal Grade...")
    const studentAbsences = await studentsAbsences()
    
    let studentSituation = []

    averageGrades.forEach( (grade, index) => {
        
        if (parseInt(studentAbsences[index]) > (totalClasses * 0.25)){
            studentSituation.push(["Reprovado por Falta", 0])
        }
        else if (grade < 50) {
            studentSituation.push(["Reprovado por nota", 0])
        }
        else  if (grade >= 50 && grade < 70) {
            const NAF = 100 - grade
            studentSituation.push(["Exame Final", NAF])
        }
        else if (grade >= 70){
            studentSituation.push(["Aprovado", 0])
        }
        
    })

    await writeSpreadsheet(studentSituation)
    console.log(studentSituation)
}


const main = async () => {

    try {
        const totalClasses = await getSpreadsheetTotalClasses()
        const averageGrades = await calculateAverageGrade()
        determineStudentSituation(averageGrades, totalClasses)
    }
    catch (error) {
        console.log(error)
    }
    
}

main()