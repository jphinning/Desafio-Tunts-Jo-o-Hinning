const {google} = require('googleapis')

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
    
    console.log("Getting SpreadSheet Data...")

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

const getSpreadsheetTotalClasses = async () => {

    const data = await readSpreadsheet()
    let getNumberFromString = 0
    const spreadSheetTotalClasses = data.values.splice(1,1)

    const stringOfTotalClasses = spreadSheetTotalClasses[0]

    stringOfTotalClasses.forEach(classes => {
        getNumberFromString = classes.match(/\d+/)[0] 
    })

    return getNumberFromString
}

const getSpreadsheetStudentInformation = async () => {

    const data = await readSpreadsheet()

    const spreadSheetWithNameAndGrades = data.values.splice(3)

    return spreadSheetWithNameAndGrades

}

const calculateAverageGrade = async () => {

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

const studentsAbsences = async () => { 

    //Storing all absences
    let absences = []

    const studentInformation = await getSpreadsheetStudentInformation()

    studentInformation.forEach(student => {

        //Getting the individual absences for each student
        const spreadSheetWithAbsences = student.splice(2,1)

        absences.push(spreadSheetWithAbsences[0])
    })

    return absences
}


const determineStudentSituation = async (averageGrades, totalClasses) => {

    const studentAbsences = await studentsAbsences()
    
    let studentSituation = []

    averageGrades.forEach( (grade, index) => {
        
        if (parseInt(studentAbsences[index]) > (totalClasses * 0.25)){
            studentSituation.push(["Reprovado por Falta", "0"])
        }
        else if (grade < 50) {
            studentSituation.push(["Reprovado por nota", "0"])
        }
        else  if (grade >= 50 && grade < 70) {
            const NAF = 100 - grade
            studentSituation.push(["Exame Final", NAF])
        }
        else if (grade >= 70){
            studentSituation.push(["Aprovado", "0"])
        }
        
    })

    console.log(studentSituation)

    await writeSpreadsheet(studentSituation)
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