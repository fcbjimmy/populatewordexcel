const XLSX = require("xlsx");

const workbook = XLSX.readFile("scores.xlsx");
const worksheet = workbook.Sheets["Sheet1"];

const arrStudents = XLSX.utils.sheet_to_json(worksheet);
const highSchoolData = {}; // {highSchool: {numStudents: 0, cumalativeScore: 0}}
// Fill out hsData

for (const student of arrStudents) {
  const highSchool = student["High School"]; //obj['property']
  const studentAverage = student["Average"];

  if (highSchool in highSchoolData === false) {
    highSchoolData[highSchool] = { numStudents: 0, cumalativeScore: 0 };
  }

  highSchoolData[highSchool].numStudents += 1;
  highSchoolData[highSchool].cumalativeScore += studentAverage;
}

// // Log out average score for each high school using highSchoolData
for (const highSchool of Object.keys(highSchoolData)) {
  const highSchoolAverage =
    highSchoolData[highSchool].cumalativeScore /
    highSchoolData[highSchool].numStudents;
  console.log(`The average score for ${highSchool} is ${highSchoolAverage}`);
}
