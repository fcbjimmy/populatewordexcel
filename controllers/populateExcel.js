const XlsxPopulate = require("xlsx-populate");
const fs = require("fs");
const path = require("path");

const populateExcel = async (req, res, next) => {
  // Load an existing workbook
  const thePath = path.resolve(__dirname, "../test", "Book1.xlsx");
  console.log(thePath, "here");
  XlsxPopulate.fromFileAsync(thePath).then((workbook) => {
    // Modify the workbook.
    // const value = workbook.sheet("Sheet1").cell("A1").value();

    // Log the value.

    workbook.sheet("Sheet1").cell("A1").value("This is neat!");
    // const cell = workbook.definedName("test", "this is it");
    // console.log(cell);
    // const sheet = cell.sheet();
    const value = workbook.definedName("date", "14-2-2023");
    workbook.sheet("Sheet1").definedName("name", "Jimmy");
    const cell = workbook.definedName("test");
    console.log(cell);
    // workbook.sheet("Sheet1").definedName("date").value(date);
    // workbook.sheet("Sheet1").definedName("name").value("foo");
    // workbook.sheet("Sheet1").definedName("price").value("$20");
    // Write to file.
    workbook.toFileAsync(path.resolve(__dirname, "../test", "out.xlsx"));
    console.log("Finish to fill in Microsoft Excel document.");
    res.render("success", { title: "Excel" });
  });
};

module.exports = { populateExcel };
