var express = require("express");
var router = express.Router();
const { wordFile, downloadWord } = require("../controllers/wordFile");
const { populateExcel } = require("../controllers/populateExcel");
/* GET home page. */
router.get("/", function (req, res, next) {
  console.log("hi");
  res.render("test", { title: "Test" });
});

router.get("/wordDocument", wordFile);
// router.get("/powerpoint", powerpointSlides);
// router.get("/excel", excel);

//download
router.get("/downloadWord", downloadWord); //download word document

//excel

router.get("/excel", populateExcel); //download excel document

module.exports = router;
