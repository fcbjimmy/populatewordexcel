const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const mime = require("mime-types");
const fs = require("fs");
const path = require("path");

const data = [
  {
    name: "Destiny Aguirre",
    phone: "(451) 431-6807",
    email: "facilisis@google.com",
    address: "1162 Magna. Avenue",
  },
  {
    name: "Carla Brady",
    phone: "1-348-876-5552",
    email: "tellus.lorem.eu@yahoo.ca",
    address: "P.O. Box 664, 9104 Ultricies Rd.",
  },
  {
    name: "Hilda Dillard",
    phone: "1-333-856-2355",
    email: "a.auctor@google.org",
    address: "945-9583 Imperdiet Rd.",
  },
  {
    name: "Heather Casey",
    phone: "(548) 895-3487",
    email: "sapien.molestie@hotmail.org",
    address: "Ap #430-5703 Enim St.",
  },
  {
    name: "Kimberly Larsen",
    phone: "(466) 924-6488",
    email: "molestie.in.tempus@aol.com",
    address: "3542 Ante, Road",
  },
];

const wordFile = async (req, res, next) => {
  // Load the docx file as binary content
  const content = fs.readFileSync(
    path.resolve(__dirname, "../test", "input.docx"),
    "binary"
  );

  const zip = new PizZip(content);

  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
  });

  // Render the document (Replace {first_name} by John, {last_name} by Doe, ...)
  doc.render({
    first_name: "Jim",
    last_name: "Chan",
    phone: "123456789",
    description: "New Website",
    content:
      "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with desktop publishing software like Aldus PageMaker including versions of Lorem Ipsum.",
    date: new Date().toLocaleDateString("en-CA", { timeZone: "GMT" }),
    users: data
      .filter((user) => user.name !== "Destiny Aguirre")
      .map((user) => {
        return { name: user.name, phone: user.phone, email: user.email };
      }),
  });

  const buf = doc.getZip().generate({
    type: "nodebuffer",
    // compression: DEFLATE adds a compression step.
    // For a 50MB output document, expect 500ms additional CPU time
    compression: "DEFLATE",
  });

  // buf is a nodejs Buffer, you can either write it to a
  // file or res.send it with express for example.
  fs.writeFileSync(path.resolve(__dirname, "../test", "output.docx"), buf);
  console.log("Finish to fill in Microsoft Word document.");
  res.render("success", { title: "Test" });
};

const downloadWord = async (req, res) => {
  var file = __dirname + "/../test/output.docx";
  var filename = path.basename(file);
  var mimetype = mime.lookup(file);
  res.setHeader("Content-disposition", "attachment; filename=" + filename);
  res.setHeader("Content-type", mimetype);
  var filestream = fs.createReadStream(file);
  filestream.pipe(res);
  res.download(file);
};

module.exports = { wordFile, downloadWord };
