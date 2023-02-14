require("express-async-errors");
// require("dotenv").config();
const express = require("express");
const app = express();
const path = require("path");

app.set("views", path.join(__dirname, "./views"));
app.set("view engine", "ejs");
app.use(express.json());
// Public
app.use("/assets", express.static(path.join(__dirname, "assets")));
app.use("/test", require("./routes/document"));

app.get("/", (req, res) => {
  res.render("home");
});

//port
const port = process.env.PORT || 4000;

//initialize server
const start = async () => {
  try {
    app.listen(port, console.log(`server is listening on port ${port}`));
  } catch (error) {
    console.error("Unable to connect to the database:", error);
  }
};

start();
