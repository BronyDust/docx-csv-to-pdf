const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const fs = require("fs");
const path = require("path");
const CsvReadableStream = require("csv-reader");
const { Transform } = require("stream");

const template = fs.readFileSync(
  path.resolve(__dirname, "template.docx"),
  "binary"
);
const templateEmpty = fs.readFileSync(
  path.resolve(__dirname, "template-empty.docx"),
  "binary"
);

const inputStream = fs.createReadStream("data.csv", "utf-8");

fs.rmSync("./docx", { recursive: true, force: true });
fs.mkdirSync("./docx", { recursive: true });

fs.rmSync("./result", { recursive: true, force: true });
fs.mkdirSync("./result", { recursive: true });

console.log("START");

// converting docx to pdf stream

const COUNT = 507;
let counter = 0;

// Transform stream that takes csv data and converts it to docx
// and then to pdf
const pdfStream = new Transform({
  objectMode: true,
  transform: async (row, encoding, callback) => {
    const templateZip = new PizZip(template);
    const templateEmptyZip = new PizZip(templateEmpty);

    const [NAME, PLACE, NOMINATOIN, TEACHER] = row;

    const isEmpty = !NOMINATOIN || !TEACHER;
    const doc = new Docxtemplater(isEmpty ? templateEmptyZip : templateZip, {
      paragraphLoop: true,
      linebreaks: true,
      errorLogging: true,
    });

    doc.render(
      isEmpty
        ? { NAME, PLACE }
        : {
            NAME,
            PLACE,
            NOMINATOIN,
            TEACHER,
          }
    );

    const buffer = doc.getZip().generate({
      type: "nodebuffer",
    });
    const fileName = `${isEmpty ? `e-${counter}` : counter}-${NAME.split(" ")
      .join("-")
      .toLowerCase()}`;

    // const html = await mammoth.convertToHtml({ buffer });

    fs.writeFileSync(
      path.resolve(__dirname, "result", `${fileName}.docx`),
      buffer
    );

    counter += 1;
    console.log(`PROCESS ${counter}/${COUNT}`);

    callback(null, null);
  },
});

inputStream
  .pipe(new CsvReadableStream({ trim: true, delimiter: ";", skipHeader: true }))
  .pipe(pdfStream)
  .on("end", () => {
    console.log("DONE");
  });
