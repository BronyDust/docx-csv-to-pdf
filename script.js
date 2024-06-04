const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const fs = require("fs");
const path = require("path");
const CsvReadableStream = require("csv-reader");
const { Transform } = require("stream");

const [dataFilePath, templateFilePath] = process.argv.slice(2);

const template = fs.readFileSync(templateFilePath, "binary");
const inputStream = fs.createReadStream(dataFilePath, "utf-8");

fs.rmSync("./docx", { recursive: true, force: true });
fs.mkdirSync("./docx", { recursive: true });

fs.rmSync("./result", { recursive: true, force: true });
fs.mkdirSync("./result", { recursive: true });

console.log("ЭТАП ГЕНЕРАЦИИ ПО ШАБЛОНУ");

// converting docx to pdf stream
let counter = 0;

// Transform stream that takes csv data and converts it to docx
// and then to pdf
const pdfStream = new Transform({
  objectMode: true,
  transform: async (row, _encoding, callback) => {
    const templateZip = new PizZip(template);

    const doc = new Docxtemplater(templateZip, {
      paragraphLoop: true,
      linebreaks: true,
      errorLogging: true,
    });

    doc.render(row);

    const buffer = doc.getZip().generate({
      type: "nodebuffer",
    });
    const fileName = `${row.NOMINATION.replace(
      "Изобразительное искусство",
      "ИЗО"
    ).replace("Декоративно-прикладное искусство", "ДПИ")} ${row.NAME}`;

    fs.writeFileSync(
      path.resolve(__dirname, "docx", `${fileName}.docx`),
      buffer
    );

    counter += 1;
    console.log(`ГЕНЕРИРОВАНО ${counter}`);

    callback(null, null);
  },
});

inputStream
  .pipe(new CsvReadableStream({ trim: true, delimiter: ";", asObject: true }))
  .pipe(pdfStream)
  .on("finish", () => {
    console.log("ГЕНЕРАЦИЯ ЗАВЕРШЕНА");
    console.log("--------------------------------------------");
  });
