const excel = require("exceljs");
const http = require("http");


/**
 * 
 * @param {http.IncomingMessage} req
 * @param {excel.Workbook} workbook
 */
const requestHandler = async (req, workbook) => {
  const url = new URL(`http://localhost:5600${req.url}`);
  const pathSegments = url.pathname.split("/").slice(1);
  let sheet = workbook.getWorksheet(pathSegments[0]);
  if (!sheet) {
    return {
      error: "[excel-server]: Table named " + "'"+ pathSegments[0] +"' does not exist"
    };
  }
  let data = [];
  let fields = [];
  let limit = url.searchParams.get("limit");
  if (!limit) {
    limit = sheet.actualRowCount - 1
  }
  else {
    limit = sheet.actualRowCount == limit ? limit : 50;
  }
  console.log("limit = %s", url.searchParams.get("limit"));
  let requiredFields = url.searchParams.get("fields") ?? [];
  if (requiredFields.length !== 0) {
    requiredFields = requiredFields.split(",");
  }
  sheet.getRow(1).eachCell(c => {
    if (requiredFields.length === 0 || requiredFields.includes(c.text)) {
      fields.push(c.text);
    }
  });

  let cellRows = sheet.getRows(2, limit);
  // console.log("cellRows: ", cellRows);
  for (let i = 0; i < cellRows.length; i++) {
    let object = {};
    cellRows[i].eachCell(c => {
      if (fields.includes(sheet.getCell(1, c.col).text)) {
        object[fields.find(colkey => colkey === sheet.getCell(1, c.col).text)] = c.text;
        console.log(object);
      }
    });
    data.push(object);
  }
  return {
    data
  };
}
const main = async () => {
  const args = process.argv.slice(2);
  if (args.length < 1) {
    return console.log("[excel-server]: Please provide an excel file to serve.");
  }
  const fileLocation = args[0];
  let database = new excel.Workbook();
  try {
    database = await database.xlsx.readFile(fileLocation);
  }
  catch (e) {
    console.error(e);
    return;
  }
  const server = http.createServer((req, res) => {
    requestHandler(req, database)
    .then(respObject => {
      res.write(JSON.stringify(respObject));
      res.end();
    })
    .catch(console.error);
  });
  server.listen(5600,() => console.log(`[excel-server]: listening on 5600`));
}

main();