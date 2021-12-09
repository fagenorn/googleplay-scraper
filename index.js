import Excel from "exceljs";
import {
  list,
  category as _category,
  collection as _collection,
  app as _app,
  search,
} from "google-play-scraper";

var workbook = new Excel.Workbook();

const sheet = workbook.addWorksheet("GAMES - NEW FREE", {
  properties: { tabColor: { argb: "E2E2E2E" } },
});

sheet.columns = [
  {
    header: "Title",
    key: "title",
    width: 32,
  },
  { header: "Summary", key: "summary", width: 52, outlineLevel: 2 },
  { header: "Installs", key: "installs", width: 32 },
  {
    header: "Max Installs",
    key: "maxInstalls",
    width: 32,
    outlineLevel: 1,
    filterButton: true,
    style: { numFmt: "#,##0_);(#,##0)" },
  },
  { header: "Score", key: "score", width: 10, outlineLevel: 2 },
  { header: "Ratings", key: "ratings", width: 32, outlineLevel: 2 },
  { header: "Reviews", key: "reviews", width: 32, outlineLevel: 2 },
  { header: "Free", key: "free", width: 10, outlineLevel: 2 },
  { header: "Developer", key: "developer", width: 32, outlineLevel: 1 },
  {
    header: "Developer Email",
    key: "developerEmail",
    width: 32,
  },
  {
    header: "Developer Website",
    key: "developerWebsite",
    width: 32,
    outlineLevel: 1,
  },
  {
    header: "Developer Address",
    key: "developerAddress",
    width: 32,
    outlineLevel: 1,
  },
  { header: "Url", key: "url", width: 32 },
];

const colBorderStyle = {
  left: { style: "thin" },
  bottom: { style: "thin" },
  right: { style: "thin" },
};

const colFontStyle = {
  name: "Calibri",
  size: 12,
  bold: true,
};

const colAlignmentStyle = { vertical: "center", horizontal: "center" };

const colFillStyle = {
  type: "gradient",
  gradient: "path",
  center: { left: 0.5, top: 0.5 },
  stops: [
    { position: 0, color: { argb: "FFCC99" } },
    { position: 1, color: { argb: "FFB66D" } },
  ],
};

sheet.getRow(1).border = colBorderStyle;
sheet.getRow(1).font = colFontStyle;
sheet.getRow(1).alignment = colAlignmentStyle;
sheet.getRow(1).fill = colFillStyle;

search({
  term: "control",
  price: "free",
  num: 250,
}).then(async (apps) => {
  // list({
  //   category: _category.APPLICATION,
  //   collection: _collection.NEW_FREE,
  //   num: 250,
  // }).then(async (apps) => {
  for (let index = 0; index < apps.length; index++) {
    const app = apps[index];
    await _app({ appId: app.appId }).then((details) => {
      sheet.addRow([
        details.title,
        details.summary,
        details.installs,
        details.maxInstalls,
        details.score,
        details.ratings,
        details.ratings,
        details.free,
        details.developer,
        details.developerEmail,
        details.developerWebsite,
        details.developerAddress,
        details.url,
      ]);
    });
  }

  await workbook.xlsx.writeFile("10.xlsx");
});
