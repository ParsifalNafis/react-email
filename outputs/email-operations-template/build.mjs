import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = new URL(".", import.meta.url).pathname;
const workbook = Workbook.create();

const navy = "#263238";
const ink = "#1F2933";
const slate = "#607D8B";
const line = "#D8E0E4";
const paper = "#F7F9FA";
const amber = "#FFF3CD";
const green = "#E8F5E9";
const red = "#FDECEC";
const font = "Arial";

function setupSheet(sheet, title, subtitle, headers, widths) {
  sheet.showGridLines = false;
  sheet.getRange("A1:O1").merge();
  sheet.getRange("A1").values = [[title]];
  sheet.getRange("A1").format = {
    font: { name: font, size: 15, bold: true, color: ink },
    verticalAlignment: "center",
  };
  sheet.getRange("A1").format.rowHeight = 28;
  sheet.getRange("A2:O2").merge();
  sheet.getRange("A2").values = [[subtitle]];
  sheet.getRange("A2").format = {
    font: { name: font, size: 10, italic: true, color: slate },
    verticalAlignment: "center",
  };
  sheet.getRange("A2").format.rowHeight = 20;
  sheet.getRange("A4").values = [["Editable fields are highlighted in pale yellow. Keep Contact ID stable when a person appears on more than one tab."]];
  sheet.getRange("A4").format = { font: { name: font, size: 9, color: slate, italic: true } };
  const headerRange = sheet.getRangeByIndexes(5, 0, 1, headers.length);
  headerRange.values = [headers];
  headerRange.format = {
    fill: navy,
    font: { name: font, size: 10, bold: true, color: "#FFFFFF" },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "outside", style: "thin", color: navy },
  };
  headerRange.format.rowHeight = 30;
  const inputRange = sheet.getRangeByIndexes(6, 0, 199, headers.length);
  inputRange.format = {
    font: { name: font, size: 10, color: ink },
    verticalAlignment: "center",
    fill: amber,
  };
  inputRange.format.borders = { preset: "insideHorizontal", style: "thin", color: line };
  inputRange.format.rowHeight = 20;
  widths.forEach((width, index) => {
    sheet.getRangeByIndexes(0, index, 205, 1).format.columnWidth = width;
  });
  sheet.freezePanes.freezeRows(6);
}

const general = workbook.worksheets.add("General Mailing List");
const generalHeaders = [
  "Contact ID", "Email", "First name", "Last name", "Organization", "Role",
  "Source", "Consent status", "Subscribed on", "Preferred content", "Lifecycle status",
  "Last email sent", "Unsubscribed on", "Notes"
];
setupSheet(
  general,
  "General Mailing List",
  "Use this for people who have consented to receive general product, blog, and company updates.",
  generalHeaders,
  [14, 28, 16, 16, 22, 18, 18, 18, 16, 20, 18, 16, 18, 30]
);
general.getRange("I7:I205").format.numberFormat = "yyyy-mm-dd";
general.getRange("L7:M205").format.numberFormat = "yyyy-mm-dd";
general.getRange("H7:H205").dataValidation = { rule: { type: "list", values: ["Subscribed", "Unsubscribed", "Pending consent", "Suppressed"] } };
general.getRange("J7:J205").dataValidation = { rule: { type: "list", values: ["All updates", "Blog posts", "Product updates", "Events only"] } };
general.getRange("K7:K205").dataValidation = { rule: { type: "list", values: ["Active", "Paused", "Bounced", "Complained"] } };
general.getRange("H7:H205").conditionalFormats.add("containsText", { text: "Subscribed", format: { fill: green, font: { color: "#1B5E20" } } });
general.getRange("H7:H205").conditionalFormats.add("containsText", { text: "Unsubscribed", format: { fill: red, font: { color: "#B71C1C" } } });

const weekly = workbook.worksheets.add("Shrine Dev Weekly");
const weeklyHeaders = [
  "Contact ID", "Email", "Name", "Organization", "Role", "Timezone", "Invite status",
  "Attendance preference", "First invited", "Last invited", "Last attended", "Engagement status", "Notes"
];
setupSheet(
  weekly,
  "Shrine Dev Weekly",
  "Roster for the recurring developer meeting. Contact ID links this list to the attendance log.",
  weeklyHeaders,
  [14, 28, 24, 22, 18, 16, 18, 22, 16, 16, 16, 18, 32]
);
weekly.getRange("I7:K205").format.numberFormat = "yyyy-mm-dd";
weekly.getRange("G7:G205").dataValidation = { rule: { type: "list", values: ["Invited", "RSVP yes", "RSVP no", "Waitlist", "Do not invite"] } };
weekly.getRange("H7:H205").dataValidation = { rule: { type: "list", values: ["Weekly", "Monthly digest", "Notes only", "Do not email"] } };
weekly.getRange("L7:L205").dataValidation = { rule: { type: "list", values: ["Active", "New", "Inactive", "Paused"] } };
weekly.getRange("G7:G205").conditionalFormats.add("containsText", { text: "Do not invite", format: { fill: red, font: { color: "#B71C1C" } } });

const attendance = workbook.worksheets.add("Weekly Attendance");
const attendanceHeaders = [
  "Meeting date", "Meeting title", "Contact ID", "Email", "Name", "RSVP", "Attendance",
  "Follow-up status", "Follow-up completed", "Notes"
];
setupSheet(
  attendance,
  "Weekly Attendance",
  "One row per person per meeting. Use this as the source for a later post-meeting follow-up workflow.",
  attendanceHeaders,
  [16, 32, 14, 28, 24, 14, 16, 20, 20, 34]
);
attendance.getRange("A7:A205").format.numberFormat = "yyyy-mm-dd";
attendance.getRange("I7:I205").format.numberFormat = "yyyy-mm-dd";
attendance.getRange("F7:F205").dataValidation = { rule: { type: "list", values: ["Yes", "No", "No response"] } };
attendance.getRange("G7:G205").dataValidation = { rule: { type: "list", values: ["Attended", "Absent", "Cancelled", "Unknown"] } };
attendance.getRange("H7:H205").dataValidation = { rule: { type: "list", values: ["Not needed", "Ready to send", "Sent", "Submitted", "Follow up"] } };
attendance.getRange("G7:G205").conditionalFormats.add("containsText", { text: "Attended", format: { fill: green, font: { color: "#1B5E20" } } });
attendance.getRange("G7:G205").conditionalFormats.add("containsText", { text: "Absent", format: { fill: red, font: { color: "#B71C1C" } } });

attendance.getRange("L1:M1").merge();
attendance.getRange("L1").values = [["Attendance summary"]];
attendance.getRange("L1").format = { fill: navy, font: { name: font, size: 10, bold: true, color: "#FFFFFF" }, horizontalAlignment: "center" };
attendance.getRange("L2:M5").values = [
  ["People marked attended", null],
  ["People absent", null],
  ["Ready for follow-up", null],
  ["Follow-up completed", null],
];
attendance.getRange("M2:M5").formulas = [
  ['=COUNTIF($G$7:$G$205,"Attended")'],
  ['=COUNTIF($G$7:$G$205,"Absent")'],
  ['=COUNTIF($H$7:$H$205,"Ready to send")'],
  ['=COUNTIF($H$7:$H$205,"Submitted")'],
];
attendance.getRange("L2:M5").format = { font: { name: font, size: 10, color: ink }, borders: { preset: "all", style: "thin", color: line } };
attendance.getRange("L2:L5").format.fill = paper;
attendance.getRange("M2:M5").format = { fill: "#E3F2FD", font: { name: font, size: 10, bold: true, color: ink }, horizontalAlignment: "center", borders: { preset: "all", style: "thin", color: line } };
attendance.getRange("L1:M5").format.columnWidth = 20;

workbook.recalculate();

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(`${outputDir}/email-operations.xlsx`);

for (const sheetName of ["General Mailing List", "Shrine Dev Weekly", "Weekly Attendance"]) {
  const preview = await workbook.render({ sheetName, autoCrop: "all", scale: 1.2, format: "png" });
  await fs.writeFile(`${outputDir}/${sheetName.replaceAll(" ", "-").toLowerCase()}.png`, new Uint8Array(await preview.arrayBuffer()));
}

const overview = await workbook.inspect({ kind: "workbook,sheet", maxChars: 4000 });
const attendanceCheck = await workbook.inspect({ kind: "table", range: "Weekly Attendance!A1:M12", include: "values,formulas", tableMaxRows: 12, tableMaxCols: 13 });
const errors = await workbook.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!", options: { useRegex: true, maxResults: 100 }, summary: "final formula error scan" });
await fs.writeFile(`${outputDir}/verification.txt`, `${overview.ndjson}\n${attendanceCheck.ndjson}\n${errors.ndjson}`);
