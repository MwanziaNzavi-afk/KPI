
function setupKpiRemarks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rosterSheet = ss.getSheetByName("Roster") || ss.insertSheet("Roster");
  if (rosterSheet.getLastRow() === 0) {
    rosterSheet.appendRow(["Name","Department","Group","Supervisor","Supervisor Email"]);
  }

  // Paste/import roster values into this sheet before running setupKpiRemarks.
  // The form generator reads these rows.
  const data = rosterSheet.getDataRange().getValues();
  if (data.length < 2) {
    throw new Error("Add roster rows to the Roster sheet first, then run setupKpiRemarks again.");
  }

  const form = FormApp.create("Daily KPI Supervisor Remarks");
  form.setDescription("Daily supervisor remarks for KPI bonus tracking. Submit one response per employee.");

  const workDate = form.addDateItem().setTitle("Work Date").setRequired(true);
  const groupItem = form.addListItem().setTitle("Group").setRequired(true);
  const memberItem = form.addListItem().setTitle("Team Member").setRequired(true);
  const supervisorItem = form.addTextItem().setTitle("Supervisor").setRequired(true);
  const remarkItem = form.addParagraphTextItem().setTitle("Remark").setRequired(true);

  // The form's response destination will be this spreadsheet.
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  const groups = [...new Set(data.slice(1).map(r => r[2]).filter(String))].sort();
  groupItem.setChoiceValues(groups);

  // Team Member choices are all roster names. The Python app validates group membership.
  const names = [...new Set(data.slice(1).map(r => r[0]).filter(String))].sort();
  memberItem.setChoiceValues(names);

  const out = ss.getSheetByName("Supervisor Links") || ss.insertSheet("Supervisor Links");
  out.clear();
  out.appendRow(["Group","Supervisor","Supervisor Email","Prefilled Form Link"]);

  // Prefilled links are generated per group/supervisor row.
  for (let i=1; i<data.length; i++) {
    const [name, dept, group, supervisor, email] = data[i];
    if (!group) continue;
    const response = form.createResponse();
    response.withItemResponse(groupItem.createResponse(String(group)));
    response.withItemResponse(supervisorItem.createResponse(String(supervisor || "")));
    const url = response.toPrefilledUrl();
    out.appendRow([group, supervisor || "", email || "", url]);
  }

  const formLinks = ss.getSheetByName("Form Links") || ss.insertSheet("Form Links");
  formLinks.clear();
  formLinks.appendRow(["Google Form", "Link"]);
  formLinks.appendRow(["Public response form", form.getPublishedUrl()]);
  formLinks.appendRow(["Form editor", form.getEditUrl()]);

  PropertiesService.getScriptProperties().setProperty("KPI_REMARKS_FORM_URL", form.getPublishedUrl());

  Logger.log("Form edit URL: " + form.getEditUrl());
  Logger.log("Form public URL: " + form.getPublishedUrl());
}

function doGet() {
  const url = PropertiesService.getScriptProperties().getProperty("KPI_REMARKS_FORM_URL");
  const content = url
    ? `<p><a href="${url}" target="_blank">Open the Daily KPI Supervisor Remarks form</a></p>`
    : "<p>Paste roster rows into the Roster sheet, then run setupKpiRemarks from the Apps Script editor.</p>";
  return HtmlService.createHtmlOutput(content).setTitle("KPI Supervisor Remarks");
}
