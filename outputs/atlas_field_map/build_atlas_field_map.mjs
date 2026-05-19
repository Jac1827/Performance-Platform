import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "/Users/jacheflin/Documents/Playground/outputs/atlas_field_map";
const outputPath = path.join(outputDir, "ATLAS_RISE_Universal_Field_Map.xlsx");

const atlasMonthlyFields = [
  "Report_Month",
  "Community_ID",
  "Community_Name",
  "Region",
  "Market",
  "Asset_Type",
  "PMS_Property_Code",
  "Total_Units",
  "Occupied_Units",
  "Leased_Units",
  "Preleased_Units",
  "Vacant_Units",
  "Physical_Occupancy_%",
  "Leased_%",
  "Economic_Occupancy_%",
  "Avg_Market_Rent",
  "Actual_Rent",
  "Concessions_$",
  "Delinquency_$",
  "Gross_Potential_Rent_$",
  "Collected_Rent_$",
  "Other_Income_$",
  "Operating_Expenses_$",
  "NOI_$",
  "Traffic",
  "Leads",
  "Tours",
  "Applications",
  "Approved_Applications",
  "Move_Ins",
  "Move_Outs",
  "Renewals_Sent",
  "Renewals_Signed",
  "Reviews_Count",
  "Avg_Review_Rating",
  "Open_Work_Orders",
  "Completed_Work_Orders",
  "Period_Status",
  "Source_System",
  "Source_File_Name",
  "Agent_Notes",
  "Last_Updated",
  "Row_Status",
  "Error_Detail",
];

const sheets = {
  README: [
    ["ATLAS RISE Universal Field Map"],
    ["Purpose", "Use this as the reusable mapping/control workbook for RISE source files before loading Atlas."],
    ["Primary monthly grain", "One row per Report_Month + Community_ID in the Atlas Monthly_Upload template."],
    ["Detail grain", "Use separate fact tabs for floor-plan, market, unit-level, and rent-roll detail. Do not flatten those into one monthly row."],
    ["Recommended upload path", "Raw files -> Source_File_Log -> Community_Aliases -> Source_Field_Map -> Atlas monthly file and optional detail tables."],
    ["What Atlas can read directly", "If Atlas only accepts the current Monthly_Upload schema, load only Target_Monthly_Upload. Detail tabs require an Atlas/Power BI model extension or import API."],
    ["Important privacy note", "Rent roll resident names, balances, and deposits are PII/sensitive. Only import those fields when Atlas has a defined need and access control."],
  ],
  Load_Order: [
    ["Step", "Object", "Grain", "Required?", "Purpose", "Notes"],
    [1, "Community_Master", "One row per community", "Yes", "Defines Community_ID and property attributes", "Update for missing communities before loading monthly facts."],
    [2, "Community_Aliases", "Many aliases per community", "Yes", "Maps raw source property names to Community_ID", "Use exact alias matching first, then reviewed fuzzy matching."],
    [3, "Source_File_Log", "One row per uploaded source file", "Yes", "Declares file type, report month, and load cadence", "The agent/importer should infer this when possible and flag exceptions."],
    [4, "Target_Monthly_Upload", "One row per Report_Month + Community_ID", "Yes", "Official Atlas dashboard load", "Matches the current Atlas Advanced Monthly Upload Template."],
    [5, "BoxScore_FloorPlanMonthly", "One row per month + community + unit type", "Optional", "Preserves box-score floor-plan performance", "Needed for floor-plan trend analysis."],
    [6, "BoxScore_PulseMonthly", "One row per month + community + unit type", "Optional", "Preserves move-ins, notices, move-outs, renewals, leased activity", "Aggregates into Monthly_Upload where applicable."],
    [7, "BoxScore_LeadMonthly", "One row per month + community + unit type", "Optional", "Preserves lead channels, follow-up activity, and tours", "Aggregates into Monthly_Upload Leads/Tours."],
    [8, "BoxScore_ApplicationMonthly", "One row per month + community + unit type", "Optional", "Preserves application/conversion funnel counts", "Aggregates into Applications and Approved_Applications."],
    [9, "BoxScore_MakeReadyMonthly", "One row per month + community + status", "Optional", "Preserves ready/not-ready vacant status", "Useful for operations but not present in current Monthly_Upload."],
    [10, "Market_PropertySnapshot", "One row per snapshot + subject/comp property", "Optional", "Stores comp market metrics", "Use snapshot date; aggregate only selected subject-property metrics to monthly row."],
    [11, "Market_FloorPlanSnapshot", "One row per snapshot + property + floor plan", "Optional", "Stores comp floor-plan rents and availability", "Use for pricing and comp detail."],
    [12, "Market_UnitSnapshot", "One row per snapshot + property + unit", "Optional", "Stores comp unit-level listing detail", "Use for pricing detail. Can be large."],
    [13, "RentRoll_UnitMonthly", "One row per month + community + unit + resident/lease segment", "Optional", "Stores unit-level occupancy and rent-roll facts", "Restrict PII fields if not needed."],
  ],
  Community_Aliases: [
    ["Community_ID", "Atlas_Community_Name", "Source_Alias", "Source_Type", "BoxScore_Total_Units_Apr_2026", "Status", "Notes"],
    ["GKP", "Glen Kernan Park", "RISE at Glen Kernan Park", "Box Score / Market / Rent Roll", 308, "Mapped", "Current Community_Master has no total units; box score suggests 308."],
    ["GKP", "Glen Kernan Park", "Rise At Glen Kernan Park", "Box Score", 308, "Mapped", "Sheet-name variant."],
    ["GKP", "Glen Kernan Park", "Glen Kernan Park", "Atlas short name", 308, "Mapped", "Short-name variant."],
    ["NOC", "Nocatee", "RISE at Nocatee", "Box Score / Rent Roll", 178, "Mapped", "Current Community_Master has no total units; box score suggests 178."],
    ["NOC", "Nocatee", "Rise At Nocatee", "Box Score", 178, "Mapped", "Sheet-name variant."],
    ["BAR", "Bartram", "RISE Bartram Park", "Box Score / Rent Roll", 297, "Mapped", "Current Community_Master has no total units; box score suggests 297."],
    ["BAY", "Baymeadows", "RISE Baymeadows", "Box Score / Market / Rent Roll", 331, "Mapped", "Current Community_Master has no total units; box score suggests 331."],
    ["CIT", "Citrus Ridge", "RISE Citrus Ridge Townhomes", "Box Score / Rent Roll", 222, "Mapped", "Current Community_Master has no total units; box score suggests 222."],
    ["DOR", "Doro", "RISE Doro", "Box Score / Rent Roll", 247, "Mapped", "Matches current Community_Master total units."],
    ["FLO", "Florence Villa", "RISE Florence Villa Townhomes", "Box Score / Rent Roll", 224, "Mapped", "Current Community_Master has no total units; box score suggests 224."],
    ["SER", "Sereno", "RISE Sereno", "Box Score / Market / Rent Roll", 320, "Mapped", "Matches current Community_Master total units."],
    ["VIE", "Viera", "RISE Viera", "Box Score / Rent Roll", 166, "Mapped", "Current Community_Master has no total units; box score suggests 166."],
    ["STA", "St. Augustine", "RISE St. Augustine", "Box Score / Rent Roll", 272, "Needs Atlas decision", "Present in sources but not in the current Community_Master."],
  ],
  Floor_Plan_Aliases: [
    ["Community_ID", "Community_Name", "Source_Floor_Plan_Code", "Resolved_Floor_Plan_Name", "Notes"],
    ["DOR", "Doro", "A4", "The Armstrong", "694-sf split in the Doro box-score export; rolls into the dashboard plan name."],
    ["DOR", "Doro", "A4-L", "The Armstrong", "Leased/large variant of the same 694-sf split."],
    ["DOR", "Doro", "A5", "The Bennett", "Patio option variant of the 694-sf split."],
  ],
  Source_File_Log: [
    ["Source_File_Name", "Detected_Source_Type", "Detected_Period", "Report_Month", "Snapshot_Date", "Community_Scope", "Load_Cadence", "Atlas_Target", "Status", "Notes"],
    ["02 - RISE - Box Score-11.xlsx", "Box Score", "04/01/2026 - 04/30/2026", "2026-04-01", "2026-04-30", "Multi-community", "Monthly", "Monthly_Upload plus BoxScore detail tabs", "Ready with mapping", "Use each property sheet as one community."],
    ["2026_04_30_15_23_RISE_Baymeadows_Market.xlsx", "Market Survey", "Updated on 04/30/2026", "2026-04-01", "2026-04-30", "RISE Baymeadows plus comps", "Daily/Weekly/Monthly snapshot", "Market detail tabs; limited monthly rollup", "Ready with mapping", "Filename timestamp can drive Snapshot_Date."],
    ["2026_04_29_11_40_RISE_Sereno_Market.xlsx", "Market Survey", "Updated on 04/29/2026", "2026-04-01", "2026-04-29", "RISE Sereno plus comps", "Daily/Weekly/Monthly snapshot", "Market detail tabs; limited monthly rollup", "Ready with mapping", "Use subject property as Community_ID SER."],
    ["2026_04_29_21_07_RISE_at_Glen_Kernan_Park_Market.xlsx", "Market Survey", "Updated on 04/29/2026", "2026-04-01", "2026-04-29", "RISE at Glen Kernan Park plus comps", "Daily/Weekly/Monthly snapshot", "Market detail tabs; limited monthly rollup", "Ready with mapping", "Use subject property as Community_ID GKP."],
    ["03 - RISE - Rent Roll-4.pdf", "Rent Roll PDF", "Mar 2025", "2025-03-01", "", "Multi-community", "Monthly", "RentRoll detail and selected monthly rollups", "Parser needed", "PDF extraction needs validation because resident lines can wrap."],
    ["03 - RISE - Rent Roll-5.pdf", "Rent Roll PDF", "Mar 2026", "2026-03-01", "", "Multi-community", "Monthly", "RentRoll detail and selected monthly rollups", "Parser needed", "PDF extraction needs validation because split leases/resident lines can wrap."],
  ],
  Source_Field_Map: [
    ["Source_Type", "Source_Location", "Source_Field", "Target_Table", "Target_Field", "Transform_or_Rule", "Grain", "Confidence", "Notes"],
    ["Box Score", "Property sheet cell A4", "Report period", "Target_Monthly_Upload", "Report_Month", "Use first day of period month.", "Community-month", "High", "Example 04/01/2026 - 04/30/2026 -> 2026-04-01."],
    ["Box Score", "Property sheet cell A3 or sheet name", "Property name", "Target_Monthly_Upload", "Community_ID", "Map through Community_Aliases.", "Community-month", "High", "Do not load if alias is unmapped."],
    ["Box Score", "Availability total row", "Units", "Target_Monthly_Upload", "Total_Units", "Use total row Units or Community_Master when approved.", "Community-month", "High", "Also update Community_Master where current total units are blank."],
    ["Box Score", "Availability total row", "Occupied", "Target_Monthly_Upload", "Occupied_Units", "Use total row occupied units.", "Community-month", "High", ""],
    ["Box Score", "Availability hidden total", "Leased Units", "Target_Monthly_Upload", "Leased_Units", "Use Availability: Leased Units, not the visible Leased percentage.", "Community-month", "High", "In sample, Baymeadows = 259."],
    ["Box Score", "Availability total row", "Leased Units - Occupied", "Target_Monthly_Upload", "Preleased_Units", "Calculate MAX(Leased_Units - Occupied_Units, 0).", "Community-month", "Medium", "Confirm Atlas definition of preleased."],
    ["Box Score", "Availability total row", "Vacant", "Target_Monthly_Upload", "Vacant_Units", "Use total row vacant units.", "Community-month", "High", ""],
    ["Box Score", "Availability total row", "Occupied %", "Target_Monthly_Upload", "Physical_Occupancy_%", "Use visible Occupied percentage or Occupied_Units / Rentable_Units.", "Community-month", "High", "Store as decimal, not 95.6."],
    ["Box Score", "Availability total row", "Leased %", "Target_Monthly_Upload", "Leased_%", "Use visible Leased percentage.", "Community-month", "High", "Store as decimal."],
    ["Box Score", "Availability total row", "Avg. Market Rent (Budgeted)", "Target_Monthly_Upload", "Avg_Market_Rent", "Use total row weighted average.", "Community-month", "High", ""],
    ["Box Score", "Availability total row", "Avg. Scheduled Rent", "Target_Monthly_Upload", "Actual_Rent", "Use only if Atlas defines Actual_Rent as average in-place/scheduled rent.", "Community-month", "Review", "If Atlas expects collected rent, leave blank and use accounting source."],
    ["Box Score", "Availability hidden total", "Total Budgeted Rent", "Target_Monthly_Upload", "Gross_Potential_Rent_$", "Use total budgeted rent if Atlas wants GPR.", "Community-month", "Medium", "Confirm whether budgeted or market rent should be GPR."],
    ["Box Score", "Availability hidden total", "Total Effective Rent", "Target_Monthly_Upload", "Concessions_$", "Do not map directly. Calculate concessions as Total Scheduled Rent - Total Effective Rent if approved.", "Community-month", "Review", "Current market files have concession fields too."],
    ["Box Score", "Property Pulse total row", "Move-Ins", "Target_Monthly_Upload", "Move_Ins", "Use total row move-ins.", "Community-month", "High", ""],
    ["Box Score", "Property Pulse total row", "Move-Outs", "Target_Monthly_Upload", "Move_Outs", "Use total row move-outs.", "Community-month", "High", ""],
    ["Box Score", "Property Pulse total row", "Renewal Leases Approved", "Target_Monthly_Upload", "Renewals_Signed", "Use as renewals signed/approved if Atlas accepts approved renewals as signed.", "Community-month", "Medium", "If renewal-sent counts exist elsewhere, use that source for Renewals_Sent."],
    ["Box Score", "Lead Activity total row", "New Leads", "Target_Monthly_Upload", "Leads", "Use total row new leads.", "Community-month", "High", ""],
    ["Box Score", "Lead Activity total row", "First Visits/Tours", "Target_Monthly_Upload", "Tours", "Use total row first visits/tours.", "Community-month", "High", ""],
    ["Box Score", "Lead Activity total row", "New Leads or First Visits/Tours", "Target_Monthly_Upload", "Traffic", "Review Atlas definition. Suggested default: Traffic = First Visits/Tours; alternate = New Leads.", "Community-month", "Review", "Avoid double-counting."],
    ["Box Score", "Lead Conversions total row", "Completed + Partially Completed", "Target_Monthly_Upload", "Applications", "Use completed application count; include partials only if Atlas wants in-progress applications.", "Community-month", "Medium", "Column labels repeat; section context required."],
    ["Box Score", "Lead Conversions total row", "Approved", "Target_Monthly_Upload", "Approved_Applications", "Use total row approved applications.", "Community-month", "High", ""],
    ["Market Survey", "Property Summary or Visual Market Survey", "Advertised Occupancy / Leased % / Exposure %", "Market_PropertySnapshot", "Occupancy and exposure fields", "Store as market snapshot. Do not overwrite official monthly ops unless selected as pricing source.", "Property-snapshot", "High", ""],
    ["Market Survey", "Visual Market Survey", "Concessions / Concession %", "Market_PropertySnapshot", "Concessions / Concession_Pct", "Store snapshot values. Optional monthly rollup by subject property and month-end snapshot.", "Property-snapshot", "High", ""],
    ["Market Survey", "Floor Plan Data", "Available Asking Rent / Leased Asking Rent / Days on Market", "Market_FloorPlanSnapshot", "Rent and availability fields", "Load one row per snapshot + property + floor plan.", "Floor-plan snapshot", "High", ""],
    ["Box Score", "Floor plan code variants", "A4 / A4-L / A5", "Community floor-plan setup", "Floor_Plan_Aliases", "Resolve code variants before updating the dashboard; aggregate by resolved floor-plan name.", "Floor-plan snapshot", "High", "Doro uses A4 and A4-L for The Armstrong and A5 for The Bennett."],
    ["Market Survey", "Unit Level Data", "Unit Status / Asking Rent / Eff Rent / Concession Details", "Market_UnitSnapshot", "Unit listing fields", "Load one row per snapshot + property + unit.", "Unit snapshot", "High", "High-volume table; keep separate from monthly summary."],
    ["Rent Roll PDF", "Unit Details", "Bldg-Unit / Unit Type / SQFT / Unit Status", "RentRoll_UnitMonthly", "Unit identity and status fields", "Parse by property and report month. Validate wrapped rows.", "Unit-month", "Medium", "PDF parsing requires exception review."],
    ["Rent Roll PDF", "Unit Details", "Market Rent (Budgeted) / Actual Charges / In-Place Rents", "RentRoll_UnitMonthly", "Rent fields", "Store unit-level fields; aggregate as approved into monthly Avg_Market_Rent or Actual_Rent.", "Unit-month", "Medium", ""],
    ["Rent Roll PDF", "Unit Details", "Balance", "Target_Monthly_Upload", "Delinquency_$", "Only aggregate positive balances if Atlas defines delinquency this way.", "Community-month", "Review", "Negative balances may be credits; do not net without business approval."],
    ["Financial source not provided", "GL / Accounting export", "Collected rent / other income / operating expenses / NOI", "Target_Monthly_Upload", "Collected_Rent_$ / Other_Income_$ / Operating_Expenses_$ / NOI_$", "Leave blank until financial export is provided.", "Community-month", "N/A", "These are not reliably available in the sample Box Score, Market, or Rent Roll files."],
    ["Maintenance source not provided", "Work order export", "Open / completed work orders", "Target_Monthly_Upload", "Open_Work_Orders / Completed_Work_Orders", "Leave blank until maintenance export is provided.", "Community-month", "N/A", ""],
    ["Reputation source not provided", "Reviews export", "Review count / average rating", "Target_Monthly_Upload", "Reviews_Count / Avg_Review_Rating", "Market Survey has ratings in some places but not true monthly review count.", "Community-month", "Low", "Use a reputation source if available."],
  ],
  Target_Monthly_Upload: [
    atlasMonthlyFields,
    ["2026-04-01", "BAY", "Baymeadows", "Florida", "Jacksonville", "Lease-Up", "BAYMEADOWS", 331, 217, 259, 42, 114, 0.65558912386707, 0.78247734138973, "", 1775.1873111782, 1721.2304147465, "", "", 587587, "", "", "", "", "", 254, 84, 56, 36, 23, 6, "", 10, "", "", "", "", "Actual", "Excel Export", "02 - RISE - Box Score-11.xlsx", "Example row from April 2026 Box Score. Review Traffic/Actual_Rent/Renewals mapping before production.", "2026-04-30", "REVIEW", "Traffic, Actual_Rent, and Renewals_Signed require Atlas definition confirmation."],
  ],
  BoxScore_FloorPlanMonthly: [
    ["Report_Month", "Community_ID", "Unit_Type", "Avg_SQFT", "Avg_Market_Rent", "Avg_Scheduled_Rent", "Avg_Net_Effective_Rent", "Units", "Excluded", "Rentable_Units", "Occupied_Units", "Vacant_Units", "Available_Units", "Occupied_No_Notice", "Notice_Rented", "Notice_Unrented", "Vacant_Rented", "Vacant_Unrented", "Occupancy_Pct", "Leased_Pct", "Trend", "Exposure_Pct", "Total_Square_Feet", "Total_Market_Rent", "Total_Scheduled_Charges", "Total_Budgeted_Rent", "Total_Scheduled_Rent", "Total_Scheduled_Other", "Total_Effective_Rent", "Total_Actual_Rent", "Vacant_Rented_Ready_Units", "Vacant_Rented_Not_Ready_Units", "Vacant_Unrented_Ready_Units", "Vacant_Unrented_Not_Ready_Units", "Leased_Units", "Exposure_Units", "Not_Exposed_Leased_Units", "Source_File_Name", "Source_Sheet"],
  ],
  BoxScore_PulseMonthly: [
    ["Report_Month", "Community_ID", "Unit_Type", "Units", "Move_Ins", "MTM", "MTM_Conversions", "Renewal_Leases_Approved", "Transfers", "Notices", "Move_Outs", "Renewal_Transfers", "Skips", "Evictions", "Leased", "Source_File_Name", "Source_Sheet"],
  ],
  BoxScore_LeadMonthly: [
    ["Report_Month", "Community_ID", "Unit_Type", "New_Leads", "Email_Leads", "Call_Leads", "Online_Leads", "Walk_In_Leads", "Off_Site_Event_Leads", "Chat_Leads", "Text_Leads", "Other_Leads", "Emails", "Calls", "Chats", "Texts", "First_Visits_Tours", "Source_File_Name", "Source_Sheet"],
  ],
  BoxScore_ApplicationMonthly: [
    ["Report_Month", "Community_ID", "Unit_Type", "Applications_Completed", "Applications_Partially_Completed", "Applications_Completed_Cancelled", "Applications_Denied", "Applications_Approved", "Applications_Approved_Cancelled", "Lease_Conversions_Completed", "Lease_Conversions_Cancelled", "Lease_Conversions_Approved", "Source_File_Name", "Source_Sheet"],
  ],
  BoxScore_MakeReadyMonthly: [
    ["Report_Month", "Community_ID", "Status", "Vacant_Rented", "Vacant_Unrented", "Total_Vacant", "Pct_Of_Total_Vacant", "Source_File_Name", "Source_Sheet"],
  ],
  Market_PropertySnapshot: [
    ["Snapshot_Date", "Report_Month", "Subject_Community_ID", "Source_Property_Name", "Is_Subject_Property", "Distance", "Year_Built", "Total_Units", "Exposure_Pct", "Leased_Pct", "Preleased_Pct", "Applications_Last_7_Days", "Applications_Last_30_Days", "Concessions", "Concession_Pct", "Total_Available_Units", "Total_Leased_Units", "Vacant_Units", "Asking_Rent", "Asking_Rent_PSF", "Effective_Rent", "Effective_Rent_PSF", "Avg_SqFt", "Source_File_Name"],
  ],
  Market_FloorPlanSnapshot: [
    ["Snapshot_Date", "Report_Month", "Subject_Community_ID", "Source_Property_Name", "Floor_Plan", "Bed_Count", "Bath_Count", "Sq_Ft", "Total_Units", "Unit_Mix", "Leased_Units_Last_7_Days", "Leased_Days_On_Market", "Leased_Asking_Rent", "Leased_Asking_Rent_PSF", "Available_Units", "Available_Days_On_Market", "Available_Asking_Rent", "Available_Asking_Rent_PSF", "Last_7_Days_Asking_Rent_Trend", "Source_File_Name"],
  ],
  Market_UnitSnapshot: [
    ["Snapshot_Date", "Report_Month", "Subject_Community_ID", "Source_Property_Name", "Address", "Floor_Plan", "Unit_Number", "Bed_Count", "Bath_Count", "Sq_Ft", "Unit_Status", "Date_First_Available", "Leased_Date", "Days_On_Market", "Times_Listed", "Asking_Rent", "Asking_Rent_PSF", "Effective_Rent", "Effective_Rent_PSF", "Concession_Pct", "Concession_Details", "Source_File_Name"],
  ],
  RentRoll_UnitMonthly: [
    ["Report_Month", "Community_ID", "Source_Property_Name", "Bldg_Unit", "Unit_Type", "SqFt", "Unit_Status", "Resident_Name", "Move_In", "Lease_Start", "Lease_End", "Expected_Move_Out", "Market_Rent_Budgeted", "Actual_Charges", "In_Place_Rent_Scheduled", "Balance", "Deposit_Held", "Ready_Status", "PII_Load_Flag", "Source_File_Name", "Source_Page"],
  ],
  QA_Rules: [
    ["Rule_ID", "Target_Table", "Severity", "Check", "Expected_Result", "Action"],
    ["QA-001", "Target_Monthly_Upload", "Error", "Report_Month and Community_ID are present", "No blanks on populated rows", "Reject row until fixed."],
    ["QA-002", "Target_Monthly_Upload", "Error", "Community_ID exists in Community_Aliases / Community_Master", "All IDs valid", "Add alias or add community to master."],
    ["QA-003", "Target_Monthly_Upload", "Error", "Occupied_Units <= Total_Units and Leased_Units <= Total_Units", "True", "Flag source or mapping issue."],
    ["QA-004", "Target_Monthly_Upload", "Review", "Preleased_Units = MAX(Leased_Units - Occupied_Units, 0)", "Formula reconciles", "Confirm Atlas preleased definition."],
    ["QA-005", "BoxScore detail", "Error", "Floor-plan totals sum to community totals", "Total rows reconcile to detail", "Flag if source table parse skipped a row."],
    ["QA-006", "Market snapshots", "Review", "Snapshot_Date assigned from workbook updated date or filename timestamp", "Date present", "Do not aggregate weekly/daily snapshots into monthly values unless business rule is selected."],
    ["QA-007", "RentRoll_UnitMonthly", "Error", "PDF parser confidence reviewed for wrapped resident/lease lines", "No unparsed unit rows", "Manual review needed for exceptions."],
    ["QA-008", "Sensitive fields", "Error", "Resident_Name, Balance, Deposit_Held loaded only when PII_Load_Flag = Yes", "No accidental PII load", "Remove or mask before non-restricted load."],
    ["QA-009", "Financial fields", "Review", "Collected_Rent_$, Other_Income_$, Operating_Expenses_$, NOI_$ not inferred from rent roll or box score", "Blank unless financial source provided", "Use GL/accounting export."],
  ],
  Atlas_Teaching_Notes: [
    ["Question", "Answer"],
    ["Is the current Atlas template appropriate?", "Yes for monthly community-level dashboard facts. It is not enough by itself for all detail in your examples."],
    ["Can Atlas read the raw files exactly as uploaded?", "Only if Atlas has a configurable import/parser layer, Power Query, API, or an agent workflow using this mapping. Otherwise the raw files must be transformed first."],
    ["What should be loaded to Atlas monthly?", "Use Target_Monthly_Upload for official month-end values by community."],
    ["Where should floor-plan, comp, and unit-level details go?", "Use the detail fact tabs. Add these tables to Atlas/Power BI if you want drill-down beyond the monthly community row."],
    ["How should Doro floor-plan codes be resolved?", "Use Floor_Plan_Aliases: A4 and A4-L map to The Armstrong, and A5 maps to The Bennett."],
    ["How should daily/weekly market files map to months?", "Keep Snapshot_Date for the actual file date and Report_Month as the first day of that month. For official monthly reporting, use the month-end/latest approved snapshot."],
    ["How should multi-community bulk upload work?", "Upload all source files, map aliases to Community_ID, create one monthly row per community, then append optional detail rows to their own tables."],
    ["What is the safest production rule?", "Do not let unknown aliases, unconfirmed financial fields, or PII fields auto-load. Route them to REVIEW."],
  ],
};

function colLabel(n) {
  let label = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    label = String.fromCharCode(65 + rem) + label;
    n = Math.floor((n - 1) / 26);
  }
  return label;
}

function normalizeRows(rows) {
  const width = Math.max(...rows.map((row) => row.length));
  return rows.map((row) => {
    const out = row.slice();
    while (out.length < width) out.push("");
    return out;
  });
}

function writeSheet(workbook, name, rows, options = {}) {
  const sheet = workbook.worksheets.add(name);
  const values = normalizeRows(rows);
  const lastCol = colLabel(values[0].length);
  const lastRow = values.length;
  const range = sheet.getRange(`A1:${lastCol}${lastRow}`);
  range.values = values;
  range.format = {
    font: { name: "Aptos", size: 10, color: "#111827" },
    verticalAlignment: "top",
    wrapText: true,
  };
  const header = sheet.getRange(`A1:${lastCol}1`);
  header.format = {
    fill: "#1F4E78",
    font: { name: "Aptos", size: 10, color: "#FFFFFF", bold: true },
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "outside", style: "thin", color: "#D1D5DB" },
  };
  range.format.borders = { preset: "inside", style: "thin", color: "#E5E7EB" };
  sheet.freezePanes.freezeRows(1);
  if (options.freezeCols) sheet.freezePanes.freezeColumns(options.freezeCols);
  range.format.autofitColumns();
  range.format.autofitRows();
  return sheet;
}

await fs.mkdir(outputDir, { recursive: true });
const workbook = Workbook.create();

for (const [name, rows] of Object.entries(sheets)) {
  const freezeCols = ["Source_Field_Map", "Community_Aliases", "Floor_Plan_Aliases"].includes(name) ? 2 : 0;
  writeSheet(workbook, name, rows, { freezeCols });
}

const inspect = await workbook.inspect({
  kind: "table",
  range: "Source_Field_Map!A1:I12",
  include: "values",
  tableMaxRows: 12,
  tableMaxCols: 9,
});
console.log(inspect.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "formula error scan",
});
console.log(errors.ndjson);

await workbook.render({ sheetName: "README", range: "A1:B7", scale: 2 });
await workbook.render({ sheetName: "Source_Field_Map", range: "A1:I18", scale: 2 });
await workbook.render({ sheetName: "Floor_Plan_Aliases", range: "A1:E6", scale: 2 });
await workbook.render({ sheetName: "Target_Monthly_Upload", range: "A1:AR2", scale: 1 });

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(outputPath);
