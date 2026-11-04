function updateMileageSystem() {
  importMissionAreas();
  updateMileageFormAreas();
}

function importMissionAreas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Create or clear the Areas sheet
  let sheet = ss.getSheetByName('Areas');

  if (!sheet) {
    sheet = ss.insertSheet('Areas');
  } else {
    sheet.clearContents();
  }

  const headers = [
    'Area',
    'District',
    'Zone',
    'Ecclesiastical Unit',
    'Vehicle',
    'Vehicle Allowance/Mo',
    'VIN Last 8',
    'Vehicle CCID',
    'Email'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const rows = [];
  let pageToken = null;

  do {
    const response = People.People.Connections.list('people/me', {
      pageSize: 100,
      pageToken: pageToken,
      personFields:
        'names,emailAddresses,phoneNumbers,organizations,biographies'
    });

    const people = response.connections || [];

    people.forEach(person => {
      const notes = person.biographies?.[0]?.value || '';

      // Only include contacts that look like mission area contacts
      if (!notes.includes('Area:') || !notes.includes('Vehicle:')) {
        return;
      }

      const email =
        person.emailAddresses?.[0]?.value || '';

      rows.push([
        extractField(notes, 'Area'),
        extractField(notes, 'District'),
        extractField(notes, 'Zone'),
        extractField(notes, 'Ecclesiastical Unit'),
        extractField(notes, 'Vehicle'),
        extractField(notes, 'Vehicle Allowance/Mo'),
        extractField(notes, 'Vehicle VIN Last 8'),
        extractField(notes, 'Vehicle CCID'),
        email
      ]);
    });

    pageToken = response.nextPageToken;

  } while (pageToken);

  // Write all areas at once
  if (rows.length > 0) {
    sheet
      .getRange(2, 1, rows.length, headers.length)
      .setValues(rows);
  }

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);

  Logger.log(`Imported ${rows.length} mission areas.`);
}


/**
 * Extracts a value from a Notes field.
 *
 * Example:
 *   Area: H' Rockwall YSA 1
 *
 * returns:
 *   H' Rockwall YSA 1
 */
function extractField(notes, fieldName) {
  const lines = notes.split(/\r?\n/);

  for (const line of lines) {
    const prefix = fieldName + ':';

    if (line.trim().startsWith(prefix)) {
      return line
        .trim()
        .substring(prefix.length)
        .trim();
    }
  }

  return '';
}

function updateMileageFormAreas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Areas');

  if (!sheet) {
    throw new Error('Could not find the "Areas" sheet.');
  }

  const formUrl = ss.getFormUrl();

  if (!formUrl) {
    throw new Error(
      'This spreadsheet is not connected to a Google Form.'
    );
  }

  Logger.log('Form URL: ' + formUrl);

  const form = FormApp.openByUrl(formUrl);

  Logger.log('Form title: ' + form.getTitle());

  // Get area names from the Areas sheet.
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    throw new Error('The Areas sheet contains no areas.');
  }

  const areas = sheet
    .getRange(2, 1, lastRow - 1, 1)
    .getValues()
    .flat()
    .map(value => String(value).trim())
    .filter(value => value !== '');

  const uniqueAreas = [...new Set(areas)]
  .sort((a, b) => a.localeCompare(b));

  Logger.log('Found ' + uniqueAreas.length + ' areas.');
  Logger.log('First area: ' + uniqueAreas[0]);

  // Show every question in the form and its type.
  const items = form.getItems();

  Logger.log('----- FORM QUESTIONS -----');

  items.forEach((item, index) => {
    Logger.log(
      index +
      ': "' +
      item.getTitle() +
      '" | Type: ' +
      item.getType()
    );
  });

  Logger.log('--------------------------');

  // Look for an Area question.
  let areaItem = null;

  for (const item of items) {
    if (item.getTitle().trim().toLowerCase() === 'area') {

      Logger.log(
        'Found Area question. Type: ' + item.getType()
      );

      if (item.getType() === FormApp.ItemType.LIST) {
        areaItem = item.asListItem();
        break;
      }

      if (item.getType() === FormApp.ItemType.MULTIPLE_CHOICE) {
        areaItem = item.asMultipleChoiceItem();
        break;
      }

      throw new Error(
        'The "Area" question exists, but its type is ' +
        item.getType() +
        '. Make it a Dropdown or Multiple choice question.'
      );
    }
  }

  if (!areaItem) {
    throw new Error(
      'Could not find a question named "Area".'
    );
  }

  // Set the choices.
  areaItem.setChoiceValues(uniqueAreas);

  Logger.log(
    'SUCCESS: Updated Area question with ' +
    uniqueAreas.length +
    ' choices.'
  );
}

function addAreaInformationToResponse(e) {
  const responseSheet = e.range.getSheet();
  const row = e.range.getRow();

  const areasSheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Areas');

  if (!areasSheet) {
    throw new Error('Could not find the "Areas" sheet.');
  }

  // Get headers from the response sheet
  const responseHeaders = responseSheet
    .getRange(1, 1, 1, responseSheet.getLastColumn())
    .getValues()[0];

  // Find the Area column
  const areaColumn = responseHeaders.indexOf('Area') + 1;

  if (areaColumn === 0) {
    throw new Error('Could not find an "Area" column in the response sheet.');
  }

  const area = responseSheet
    .getRange(row, areaColumn)
    .getValue()
    .toString()
    .trim();

  if (!area) {
    return;
  }

  // Get Areas sheet data
  const areasData = areasSheet
    .getRange(1, 1, areasSheet.getLastRow(), areasSheet.getLastColumn())
    .getValues();

  const areasHeaders = areasData[0];

  const areasAreaColumn =
    areasHeaders.indexOf('Area');

  const zoneColumn =
    areasHeaders.indexOf('Zone');

  const districtColumn =
    areasHeaders.indexOf('District');

  if (areasAreaColumn === -1 || zoneColumn === -1) {
    throw new Error(
      'The Areas sheet must contain "Area" and "Zone" columns.'
    );
  }

  // Find the matching area
  let matchingRow = null;

  for (let i = 1; i < areasData.length; i++) {
    if (
      String(areasData[i][areasAreaColumn]).trim() === area
    ) {
      matchingRow = areasData[i];
      break;
    }
  }

  if (!matchingRow) {
    Logger.log('No matching area found for: ' + area);
    return;
  }

  const zone = matchingRow[zoneColumn];

  // Find the Zone column in the response sheet
  const responseZoneColumn =
    responseHeaders.indexOf('Zone') + 1;

  if (responseZoneColumn === 0) {
    throw new Error(
      'Could not find a "Zone" column in the response sheet.'
    );
  }

  // Write the correct Zone
  responseSheet
    .getRange(row, responseZoneColumn)
    .setValue(zone);

  // If there is a District column, populate it too
  if (districtColumn !== -1) {
    const district = matchingRow[districtColumn];

    let responseDistrictColumn =
      responseHeaders.indexOf('District') + 1;

    if (responseDistrictColumn === 0) {
      responseDistrictColumn =
        responseSheet.getLastColumn() + 1;

      responseSheet
        .getRange(1, responseDistrictColumn)
        .setValue('District');
    }

    responseSheet
      .getRange(row, responseDistrictColumn)
      .setValue(district);
  }

  Logger.log(
    `Updated ${area}: Zone=${zone}`
  );
}

function updateMonthlyMileageSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const responseSheet = ss.getSheets().find(sheet =>
    sheet.getName().toLowerCase().includes('form responses')
  );

  if (!responseSheet) {
    throw new Error('Could not find the Form Responses sheet.');
  }

  const areasSheet = ss.getSheetByName('Areas');

  if (!areasSheet) {
    throw new Error('Could not find the "Areas" sheet.');
  }

  // ------------------------------------------------------------
  // Read Form Responses
  // ------------------------------------------------------------

  const responseData = responseSheet.getDataRange().getValues();

  if (responseData.length < 2) {
    Logger.log('No form responses found.');
    return;
  }

  const responseHeaders = responseData[0];

  const dateColumn = findColumn(responseHeaders, 'Date of event');
  const areaColumn = findColumn(responseHeaders, 'Area');
  const categoryColumn = findColumn(responseHeaders, 'Mileage Category');
  const milesColumn = findColumn(responseHeaders, 'Number of miles used (Please enter just the number!)');

  if (
    dateColumn === -1 ||
    areaColumn === -1 ||
    categoryColumn === -1 ||
    milesColumn === -1
  ) {
    throw new Error(
      'Could not find one or more required Form Response columns.'
    );
  }

  // ------------------------------------------------------------
  // Read Areas sheet
  // ------------------------------------------------------------

  const areasData = areasSheet.getDataRange().getValues();

  if (areasData.length < 2) {
    throw new Error('The Areas sheet contains no areas.');
  }

  const areaHeaders = areasData[0];

  const areasAreaColumn =
    findColumn(areaHeaders, 'Area');

  const districtColumn =
    findColumn(areaHeaders, 'District');

  const zoneColumn =
    findColumn(areaHeaders, 'Zone');

  const vehicleColumn =
    findColumn(areaHeaders, 'Vehicle');

  const allowanceColumn =
    findColumn(areaHeaders, 'Vehicle Allowance/Mo');

  if (areasAreaColumn === -1 || zoneColumn === -1) {
    throw new Error(
      'The Areas sheet must contain Area and Zone columns.'
    );
  }

  // Create a lookup table for areas.
  const areaLookup = {};

  for (let i = 1; i < areasData.length; i++) {
    const row = areasData[i];

    const area = String(row[areasAreaColumn]).trim();

    if (!area) {
      continue;
    }

    areaLookup[area] = {
      district:
        districtColumn >= 0
          ? row[districtColumn]
          : '',

      zone:
        zoneColumn >= 0
          ? row[zoneColumn]
          : '',

      vehicle:
        vehicleColumn >= 0
          ? row[vehicleColumn]
          : '',

      allowance:
        allowanceColumn >= 0
          ? row[allowanceColumn]
          : ''
    };
  }

  // ------------------------------------------------------------
  // Build monthly/area/category totals
  // ------------------------------------------------------------

  const summary = {};

  for (let i = 1; i < responseData.length; i++) {
    const row = responseData[i];

    const date = row[dateColumn];
    const area = String(row[areaColumn]).trim();
    const category = String(row[categoryColumn]).trim();
    const miles = Number(row[milesColumn]) || 0;

    if (!date || !area || !category || miles === 0) {
      continue;
    }

    // Convert date to YYYY-MM using the shared date helper.
    const month = normalizeMileageMonth(date);

    if (!month) {
      Logger.log(
        `Skipping row ${i + 1}: Invalid date value: ${date}`
      );
    continue;
  }

Logger.log(
  `Processing row ${i + 1}: Date=${date}, Month=${month}, Area=${area}, Miles=${miles}`
);

    const key = month + '|' + area;

    if (!summary[key]) {
      const areaInfo = areaLookup[area] || {};

      summary[key] = {
        month: month,
        area: area,
        district: areaInfo.district || '',
        zone: areaInfo.zone || '',
        vehicle: areaInfo.vehicle || '',
        allowance: areaInfo.allowance || '',
        totalMiles: 0,
        categories: {}
      };
    }

    summary[key].totalMiles += miles;

    if (!summary[key].categories[category]) {
      summary[key].categories[category] = 0;
    }

    summary[key].categories[category] += miles;
  }

  // ------------------------------------------------------------
  // Get categories currently used in the Form
  // ------------------------------------------------------------

  const formUrl = ss.getFormUrl();

  if (!formUrl) {
    throw new Error(
      'This spreadsheet is not connected to a Google Form.'
    );
  }

  const form = FormApp.openByUrl(formUrl);

  const categoryQuestions = form
    .getItems(FormApp.ItemType.MULTIPLE_CHOICE)
    .map(item => item.asMultipleChoiceItem())
    .filter(item =>
      item.getTitle().trim() === 'Mileage Category'
    );

  const categories = [];

  if (categoryQuestions.length > 0) {
    const choices = categoryQuestions[0].getChoices();

    choices.forEach(choice => {
      const value = choice.getValue().trim();

      if (value && !categories.includes(value)) {
        categories.push(value);
      }
    });
  }

  // Also include categories that appear in historical responses.
  Object.values(summary).forEach(record => {
    Object.keys(record.categories).forEach(category => {
      if (!categories.includes(category)) {
        categories.push(category);
      }
    });
  });

  // ------------------------------------------------------------
  // Create / clear Monthly Summary sheet
  // ------------------------------------------------------------

  let summarySheet = ss.getSheetByName('Monthly Summary');

  if (!summarySheet) {
    summarySheet = ss.insertSheet('Monthly Summary');
  } else {
    summarySheet.clearContents();
  }

  // Header
  const headers = [
    'Month',
    'Area',
    'District',
    'Zone',
    'Vehicle',
    'Monthly Allowance',
    'Total Exception Miles',
    ...categories
  ];

  summarySheet
    .getRange(1, 1, 1, headers.length)
    .setValues([headers]);

  // ------------------------------------------------------------
  // Create output rows
  // ------------------------------------------------------------

  const output = Object.values(summary)
    .sort((a, b) => {
      if (a.month !== b.month) {
        return b.month.localeCompare(a.month);
      }

      return a.area.localeCompare(b.area);
    })
    .map(record => {
      const categoryMiles = categories.map(
        category => record.categories[category] || 0
      );

      return [
        record.month,
        record.area,
        record.district,
        record.zone,
        record.vehicle,
        record.allowance,
        record.totalMiles,
        ...categoryMiles
      ];
    });

  if (output.length > 0) {
    summarySheet
      .getRange(
        2,
        1,
        output.length,
        headers.length
      )
      .setValues(output);
  }

  // Format the header
const headerRange = summarySheet.getRange(
  1,
  1,
  1,
  headers.length
);

headerRange
  .setFontWeight('bold')
  .setHorizontalAlignment('center');

// Freeze the header row
summarySheet.setFrozenRows(1);

// Add a filter to the entire table
if (summarySheet.getFilter()) {
  summarySheet.getFilter().remove();
}

if (output.length > 0) {
  summarySheet
    .getRange(
      1,
      1,
      output.length + 1,
      headers.length
    )
    .createFilter();
}

// Format mileage columns as whole numbers
if (output.length > 0) {
  summarySheet
    .getRange(
      2,
      7,
      output.length,
      headers.length - 6
    )
    .setNumberFormat('0');
}

// Make the columns readable
summarySheet.autoResizeColumns(
  1,
  headers.length
);

// Give the Area column a little extra room
summarySheet.setColumnWidth(2, 180);

syncMonthlyMileageSheets();
updateMileageDashboards();

Logger.log(
  `Monthly Summary updated: ${output.length} area/month records.`
);
}


/**
 * Finds a column by its header name.
 * Returns the zero-based column index.
 */
function findColumn(headers, name) {
  const target = name.trim().toLowerCase();

  return headers.findIndex(header =>
    String(header).trim().toLowerCase() === target
  );
}

function updateMissionData() {
  importMissionAreas();
  updateMileageFormAreas();

  Logger.log('Mission data update complete.');
}


function processMileageSubmission(e) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    addAreaInformationToResponse(e);
    updateSingleAreaMonth(e);

    Logger.log('Mileage submission processed successfully.');

  } finally {
    lock.releaseLock();
  }
}


function dailyMissionDataUpdate() {
  updateMissionData();

  // Full reconciliation: rebuilds the master summary,
  // monthly tabs, dashboards, charts, and tab order.
  updateMonthlyMileageSummary();

  Logger.log('Daily full mileage reconciliation complete.');
}


function syncMonthlyMileageSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = ss.getSheetByName('Monthly Summary');

  if (!summarySheet) {
    throw new Error('Could not find the "Monthly Summary" sheet.');
  }

  const data = summarySheet.getDataRange().getValues();

  if (data.length === 0 || data[0].length === 0) {
    throw new Error('The Monthly Summary sheet is empty.');
  }

  const headers = data[0];
  const monthlyData = {};

  // Group the master summary rows by month.
  // Group the master summary rows by month.
for (let i = 1; i < data.length; i++) {
  const row = data[i];
  const value = row[0];

  let month = '';

  if (value instanceof Date && !isNaN(value.getTime())) {
    month = Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      'yyyy-MM'
    );
  } else {
    const textValue = String(value).trim();

    if (/^\d{4}-\d{2}$/.test(textValue)) {
      month = textValue;
    } else if (/^\d{4}-\d{2}-\d{2}/.test(textValue)) {
      month = textValue.substring(0, 7);
    }
  }

  if (!month) {
    Logger.log(
      `Skipping row ${i + 1}: unrecognized month value "${value}"`
    );
    continue;
  }

  if (!monthlyData[month]) {
    monthlyData[month] = [];
  }

  monthlyData[month].push(row);
}

  // Ensure the current month's tab exists even before
  // the first mileage response of that month.
  const currentMonth = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyy-MM'
  );

  if (!monthlyData[currentMonth]) {
    monthlyData[currentMonth] = [];
  }

  // Refresh every month represented in the master summary.
  Object.keys(monthlyData).sort().forEach(month => {
    const parts = month.split('-');
    const year = Number(parts[0]);
    const monthNumber = Number(parts[1]);

    const monthName = Utilities.formatDate(
      new Date(year, monthNumber - 1, 1),
      Session.getScriptTimeZone(),
      'MMMM yyyy'
    );

    let sheet = ss.getSheetByName(monthName);

    if (!sheet) {
      sheet = ss.insertSheet(monthName);
    }

    // Remove the old filter before rebuilding the snapshot.
    if (sheet.getFilter()) {
      sheet.getFilter().remove();
    }

    // Clear old content and formatting so the refreshed
    // snapshot doesn't retain stale formatting or values.
    sheet.clear();

    const rows = monthlyData[month];
    const output = [headers, ...rows];

    // Write the header and this month's records.
    sheet.getRange(
      1,
      1,
      output.length,
      headers.length
    ).setValues(output);

    // Match the master summary's column widths.
    for (let col = 1; col <= headers.length; col++) {
      sheet.setColumnWidth(
        col,
        summarySheet.getColumnWidth(col)
      );
    }

    // Format the header.
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');

    sheet.setFrozenRows(1);

    // Format Month column.
if (rows.length > 0) {
  sheet.getRange(2, 1, rows.length, 1)
    .setNumberFormat('mmmm yyyy');
}

// Format mileage totals and category columns.
if (rows.length > 0 && headers.length >= 7) {
  sheet.getRange(
    2,
    7,
    rows.length,
    headers.length - 6
  ).setNumberFormat('0');
}

    // Make the monthly tab filterable.
    sheet.getRange(
      1,
      1,
      output.length,
      headers.length
    ).createFilter();

    // Keep the Area column readable.
    if (headers.length >= 2) {
      sheet.setColumnWidth(2, 180);
    }

    Logger.log(
      `Updated ${monthName}: ${rows.length} area records.`
    );
  });

  Logger.log('All monthly mileage tabs are up to date.');

  organizeMileageSheets();
}


function updateMileageDashboards() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = ss.getSheetByName('Monthly Summary');

  if (!source) {
    throw new Error('Could not find Monthly Summary.');
  }

  const data = source.getDataRange().getValues();

  if (data.length === 0 || data[0].length < 7) {
    throw new Error('Monthly Summary is missing required columns.');
  }

  const headers = data[0];
  const categories = headers.slice(7);
  const records = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const month = normalizeMileageMonth(row[0]);
    const area = String(row[1] || '').trim();

    if (!month || !area) continue;

    records.push({
      month,
      area,
      total: Number(row[6]) || 0,
      categoryMiles: categories.map(
        (_, index) => Number(row[index + 7]) || 0
      )
    });
  }

  updateMonthlyOverview(ss, records, categories);
  updateMissionOverview(ss, records, categories);
  addMissionOverviewCharts(
    ss.getSheetByName('Mission Overview')
  );

  Logger.log('Monthly Overview and Mission Overview updated.');
}


function updateMonthlyOverview(ss, records, categories) {
  const totals = {};

  records.forEach(record => {
    if (!totals[record.month]) {
      totals[record.month] = {
        total: 0,
        categories: categories.map(() => 0)
      };
    }

    totals[record.month].total += record.total;

    record.categoryMiles.forEach((miles, i) => {
      totals[record.month].categories[i] += miles;
    });
  });

  const sheet = getOrCreateMileageSheet(ss, 'Monthly Overview');

  const headers = [
    'Month',
    'Total Exception Miles',
    ...categories
  ];

  const months = Object.keys(totals).sort(
    (a, b) => b.localeCompare(a)
  );

  const output = months.map(month => [
    month,
    totals[month].total,
    ...totals[month].categories
  ]);

  writeMileageTable(sheet, headers, output);

  Logger.log(`Monthly Overview: ${output.length} months.`);
}


function updateMissionOverview(ss, records, categories) {
  const totalMiles = records.reduce(
    (sum, record) => sum + record.total, 0
  );

  const months = [...new Set(records.map(r => r.month))].sort(
    (a, b) => b.localeCompare(a)
  );

  const areas = {};

  const categoryTotals = categories.map(() => 0);
  const monthlyTotals = {};

  records.forEach(record => {
    if (!areas[record.area]) {
      areas[record.area] = 0;
    }

    areas[record.area] += record.total;

    record.categoryMiles.forEach((miles, i) => {
      categoryTotals[i] += miles;
    });

    if (!monthlyTotals[record.month]) {
      monthlyTotals[record.month] = {
        total: 0,
        categories: categories.map(() => 0)
      };
    }

    monthlyTotals[record.month].total += record.total;

    record.categoryMiles.forEach((miles, i) => {
      monthlyTotals[record.month].categories[i] += miles;
    });
  });

  const sheet = getOrCreateMileageSheet(ss, 'Mission Overview');

  // Remove old filters and rebuild the dashboard.
  if (sheet.getFilter()) sheet.getFilter().remove();
  sheet.clear();

  let row = 1;

  sheet.getRange(row, 1)
    .setValue('MISSION-WIDE MILEAGE OVERVIEW')
    .setFontWeight('bold')
    .setFontSize(14);

  row += 2;

  // Top-level statistics.
  const stats = [
    ['All-Time Exception Miles', totalMiles],
    ['Areas with Recorded Exceptions', Object.keys(areas).length],
    ['Months with Recorded Exceptions', months.length]
  ];

  sheet.getRange(row, 1, stats.length, 2).setValues(stats);
  sheet.getRange(row, 1, stats.length, 1).setFontWeight('bold');
  sheet.getRange(row, 2, 1, 1).setNumberFormat('#,##0');

  row += stats.length + 2;

  // Category totals.
  row = writeDashboardSection(
    sheet,
    row,
    'ALL-TIME MILES BY CATEGORY',
    ['Mileage Category', 'Total Miles'],
    categories.map((category, i) => [
      category,
      categoryTotals[i]
    ]).sort((a, b) => b[1] - a[1])
  );

  // Area totals.
  row = writeDashboardSection(
    sheet,
    row,
    'ALL-TIME MILES BY AREA',
    ['Area', 'Total Exception Miles'],
    Object.entries(areas)
      .map(([area, miles]) => [area, miles])
      .sort((a, b) => b[1] - a[1])
  );

  // Month totals, newest first.
  row = writeDashboardSection(
    sheet,
    row,
    'ALL-TIME MONTHLY TOTALS',
    ['Month', 'Total Exception Miles', ...categories],
    months.map(month => [
      month,
      monthlyTotals[month].total,
      ...monthlyTotals[month].categories
    ])
  );

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, Math.max(2, categories.length + 2));
  sheet.setColumnWidth(1, 240);
  sheet.setColumnWidth(2, 170);

  Logger.log('Mission Overview refreshed.');
}


function writeDashboardSection(sheet, startRow, title, headers, rows) {
  sheet.getRange(startRow, 1)
    .setValue(title)
    .setFontWeight('bold');

  const headerRow = startRow + 1;

  sheet.getRange(
    headerRow, 1, 1, headers.length
  ).setValues([headers])
    .setFontWeight('bold');

  if (rows.length > 0) {
    sheet.getRange(
      headerRow + 1, 1, rows.length, headers.length
    ).setValues(rows);

    if (headers.length > 1) {
      sheet.getRange(
        headerRow + 1, 2, rows.length, headers.length - 1
      ).setNumberFormat('#,##0');
    }
  }

  return headerRow + Math.max(rows.length, 1) + 3;
}


function getOrCreateMileageSheet(ss, name) {
  let sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
  }

  return sheet;
}


function writeMileageTable(sheet, headers, rows) {
  if (sheet.getFilter()) sheet.getFilter().remove();
  sheet.clear();

  const values = [headers, ...rows];

  sheet.getRange(
    1, 1, values.length, headers.length
  ).setValues(values);

  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet.setFrozenRows(1);

  if (rows.length > 0) {
    sheet.getRange(
      2, 2, rows.length, headers.length - 1
    ).setNumberFormat('#,##0');
  }

  sheet.getRange(
    1, 1, values.length, headers.length
  ).createFilter();

  sheet.autoResizeColumns(1, headers.length);

  if (headers.length >= 2) {
    sheet.setColumnWidth(2, 170);
  }
}

function normalizeMileageMonth(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      'yyyy-MM'
    );
  }

  const textValue = String(value ?? '').trim();

  // Already YYYY-MM
  if (/^\d{4}-\d{2}$/.test(textValue)) {
    return textValue;
  }

  // YYYY-MM-DD or YYYY-MM-DD with a time
  if (/^\d{4}-\d{2}-\d{2}/.test(textValue)) {
    return textValue.substring(0, 7);
  }

  // MM/DD/YYYY
  const usDate = textValue.match(
    /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/
  );

  if (usDate) {
    return `${usDate[3]}-${usDate[1].padStart(2, '0')}`;
  }

  return '';
}


function addMissionOverviewCharts(sheet) {
  if (!sheet) {
    throw new Error('Mission Overview sheet was not found.');
  }

  // Remove old charts so refreshing the dashboard does not duplicate them.
  sheet.getCharts().forEach(chart => sheet.removeChart(chart));

  const categoryRange = findMissionOverviewChartData(
    sheet,
    'category'
  );

  const monthlyRange = findMissionOverviewChartData(
    sheet,
    'monthly'
  );

  // Chart 1: All-time exception miles by category.
  if (categoryRange) {
    const categoryChart = sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(categoryRange)
      .setNumHeaders(1)
      .setOption('title', 'All-Time Exception Miles by Category')
      .setOption('hAxis', { title: 'Exception Category' })
      .setOption('vAxis', {
        title: 'Exception Miles',
        minValue: 0
      })
      .setOption('legend', { position: 'none' })
      .setOption('width', 650)
      .setOption('height', 380)
      .setPosition(1, 5, 0, 0)
      .build();

    sheet.insertChart(categoryChart);
  }

  // Chart 2: Monthly exception miles trend.
  if (monthlyRange) {
    const monthlyChart = sheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(monthlyRange)
      .setNumHeaders(1)
      .setOption('title', 'Monthly Exception Miles Trend')
      .setOption('hAxis', {
        title: 'Month',
        slantedText: true
      })
      .setOption('vAxis', {
        title: 'Exception Miles',
        minValue: 0
      })
      .setOption('legend', { position: 'none' })
      .setOption('width', 650)
      .setOption('height', 380)
      .setPosition(21, 5, 0, 0)
      .build();

    sheet.insertChart(monthlyChart);
  }
}


function findMissionOverviewChartData(sheet, type) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 3) return null;

  const columnA = sheet
    .getRange(1, 1, lastRow, 1)
    .getDisplayValues()
    .flat();

  // Find the relevant section heading in column A.
  const titleRowIndex = columnA.findIndex(value => {
    const title = String(value).trim().toLowerCase();

    if (type === 'category') {
      return title.includes('miles by category');
    }

    return title.includes('monthly totals');
  });

  if (titleRowIndex === -1) {
    throw new Error(
      'Could not find the ' + type +
      ' chart section in Mission Overview.'
    );
  }

  const titleRow = titleRowIndex + 1;
  const headerRow = titleRow + 1;
  const firstDataRow = headerRow + 1;

  // Find the last data row using columns A and B.
  let endRow = firstDataRow - 1;

  for (let row = firstDataRow; row <= lastRow; row++) {
    const values = sheet.getRange(row, 1, 1, 2).getDisplayValues()[0];

    if (values[0] === '' || values[1] === '') break;

    endRow = row;
  }

  if (endRow < firstDataRow) return null;

  // Include the header row for chart labels.
  return sheet.getRange(
    headerRow,
    1,
    endRow - headerRow + 1,
    2
  );
}


function organizeMileageSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Get the current month tab name.
  const currentMonth = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'MMMM yyyy'
  );

  // These tabs should always appear first.
  const firstSheets = [
    'Mission Overview',
    'Monthly Overview',
    'Monthly Summary',
    currentMonth,
    'Form Responses 1'
  ];

  // Find all previous monthly tabs.
  const monthPattern =
    /^(January|February|March|April|May|June|July|August|September|October|November|December) \d{4}$/;

  const previousMonths = ss.getSheets()
    .map(sheet => sheet.getName())
    .filter(name =>
      monthPattern.test(name) &&
      name !== currentMonth
    )
    .sort((a, b) => {
      return new Date(b + ' 1') - new Date(a + ' 1');
    });

  // These tabs should always appear at the end.
  const lastSheets = [
    'Contacts Import',
    'Areas'
  ];

  // Combine the desired order.
  const desiredOrder = [
    ...firstSheets,
    ...previousMonths,
    ...lastSheets
  ];

  // Move each existing sheet into its desired position.
  let position = 1;

  desiredOrder.forEach(name => {
    const sheet = ss.getSheetByName(name);

    if (sheet) {
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(position);
      position++;
    }
  });

  Logger.log('Mileage spreadsheet tabs organized.');
}


function updateSingleAreaMonth(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const responseSheet = e.range.getSheet();
  const responseRow = e.range.getRow();

  // Use a status column to prevent duplicate processing.
  let responseHeaders = responseSheet
    .getRange(1, 1, 1, responseSheet.getLastColumn())
    .getValues()[0];

  let statusColumn = responseHeaders.indexOf(
    'Mileage Summary Processed'
  ) + 1;

  if (statusColumn === 0) {
    statusColumn = responseSheet.getLastColumn() + 1;

    responseSheet
      .getRange(1, statusColumn)
      .setValue('Mileage Summary Processed');
  }

  if (
    responseSheet.getRange(responseRow, statusColumn).getValue()
    === 'PROCESSED'
  ) {
    Logger.log(`Response row ${responseRow} was already processed.`);
    return;
  }

  responseHeaders = responseSheet
    .getRange(1, 1, 1, responseSheet.getLastColumn())
    .getValues()[0];

  const findHeader = (headers, name) =>
    headers.findIndex(h =>
      String(h).trim().toLowerCase() === name.toLowerCase()
    );

  const dateColumn = findHeader(responseHeaders, 'Date of event');
  const areaColumn = findHeader(responseHeaders, 'Area');
  const categoryColumn = findHeader(responseHeaders, 'Mileage Category');

  // The mileage question has been renamed in the Form.
  const milesColumn = findHeader(
    responseHeaders,
    'Number of miles used (Please enter just the number!)'
  );

  if (
    dateColumn < 0 ||
    areaColumn < 0 ||
    categoryColumn < 0 ||
    milesColumn < 0
  ) {
    throw new Error('Could not find all required response columns.');
  }

  const response = responseSheet
    .getRange(responseRow, 1, 1, responseSheet.getLastColumn())
    .getValues()[0];

  const month = normalizeMileageMonth(response[dateColumn]);
  const area = String(response[areaColumn]).trim();
  const category = String(response[categoryColumn]).trim();
  const miles = Number(response[milesColumn]);

  if (!month || !area || !category || !Number.isFinite(miles) || miles === 0) {
    throw new Error(
      `Invalid mileage response in row ${responseRow}.`
    );
  }

  const summarySheet = ss.getSheetByName('Monthly Summary');

  if (!summarySheet) {
    throw new Error('Monthly Summary does not exist. Run a full rebuild first.');
  }

  // Find or add the category column.
  let summaryHeaders = summarySheet
    .getRange(1, 1, 1, summarySheet.getLastColumn())
    .getValues()[0];

  let categoryColumnInSummary = summaryHeaders.indexOf(category) + 1;

  if (categoryColumnInSummary === 0) {
    categoryColumnInSummary = summaryHeaders.length + 1;

    summarySheet
      .getRange(1, categoryColumnInSummary)
      .setValue(category);

    summaryHeaders.push(category);
  }

  const summaryData = summarySheet.getDataRange().getValues();

  const monthColumn = summaryHeaders.indexOf('Month');
  const summaryAreaColumn = summaryHeaders.indexOf('Area');
  const totalColumn = summaryHeaders.indexOf('Total Exception Miles');

  if (monthColumn < 0 || summaryAreaColumn < 0 || totalColumn < 0) {
    throw new Error('Monthly Summary is missing required columns.');
  }

  // Find the existing Month + Area row.
  let targetRow = -1;

  for (let i = 1; i < summaryData.length; i++) {
    const existingMonth = normalizeMileageMonth(summaryData[i][monthColumn]);
    const existingArea = String(summaryData[i][summaryAreaColumn]).trim();

    if (existingMonth === month && existingArea === area) {
      targetRow = i + 1;
      break;
    }
  }

  // Create a new Month + Area row at the top.
if (targetRow === -1) {
  summarySheet.insertRowBefore(2);
  targetRow = 2;

    const areasSheet = ss.getSheetByName('Areas');
    const areasData = areasSheet.getDataRange().getValues();
    const areaHeaders = areasData[0];

    const areaIndex = areaHeaders.indexOf('Area');
    const districtIndex = areaHeaders.indexOf('District');
    const zoneIndex = areaHeaders.indexOf('Zone');
    const vehicleIndex = areaHeaders.indexOf('Vehicle');
    const allowanceIndex = areaHeaders.indexOf('Vehicle Allowance/Mo');

    const areaInfo = areasData.slice(1).find(row =>
      String(row[areaIndex]).trim() === area
    ) || [];

    const newRow = new Array(summaryHeaders.length).fill('');

    newRow[monthColumn] = month;
    newRow[summaryAreaColumn] = area;

    const setIfAvailable = (header, index) => {
      const column = summaryHeaders.indexOf(header);
      if (column >= 0 && index >= 0) {
        newRow[column] = areaInfo[index] || '';
      }
    };

    setIfAvailable('District', districtIndex);
    setIfAvailable('Zone', zoneIndex);
    setIfAvailable('Vehicle', vehicleIndex);
    setIfAvailable('Monthly Allowance', allowanceIndex);

    summarySheet
      .getRange(targetRow, 1, 1, newRow.length)
      .setValues([newRow]);
  }

  // Add the new miles to the existing totals.
  const totalCell = summarySheet.getRange(targetRow, totalColumn + 1);
  const categoryCell = summarySheet.getRange(
    targetRow,
    categoryColumnInSummary
  );

  totalCell.setValue((Number(totalCell.getValue()) || 0) + miles);

  categoryCell.setValue(
    (Number(categoryCell.getValue()) || 0) + miles
  );

  // Refresh only the affected monthly tab.
  refreshSingleMonthlyTab(ss, month);

  // The dashboards are based on Monthly Summary, not raw responses.
  updateMileageDashboards();

  // Mark the response only after successful processing.
  responseSheet
    .getRange(responseRow, statusColumn)
    .setValue('PROCESSED');

  Logger.log(
    `Incremental update: ${area}, ${month}, ${category}, ${miles} miles.`
  );
}


function refreshSingleMonthlyTab(ss, month) {
  const summarySheet = ss.getSheetByName('Monthly Summary');
  const data = summarySheet.getDataRange().getValues();

  const headers = data[0];

  const rows = data.slice(1).filter(row =>
    normalizeMileageMonth(row[0]) === month
  );

  const parts = month.split('-');

  const monthName = Utilities.formatDate(
    new Date(Number(parts[0]), Number(parts[1]) - 1, 1),
    Session.getScriptTimeZone(),
    'MMMM yyyy'
  );

  let sheet = ss.getSheetByName(monthName);

  const isNewSheet = !sheet;

  if (!sheet) {
    sheet = ss.insertSheet(monthName);
  }

  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }

  sheet.clear();

  const output = [headers, ...rows];

  sheet.getRange(
    1, 1, output.length, headers.length
  ).setValues(output);

  for (let col = 1; col <= headers.length; col++) {
    sheet.setColumnWidth(
      col,
      summarySheet.getColumnWidth(col)
    );
  }

  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet.setFrozenRows(1);

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 1)
      .setNumberFormat('mmmm yyyy');

    if (headers.length >= 7) {
      sheet.getRange(
        2, 7, rows.length, headers.length - 6
      ).setNumberFormat('0');
    }
  }

  sheet.getRange(
    1, 1, output.length, headers.length
  ).createFilter();

  if (headers.length >= 2) {
    sheet.setColumnWidth(2, 180);
  }

  // Only reorganize tabs when a new month is created.
  if (isNewSheet) {
    organizeMileageSheets();
  }
}


