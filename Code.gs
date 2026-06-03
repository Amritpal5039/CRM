function doGet() {
  return HtmlService.createTemplateFromFile("Index")
    .evaluate()
    .setTitle("Customer Care Dashboard")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

const SHEET_NAME = "Customers";

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      "ID",
      "Name",
      "City",
      "Phone Number",
      "Enquiry Date",
      "Source",
      "Status",
      "Branch",
      "Contact Again Date",
      "Conversations",
      "Registration ID",
    ]);
  }
  return sheet;
}

function login(username, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let idSheet = ss.getSheetByName("ID");
  if (!idSheet) {
    idSheet = ss.insertSheet("ID");
    idSheet.appendRow(["Username", "Password", "Name"]);
    idSheet.appendRow(["admin", "admin123", "Admin User"]); // default user
  }

  const data = idSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == username && data[i][1] == password) {
      return { success: true, name: data[i][2] || username };
    }
  }
  return { success: false, message: "Invalid credentials" };
}

// Add new customer data
function addCustomer(data) {
  const sheet = getSheet();
  const id = Utilities.getUuid();

  sheet.appendRow([
    id,
    data.name,
    data.city,
    data.phone,
    data.enquiryDate,
    data.source,
    "Pending", // Default Status
    "", // Branch
    "", // Contact Again Date
    "[]", // empty conversations JSON
    "", // Registration ID
  ]);
  return { success: true, message: "Customer added successfully" };
}

function formatDateToDayMonthYear(dateObj) {
  if (!dateObj) return "";
  if (!(dateObj instanceof Date)) {
    return dateObj; // string fallback
  }
  const day = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "d");
  const monthStr = Utilities.formatDate(
    dateObj,
    Session.getScriptTimeZone(),
    "MMMM",
  );
  const yearStr = Utilities.formatDate(
    dateObj,
    Session.getScriptTimeZone(),
    "yy",
  );

  let suffix = "th";
  if (day === "1" || day === "21" || day === "31") suffix = "st";
  else if (day === "2" || day === "22") suffix = "nd";
  else if (day === "3" || day === "23") suffix = "rd";

  return day + suffix + " " + monthStr.toLowerCase() + " " + yearStr;
}

// Unified function for fetching leads with filters, sorting, and pagination
function getFilteredLeads(params) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { leads: [], totalPages: 0, currentPage: 1 };

  const rows = data.slice(1);
  let results = [];

  const sName = (params.name || "").trim().toLowerCase();
  const sPhone = (params.phone || "").trim().toLowerCase();
  const sCity = (params.city || "").trim().toLowerCase();
  const isTodayOnly = params.isTodayOnly || false;
  const excludeConfirmed = params.excludeConfirmed || false;
  const requireStatus = params.requireStatus || null;
  const filterStatus = params.status || null; // For the "All Leads" dropdown

  // HTML date inputs return "YYYY-MM-DD". We can compare dates lexicographically since they are zero-padded.
  const startDateStr = params.startDate ? params.startDate : null;
  const endDateStr = params.endDate ? params.endDate : null;

  const todayDateStr = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd",
  );

  rows.forEach((row, index) => {
    let match = true;
    const name = row[1] ? String(row[1]).toLowerCase() : "";
    const city = row[2] ? String(row[2]).toLowerCase() : "";
    const phone = row[3] ? String(row[3]).toLowerCase() : "";
    const statusStr = String(row[6]).trim();
    let enqDate = row[4];
    let contactAgain = row[8];

    // Status filters
    if (excludeConfirmed && statusStr.toLowerCase() === "confirmed")
      match = false;
    if (
      requireStatus &&
      statusStr.toLowerCase() !== requireStatus.toLowerCase()
    )
      match = false;
    if (
      match &&
      filterStatus &&
      statusStr.toLowerCase() !== filterStatus.toLowerCase()
    )
      match = false;

    // Registration ID filter
    if (match && params.regIdFilter) {
      const regId = String(row[10] || "").trim();
      if (params.regIdFilter === "with" && !regId) match = false;
      if (params.regIdFilter === "without" && regId) match = false;
    }

    // Text filters
    if (match && sName && !name.includes(sName)) match = false;
    if (match && sPhone && !phone.includes(sPhone)) match = false;
    if (match && sCity && !city.includes(sCity)) match = false;

    // Date Range filters (comparing YYYY-MM-DD strings directly)
    if (match && (startDateStr || endDateStr)) {
      let dStr = "";
      if (enqDate instanceof Date) {
        dStr = Utilities.formatDate(
          enqDate,
          Session.getScriptTimeZone(),
          "yyyy-MM-dd",
        );
      } else if (typeof enqDate === "string" && enqDate) {
        // If it's a string, attempt to extract YYYY-MM-DD
        dStr = enqDate.substring(0, 10);
      }

      if (!dStr) {
        match = false; // Exclude if no date but filter is applied
      } else {
        if (startDateStr && dStr < startDateStr) match = false;
        if (endDateStr && dStr > endDateStr) match = false;
      }
    }

    // Today Only filter logic
    if (match && isTodayOnly) {
      let isToday = false;

      if (contactAgain) {
        if (contactAgain instanceof Date) {
          const dStr = Utilities.formatDate(
            contactAgain,
            Session.getScriptTimeZone(),
            "yyyy-MM-dd",
          );
          if (dStr === todayDateStr) isToday = true;
        } else if (
          typeof contactAgain === "string" &&
          contactAgain.startsWith(todayDateStr)
        ) {
          isToday = true;
        }
      } else {
        if (enqDate instanceof Date) {
          const dStr = Utilities.formatDate(
            enqDate,
            Session.getScriptTimeZone(),
            "yyyy-MM-dd",
          );
          if (dStr === todayDateStr) isToday = true;
        } else if (
          typeof enqDate === "string" &&
          enqDate.startsWith(todayDateStr)
        ) {
          isToday = true;
        }
      }

      if (!isToday) match = false;
    }

    if (match) {
      let enqDateStrFormatted = formatDateToDayMonthYear(row[4]);
      let contactAgainDateStr = "";
      if (row[8]) {
        if (row[8] instanceof Date) {
          contactAgainDateStr = Utilities.formatDate(
            row[8],
            Session.getScriptTimeZone(),
            "yyyy-MM-dd",
          );
        } else {
          contactAgainDateStr = String(row[8]);
        }
      }

      let conversations = [];
      try {
        conversations = JSON.parse(row[9]) || [];
      } catch (e) {}

      let latestConversation =
        conversations.length > 0
          ? conversations[conversations.length - 1].text
          : "No conversations yet";

      results.push({
        rowIndex: index + 2, // 1-based, +1 for header
        id: row[0],
        name: row[1],
        city: row[2],
        phone: row[3],
        enquiryDate: enqDateStrFormatted,
        source: row[5],
        status: row[6],
        branch: row[7],
        contactAgainDate: contactAgainDateStr,
        conversations: conversations,
        latestConversation: latestConversation,
        registrationId: row[10] || "",
        rawDateForSort:
          row[4] instanceof Date
            ? row[4].getTime()
            : new Date(row[4] || 0).getTime(),
      });
    }
  });

  // Sorting
  const sortOrder = params.sortOrder || "fresh"; // 'fresh' (newest first) or 'oldest'
  if (sortOrder === "fresh") {
    // Largest row index (newest added) first
    results.sort((a, b) => b.rowIndex - a.rowIndex);
  } else {
    // Smallest row index (oldest added) first
    results.sort((a, b) => a.rowIndex - b.rowIndex);
  }

  // Pagination
  const pageSize = 20;
  const page = parseInt(params.page) || 1;
  const totalItems = results.length;
  const totalPages = Math.ceil(totalItems / pageSize);

  const startIndex = (page - 1) * pageSize;
  const paginatedLeads = results.slice(startIndex, startIndex + pageSize);

  return {
    leads: paginatedLeads,
    totalPages: totalPages,
    currentPage: page,
    totalItems: totalItems,
  };
}

// Fetch single user by row index
function getUserDetails(rowIndex) {
  const sheet = getSheet();
  const row = sheet.getRange(rowIndex, 1, 1, 11).getValues()[0]; // Fetch up to 11 cols now

  let enqDateStr = formatDateToDayMonthYear(row[4]);
  let contactAgainDateStr = "";
  if (row[8]) {
    if (row[8] instanceof Date) {
      contactAgainDateStr = Utilities.formatDate(
        row[8],
        Session.getScriptTimeZone(),
        "yyyy-MM-dd",
      );
    } else {
      contactAgainDateStr = String(row[8]);
    }
  }

  let conversations = [];
  try {
    conversations = JSON.parse(row[9]) || [];
  } catch (e) {}

  return {
    rowIndex: rowIndex,
    id: row[0],
    name: row[1],
    city: row[2],
    phone: row[3],
    enquiryDate: enqDateStr,
    source: row[5],
    status: row[6],
    branch: row[7],
    contactAgainDate: contactAgainDateStr,
    conversations: conversations,
    registrationId: row[10] || "",
  };
}

function addConversation(
  rowIndex,
  text,
  agentName,
  status,
  branch,
  contactAgainDate,
  registrationId,
  originalStatus
) {
  const sheet = getSheet();

  if (status && status !== originalStatus) {
    if (!text || text.trim() === "") {
      throw new Error("Conversation notes are required when updating status.");
    }
  }

  if (status) {
    sheet.getRange(rowIndex, 7).setValue(status);
  }

  if (branch !== undefined) sheet.getRange(rowIndex, 8).setValue(branch);
  if (contactAgainDate !== undefined)
    sheet.getRange(rowIndex, 9).setValue(contactAgainDate);
  if (registrationId !== undefined)
    sheet.getRange(rowIndex, 11).setValue(registrationId);

  const convCell = sheet.getRange(rowIndex, 10); // Column J
  let conversations = [];
  try {
    const val = convCell.getValue();
    if (val) conversations = JSON.parse(val);
  } catch (e) {}

  if (text && text.trim() !== "") {
    const dateStr = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd HH:mm:ss",
    );

    conversations.push({
      text: text.trim(),
      date: dateStr,
      agent: agentName,
    });

    convCell.setValue(JSON.stringify(conversations));
  }

  return {
    success: true,
    message: "Data updated",
    conversations: conversations,
    status: status,
    branch: branch,
    contactAgainDate: contactAgainDate,
    registrationId: registrationId,
  };
}

function getDashboardStats(startDateStr, endDateStr) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { total: 0, confirmed: 0, timeline: {} };

  const rows = data.slice(1);
  let start = startDateStr ? new Date(startDateStr) : new Date(0);
  let end = endDateStr ? new Date(endDateStr) : new Date();

  end.setHours(23, 59, 59, 999);

  let total = 0;
  let confirmed = 0;
  let timeline = {};

  rows.forEach((row) => {
    let enqDate = row[4];
    let dStr = "";

    if (enqDate instanceof Date) {
      dStr = Utilities.formatDate(
        enqDate,
        Session.getScriptTimeZone(),
        "yyyy-MM-dd",
      );
    } else if (typeof enqDate === "string" && enqDate) {
      dStr = enqDate.substring(0, 10);
    }

    if (
      dStr &&
      (!startDateStr || dStr >= startDateStr) &&
      (!endDateStr || dStr <= endDateStr)
    ) {
      total++;
      if (String(row[6]).toLowerCase() === "confirmed") confirmed++;

      if (!timeline[dStr]) timeline[dStr] = { total: 0, confirmed: 0 };

      timeline[dStr].total++;
      if (String(row[6]).toLowerCase() === "confirmed")
        timeline[dStr].confirmed++;
    }
  });

  const sortedTimeline = {};
  Object.keys(timeline)
    .sort()
    .forEach((k) => {
      sortedTimeline[k] = timeline[k];
    });

  return { total: total, confirmed: confirmed, timeline: sortedTimeline };
}
