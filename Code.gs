function doGet() {
  // Check and run migration if old 'Customers' sheet still exists
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName("Customers")) {
    runMigration();
  }

  return HtmlService.createTemplateFromFile("Index")
    .evaluate()
    .setTitle("Customer Care Dashboard")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Keep-alive ping — called by the frontend to prevent session expiry
function ping() {
  return true;
}

const CLIENTS_SHEET = "Clients";
const LEADS_SHEET = "Leads";

function getClientsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CLIENTS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CLIENTS_SHEET);
    sheet.appendRow([
      "Client ID",
      "Name",
      "City",
      "Phone Number",
      "Registration ID",
      "First Enquiry Date",
      "Original Source",
      "Edit History"
    ]);
  }
  return sheet;
}

function getLeadsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LEADS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(LEADS_SHEET);
    sheet.appendRow([
      "Lead ID",
      "Client ID",
      "Enquiry Date",
      "Source",
      "Status",
      "Branch",
      "Contact Again Date",
      "Conversations",
      "Lead By",
      "Client Type", // "New" or "Returning"
      "Arrival Status" // "Reached", "Not Reached", "Closed"
    ]);
  }
  return sheet;
}

// Temporary compatibility wrapper if needed during migration
function getSheet() {
  return getClientsSheet(); // Defaulting to something to avoid breaks, will be refactored
}

function runMigration() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const oldSheet = ss.getSheetByName("Customers");
  if (!oldSheet) return; // No migration needed
  
  const clientsSheet = getClientsSheet();
  const leadsSheet = getLeadsSheet();
  
  const oldData = oldSheet.getDataRange().getValues();
  if (oldData.length <= 1) return; // Only headers or empty
  
  // To avoid duplicates, track phones we've seen
  const existingClients = {};
  
  for (let i = 1; i < oldData.length; i++) {
    const row = oldData[i];
    const id = row[0];
    const name = row[1];
    const city = row[2];
    const phone = String(row[3]).trim();
    const enqDate = row[4];
    const source = row[5];
    const status = row[6];
    const branch = row[7];
    const cad = row[8];
    const convs = row[9];
    const regId = row[10];
    const leadBy = row[11];
    const editHist = row[12];
    
    let clientId;
    let clientType = "New";
    
    if (existingClients[phone]) {
      clientId = existingClients[phone];
      clientType = "Returning";
    } else {
      clientId = Utilities.getUuid();
      existingClients[phone] = clientId;
      clientsSheet.appendRow([
        clientId, name, city, phone, regId, enqDate, source, editHist
      ]);
    }
    
    const leadId = id || Utilities.getUuid();
    leadsSheet.appendRow([
      leadId, clientId, enqDate, source, status, branch, cad, convs, leadBy, clientType
    ]);
  }
  
  // Rename old sheet so migration doesn't run twice
  oldSheet.setName("Customers_Migrated_Backup");
}

function login(username, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let idSheet = ss.getSheetByName("ID");
  if (!idSheet) {
    idSheet = ss.insertSheet("ID");
    idSheet.appendRow(["Username", "Password", "Name", "Role"]);
    idSheet.appendRow(["admin", "admin123", "Admin User", "Admin"]); // default user
  }

  const data = idSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == username && data[i][1] == password) {
      // Check if user is an admin by looking at the 4th column (Role). 
      // If the column is missing/empty, default to standard user.
      const role = String(data[i][3] || "").trim().toLowerCase();
      const isAdmin = (role === 'admin' || username === 'admin'); 
      return { success: true, name: data[i][2] || username, isAdmin: isAdmin };
    }
  }
  return { success: false, message: "Invalid credentials" };
}

// Admin only function to delete a lead
function deleteLead(rowIndex, username) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const idSheet = ss.getSheetByName("ID");
  let isAuthorized = false;

  if (idSheet) {
    const data = idSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == username) {
        const role = String(data[i][3] || "").trim().toLowerCase();
        if (role === 'admin' || username === 'admin') {
          isAuthorized = true;
        }
        break;
      }
    }
  }

  if (!isAuthorized) {
    throw new Error("Unauthorized: Only users with the Admin role can delete leads.");
  }
  
  const rIndex = parseInt(rowIndex, 10);
  if (isNaN(rIndex)) {
    throw new Error("Invalid row index.");
  }

  const sheet = getLeadsSheet();
  // Ensure we don't delete headers
  if (rIndex <= 1) {
    throw new Error("Cannot delete header row.");
  }
  
  sheet.deleteRow(rIndex);
  return { success: true, message: "Lead deleted successfully." };
}

// Helper to get client by phone
function getClientByPhone(phone) {
  const clientsSheet = getClientsSheet();
  const data = clientsSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][3]).trim() === String(phone).trim()) {
      let enqDateStr = data[i][5];
      if (enqDateStr instanceof Date) {
         enqDateStr = Utilities.formatDate(enqDateStr, Session.getScriptTimeZone(), "yyyy-MM-dd");
      }
      return {
        rowIndex: i + 1,
        clientId: data[i][0],
        name: data[i][1],
        city: data[i][2],
        phone: data[i][3],
        regId: data[i][4],
        enqDate: enqDateStr,
        source: data[i][6]
      };
    }
  }
  return null;
}

function isOpenLeadRow(row) {
  const status = String(row[4] || "Pending").trim().toLowerCase();
  const arrivalStatus = String(row[10] || "").trim().toLowerCase();

  if (status === "closed") return false;
  if (
    status === "confirmed" &&
    ["reached", "not reached", "closed"].indexOf(arrivalStatus) !== -1
  ) {
    return false;
  }

  return true;
}

function getLastOpenLeadForClient(clientId) {
  const leadsSheet = getLeadsSheet();
  const leadsData = leadsSheet.getDataRange().getValues();
  let lastOpenLead = null;

  for (let i = 1; i < leadsData.length; i++) {
    const row = leadsData[i];
    if (row[1] !== clientId || !isOpenLeadRow(row)) continue;

    let conversations = [];
    try {
      conversations = JSON.parse(row[7]) || [];
    } catch (e) {}

    const rawDate = row[2] instanceof Date ? row[2].getTime() : new Date(row[2] || 0).getTime();
    const lead = {
      rowIndex: i + 1,
      enquiryDate: formatDateToDayMonthYear(row[2]),
      status: row[4] || "Pending",
      branch: row[5] || "",
      contactAgainDate: row[6] instanceof Date
        ? Utilities.formatDate(row[6], Session.getScriptTimeZone(), "yyyy-MM-dd")
        : String(row[6] || ""),
      latestConversation: conversations.length > 0
        ? conversations[conversations.length - 1].text
        : "No conversations yet",
      rawDate: isNaN(rawDate) ? 0 : rawDate
    };

    if (!lastOpenLead || lead.rowIndex > lastOpenLead.rowIndex) {
      lastOpenLead = lead;
    }
  }

  return lastOpenLead;
}

// Called from UI when caller enters phone number
function checkPhoneExists(phone) {
  const client = getClientByPhone(phone);
  if (client) {
    return {
      exists: true,
      client: client,
      openLead: getLastOpenLeadForClient(client.clientId)
    };
  }
  return { exists: false };
}

// Add new lead for existing client
function addLeadForExistingClient(clientId, data) {
  const clientsSheet = getClientsSheet();
  const leadsSheet = getLeadsSheet();
  const leadId = Utilities.getUuid();
  const clientsData = clientsSheet.getDataRange().getValues();
  let client = null;

  for (let i = 1; i < clientsData.length; i++) {
    if (clientsData[i][0] === clientId) {
      client = clientsData[i];
      break;
    }
  }

  if (!client) {
    throw new Error("Client not found.");
  }
  
  leadsSheet.appendRow([
    leadId,
    clientId,
    data.enquiryDate,
    client[6] || "", // Returning leads keep the client's original source.
    "Pending", // Default Status
    "", // Branch
    "", // Contact Again Date
    "[]", // empty conversations JSON
    data.leadBy || "", 
    "Returning",
    "" // Arrival Status
  ]);
  
  return { success: true, message: "New lead added to existing client successfully." };
}

// Add completely new client and their first lead
function addCustomer(data) {
  const clientsSheet = getClientsSheet();
  const leadsSheet = getLeadsSheet();
  
  // Check if they slipped through
  if (getClientByPhone(data.phone)) {
     return { success: false, message: "Client already exists. Please use 'Add Lead to Existing Client'." };
  }

  const clientId = Utilities.getUuid();
  const leadId = Utilities.getUuid();

  clientsSheet.appendRow([
    clientId,
    data.name,
    data.city,
    data.phone,
    "", // Reg ID
    data.enquiryDate,
    data.source,
    "[]" // Edit History
  ]);
  
  leadsSheet.appendRow([
    leadId,
    clientId,
    data.enquiryDate,
    data.source,
    "Pending", // Default Status
    "", // Branch
    "", // Contact Again Date
    "[]", // empty conversations JSON
    data.leadBy || "", // Lead By
    "New",
    "" // Arrival Status
  ]);

  return { success: true, message: "New client and lead created successfully" };
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

// Function for fetching clients for the Clients Database section
function getFilteredClients(params) {
  const clientsSheet = getClientsSheet();
  const leadsSheet = getLeadsSheet();
  const clientsData = clientsSheet.getDataRange().getValues();
  const leadsData = leadsSheet.getDataRange().getValues();
  
  if (clientsData.length <= 1) return { clients: [] };
  
  const sName = (params.name || "").trim().toLowerCase();
  const sPhone = (params.phone || "").trim().toLowerCase();
  
  // Map lead counts
  const leadCounts = {};
  for(let i=1; i < leadsData.length; i++) {
    const cid = leadsData[i][1];
    if(!leadCounts[cid]) leadCounts[cid] = 0;
    leadCounts[cid]++;
  }

  let results = [];
  const uniquePhones = new Set();
  
  // Sort to process newer entries first, or keep order. We'll loop normally
  for(let i=1; i < clientsData.length; i++) {
    const row = clientsData[i];
    const cid = row[0];
    const count = leadCounts[cid] || 0;
    
    // Discard clients with 0 leads
    if (count === 0) continue;
    
    const phoneRaw = String(row[3]).trim();
    const phone = phoneRaw.toLowerCase();
    
    // Prevent UI duplicates if they somehow exist in DB
    if (uniquePhones.has(phone)) continue;
    uniquePhones.add(phone);
    
    let match = true;
    const name = String(row[1]).toLowerCase();
    
    if (sName && !name.includes(sName)) match = false;
    if (sPhone && !phone.includes(sPhone)) match = false;
    
    if (match) {
      results.push({
        clientId: cid,
        name: row[1],
        city: row[2],
        phone: phoneRaw,
        regId: row[4],
        leadCount: count
      });
    }
  }
  
  return { clients: results };
}

// Unified function for fetching leads with filters, sorting, and pagination
function getFilteredLeads(params) {
  const clientsSheet = getClientsSheet();
  const leadsSheet = getLeadsSheet();
  
  const clientsData = clientsSheet.getDataRange().getValues();
  const leadsData = leadsSheet.getDataRange().getValues();
  
  if (leadsData.length <= 1 || clientsData.length <= 1) return { leads: [], totalPages: 0, currentPage: 1 };

  // Map clients by ID for easy lookup
  const clientsMap = {};
  for(let i = 1; i < clientsData.length; i++) {
    const row = clientsData[i];
    clientsMap[row[0]] = {
      rowIndex: i + 1,
      name: row[1],
      city: row[2],
      phone: row[3],
      regId: row[4],
      enqDate: row[5],
      source: row[6],
      editHistory: row[7]
    };
  }

  const rows = leadsData.slice(1);
  let results = [];

  const sName = (params.name || "").trim().toLowerCase();
  const sPhone = (params.phone || "").trim().toLowerCase();
  const sCity = (params.city || "").trim().toLowerCase();
  const isTodayOnly = params.isTodayOnly || false;
  const excludeConfirmed = params.excludeConfirmed || false;
  const requireStatus = params.requireStatus || null;
  const filterStatus = params.status || null;

  const startDateStr = params.startDate ? params.startDate : null;
  const endDateStr = params.endDate ? params.endDate : null;

  const todayDateStr = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd",
  );

  rows.forEach((row, index) => {
    const clientId = row[1];
    const client = clientsMap[clientId];
    if (!client) return; // Orphaned lead?
    
    let match = true;
    const name = client.name ? String(client.name).toLowerCase() : "";
    const city = client.city ? String(client.city).toLowerCase() : "";
    const phone = client.phone ? String(client.phone).toLowerCase() : "";
    
    const enqDate = row[2];
    const source = row[3];
    const statusStr = String(row[4] || "Pending").trim();
    const branch = row[5];
    const contactAgain = row[6];
    const leadBy = row[8];
    const clientType = row[9];
    const regId = String(client.regId || "").trim();

    // Status filters
    if (excludeConfirmed && statusStr.toLowerCase() === "confirmed") match = false;
    if (requireStatus && statusStr.toLowerCase() !== requireStatus.toLowerCase()) match = false;
    if (match && filterStatus && statusStr.toLowerCase() !== filterStatus.toLowerCase()) match = false;

    // Confirmation Detail filter
    if (match && params.regIdFilter) {
      const arrStatus = String(row[10] || "").trim().toLowerCase(); // Arrival Status
      if (params.regIdFilter === "with" && !regId) match = false;
      if (params.regIdFilter === "without" && regId) match = false;
      if (params.regIdFilter === "reached" && arrStatus !== "reached") match = false;
      if (params.regIdFilter === "not_reached" && arrStatus !== "not reached") match = false;
    }

    // Text filters
    if (match && sName && !name.includes(sName)) match = false;
    if (match && sPhone && !phone.includes(sPhone)) match = false;
    if (match && sCity && !city.includes(sCity)) match = false;

    // Date Range filters
    if (match && (startDateStr || endDateStr)) {
      let dStr = "";
      if (enqDate instanceof Date) {
        dStr = Utilities.formatDate(enqDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else if (typeof enqDate === "string" && enqDate) {
        dStr = enqDate.substring(0, 10);
      }

      if (!dStr) {
        match = false;
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
          const dStr = Utilities.formatDate(contactAgain, Session.getScriptTimeZone(), "yyyy-MM-dd");
          if (dStr === todayDateStr) isToday = true;
        } else if (typeof contactAgain === "string" && contactAgain.startsWith(todayDateStr)) {
          isToday = true;
        }
      } else {
        if (enqDate instanceof Date) {
          const dStr = Utilities.formatDate(enqDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
          if (dStr === todayDateStr) isToday = true;
        } else if (typeof enqDate === "string" && enqDate.startsWith(todayDateStr)) {
          isToday = true;
        }
      }

      if (!isToday) {
        match = false;
      } else {
        if (statusStr.toLowerCase() === "confirmed" && regId !== "") {
          match = false;
        }
      }
    }

    if (match) {
      let enqDateStrFormatted = formatDateToDayMonthYear(enqDate);
      let contactAgainDateStr = "";
      if (contactAgain) {
        if (contactAgain instanceof Date) {
          contactAgainDateStr = Utilities.formatDate(contactAgain, Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else {
          contactAgainDateStr = String(contactAgain);
        }
      }

      let conversations = [];
      try {
        conversations = JSON.parse(row[7]) || [];
      } catch (e) {}

      let latestConversation = conversations.length > 0
          ? conversations[conversations.length - 1].text
          : "No conversations yet";

      results.push({
        rowIndex: index + 2, // 1-based, +1 for header inside LEADS sheet
        id: row[0],
        clientId: clientId,
        clientRowIndex: client.rowIndex, // to update client details
        name: client.name,
        city: client.city,
        phone: client.phone,
        enquiryDate: enqDateStrFormatted,
        source: source,
        status: statusStr,
        branch: branch,
        contactAgainDate: contactAgainDateStr,
        conversations: conversations,
        latestConversation: latestConversation,
        registrationId: regId,
        leadBy: leadBy,
        clientType: clientType,
        arrivalStatus: String(row[10] || "").trim(),
        rawDateForSort: enqDate instanceof Date ? enqDate.getTime() : new Date(enqDate || 0).getTime(),
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

function getClientDetailsWithLeads(clientId) {
  const clientsSheet = getClientsSheet();
  const leadsSheet = getLeadsSheet();
  
  const clientsData = clientsSheet.getDataRange().getValues();
  const leadsData = leadsSheet.getDataRange().getValues();
  
  let clientRow = null;
  for (let i = 1; i < clientsData.length; i++) {
    if (clientsData[i][0] === clientId) {
      clientRow = clientsData[i];
      break;
    }
  }
  
  if (!clientRow) throw new Error("Client not found.");

  let enqDateStr = formatDateToDayMonthYear(clientRow[5]);
  
  let editHistory = [];
  try { editHistory = JSON.parse(clientRow[7]) || []; } catch(e) {}

  let leads = [];
  for (let i = 1; i < leadsData.length; i++) {
    if (leadsData[i][1] === clientId) {
      const row = leadsData[i];
      let convs = [];
      try { convs = JSON.parse(row[7]) || []; } catch(e) {}
      
      let latestConv = convs.length > 0 ? convs[convs.length - 1].text : "No conversations yet";
      let leadEnqDate = formatDateToDayMonthYear(row[2]);
      
      leads.push({
        rowIndex: i + 1,
        enquiryDate: leadEnqDate,
        source: row[3],
        status: row[4],
        branch: row[5],
        latestConversation: latestConv,
        leadBy: row[8],
        clientType: row[9],
        rawDate: row[2] instanceof Date ? row[2].getTime() : new Date(row[2] || 0).getTime()
      });
    }
  }
  
  // Sort leads newest first
  leads.sort((a, b) => b.rawDate - a.rawDate);

  return {
    clientId: clientRow[0],
    name: clientRow[1],
    city: clientRow[2],
    phone: clientRow[3],
    registrationId: clientRow[4] || "",
    enquiryDate: enqDateStr,
    source: clientRow[6],
    editHistory: editHistory,
    leads: leads
  };
}

// Fetch single user by row index
function getUserDetails(leadRowIndex) {
  const rIndex = parseInt(leadRowIndex, 10);
  if (isNaN(rIndex)) {
    throw new Error("Invalid lead row index.");
  }
  const leadsSheet = getLeadsSheet();
  const leadRow = leadsSheet.getRange(rIndex, 1, 1, 11).getValues()[0];
  
  const clientId = leadRow[1];
  
  const clientsSheet = getClientsSheet();
  const clientsData = clientsSheet.getDataRange().getValues();
  
  let clientRow = null;
  let clientRowIndex = -1;
  for (let i = 1; i < clientsData.length; i++) {
    if (clientsData[i][0] === clientId) {
      clientRow = clientsData[i];
      clientRowIndex = i + 1;
      break;
    }
  }
  
  if (!clientRow) throw new Error("Client not found for this lead.");

  let enqDateStr = formatDateToDayMonthYear(leadRow[2]);
  let contactAgainDateStr = "";
  if (leadRow[6]) {
    if (leadRow[6] instanceof Date) {
      contactAgainDateStr = Utilities.formatDate(leadRow[6], Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else {
      contactAgainDateStr = String(leadRow[6]);
    }
  }

  let conversations = [];
  try {
    conversations = JSON.parse(leadRow[7]) || [];
  } catch (e) {}
  
  let editHistory = [];
  try {
    editHistory = JSON.parse(clientRow[7]) || [];
  } catch(e) {}

  return {
    rowIndex: leadRowIndex,
    clientRowIndex: clientRowIndex,
    id: leadRow[0],
    clientId: clientRow[0],
    name: clientRow[1],
    city: clientRow[2],
    phone: clientRow[3],
    enquiryDate: enqDateStr,
    source: leadRow[3],
    status: leadRow[4],
    branch: leadRow[5],
    contactAgainDate: contactAgainDateStr,
    conversations: conversations,
    registrationId: clientRow[4] || "",
    leadBy: leadRow[8] || "",
    clientType: leadRow[9] || "",
    arrivalStatus: String(leadRow[10] || "").trim(),
    editHistory: editHistory
  };
}

function addConversation(
  leadRowIndex,
  text,
  agentName,
  status,
  branch,
  contactAgainDate,
  registrationId,
  originalStatus,
  name,
  city,
  originalName,
  originalCity,
  clientRowIndex,
  arrivalStatus
) {
  const rIndex = parseInt(leadRowIndex, 10);
  const cIndex = parseInt(clientRowIndex, 10);
  if (isNaN(rIndex)) {
    throw new Error("Invalid lead row index.");
  }
  if (isNaN(cIndex)) {
    throw new Error("Invalid client row index.");
  }

  const leadsSheet = getLeadsSheet();
  const clientsSheet = getClientsSheet();

  // Update Lead
  if (status) leadsSheet.getRange(rIndex, 5).setValue(status);
  if (branch !== undefined) leadsSheet.getRange(rIndex, 6).setValue(branch);
  if (contactAgainDate !== undefined) leadsSheet.getRange(rIndex, 7).setValue(contactAgainDate);
  if (arrivalStatus !== undefined) leadsSheet.getRange(rIndex, 11).setValue(arrivalStatus);
  
  // Update Client
  if (registrationId !== undefined) clientsSheet.getRange(cIndex, 5).setValue(registrationId);
  if (name !== undefined) clientsSheet.getRange(cIndex, 2).setValue(name);
  if (city !== undefined) clientsSheet.getRange(cIndex, 3).setValue(city);

  const dateStr = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd HH:mm:ss",
  );

  const convCell = leadsSheet.getRange(rIndex, 8); // Column H
  let conversations = [];
  try {
    const val = convCell.getValue();
    if (val) conversations = JSON.parse(val);
  } catch (e) {}

  if (text && text.trim() !== "") {
    conversations.push({
      text: text.trim(),
      date: dateStr,
      agent: agentName,
    });
    convCell.setValue(JSON.stringify(conversations));
  }

  // Handle Edit History on Client Record
  const historyCell = clientsSheet.getRange(cIndex, 8); // Column H
  let editHistory = [];
  try {
    const hVal = historyCell.getValue();
    if (hVal) editHistory = JSON.parse(hVal);
  } catch (e) {}

  let edited = false;
  let editNotes = [];
  if (name && name !== originalName) { editNotes.push(`Name changed from "${originalName}" to "${name}"`); edited = true; }
  if (city && city !== originalCity) { editNotes.push(`City changed from "${originalCity}" to "${city}"`); edited = true; }
  
  if (edited) {
    editHistory.push({
      date: dateStr,
      agent: agentName,
      changes: editNotes.join(", ")
    });
    historyCell.setValue(JSON.stringify(editHistory));
  }

  return {
    success: true,
    message: "Data updated",
    conversations: conversations,
    editHistory: editHistory,
    status: status,
    branch: branch,
    contactAgainDate: contactAgainDate,
    registrationId: registrationId,
    name: name,
    city: city
  };
}

function getDashboardStats(startDateStr, endDateStr) {
  const clientsSheet = getClientsSheet();
  const leadsSheet = getLeadsSheet();
  const clientsData = clientsSheet.getDataRange().getValues();
  const leadsData = leadsSheet.getDataRange().getValues();
  
  if (leadsData.length <= 1) return { 
    total: 0, confirmed: 0, confirmedWithReg: 0, confirmedWithoutReg: 0, 
    pending: 0, followUp: 0, notInterested: 0, 
    closed: 0, callNotAnswered: 0, generalInquiry: 0, callDropped: 0, patientWillContact: 0,
    overallConversionRate: "0.00", newClientConversionRate: "0.00", returningReengagementRate: "0.00",
    timeline: {}, sources: {}, leadBy: {} 
  };

  const clientsMap = {};
  for(let i = 1; i < clientsData.length; i++) {
    clientsMap[clientsData[i][0]] = { 
      regId: clientsData[i][4],
      city: String(clientsData[i][2] || "Unknown").trim()
    };
  }

  const rows = leadsData.slice(1);

  let stats = {
    total: 0,
    newLeadsTotal: 0,
    returningLeadsTotal: 0,
    newLeadsConverted: 0,
    returningLeadsReengaged: 0,
    confirmedWithReg: 0,
    confirmedWithoutReg: 0,
    confirmedReached: 0,      // Reached: both New (with reg) AND Returning (reached)
    confirmedNotReached: 0,   // Not Reached: both New and Returning
    confirmedClosed: 0,
    otherGroup: 0,            // Pending + FollowUp + NotInt + NoAns + Dropped (for bar chart)
    pending: 0,
    followUp: 0,
    notInterested: 0,
    closed: 0,
    callNotAnswered: 0,
    generalInquiry: 0,
    callDropped: 0,
    patientWillContact: 0,
    timeline: {},
    sources: {},
    leadBy: {},
    branchStats: {},
    citiesLead: {},
    citiesConv: {}
  };

  rows.forEach((row) => {
    let enqDate = row[2];
    let dStr = "";

    if (enqDate instanceof Date) {
      dStr = Utilities.formatDate(enqDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else if (typeof enqDate === "string" && enqDate) {
      dStr = enqDate.substring(0, 10);
    }

    if (dStr && (!startDateStr || dStr >= startDateStr) && (!endDateStr || dStr <= endDateStr)) {
      stats.total++;
      
      const clientId = row[1];
      const sourceStr = String(row[3] || "Unknown").trim();
      const statusStr = String(row[4] || "Pending").trim().toLowerCase();
      const leadByStr = String(row[8] || "Unknown").trim();
      const clientType = String(row[9] || "New").trim(); // "New" or "Returning"
      
      const client = clientsMap[clientId] || {};
      const regId = String(client.regId || "").trim();
      const arrivalStatusStr = String(row[10] || "").trim().toLowerCase();
      
      // Basic fuzzy grouping: lowercase and strip non-letters to group spelling variations 
      // (e.g., "New York", "newyork", " New  York")
      let cityRaw = client.city || "Unknown";
      let city = cityRaw.toLowerCase().replace(/[^a-z0-9]/g, '');
      if (city === "") city = "unknown";
      
      let isConverted = false;
      // A lead is "reached" if New+confirmed+arrivalStatus=reached, OR Returning+confirmed+arrivalStatus=reached
      const isReached = statusStr === "confirmed" && arrivalStatusStr === "reached";

      if (clientType === "New") {
        stats.newLeadsTotal++;
        if (statusStr === "confirmed" && regId !== "") {
          stats.newLeadsConverted++;
          isConverted = true;
        }
      } else if (clientType === "Returning") {
        stats.returningLeadsTotal++;
        if (isReached) {
          stats.returningLeadsReengaged++;
          isConverted = true;
        }
      }

      // City breakdown
      if (city) {
        if (!stats.citiesLead[city]) stats.citiesLead[city] = { count: 0, originalName: cityRaw };
        stats.citiesLead[city].count++;
        if (cityRaw.length > stats.citiesLead[city].originalName.length) {
          stats.citiesLead[city].originalName = cityRaw;
        }
        if (isConverted) {
          if (!stats.citiesConv[city]) stats.citiesConv[city] = 0;
          stats.citiesConv[city]++;
        }
      }

      // Status breakdown — unified reached bucket for both New and Returning
      if (statusStr === "confirmed") {
        if (arrivalStatusStr === "reached") {
          // Both New (with reg) AND Returning confirmed+reached count here
          stats.confirmedReached++;
          if (clientType === "New" && regId !== "") stats.confirmedWithReg++; // keep old counter for compat
        } else if (arrivalStatusStr === "not reached") {
          stats.confirmedNotReached++;
        } else if (arrivalStatusStr === "closed") {
          stats.confirmedClosed++;
        } else {
          // New with no arrival status yet, or returning with no arrival — fallback
          stats.confirmedWithoutReg++;
        }
      } else if (statusStr === "pending") {
        stats.pending++;
        stats.otherGroup++;
      } else if (statusStr === "follow up") {
        stats.followUp++;
        stats.otherGroup++;
      } else if (statusStr === "not interested") {
        stats.notInterested++;
        stats.otherGroup++;
      } else if (statusStr === "closed") {
        stats.closed++;
      } else if (statusStr === "call not answered") {
        stats.callNotAnswered++;
        stats.otherGroup++;
      } else if (statusStr === "general inquiry") {
        stats.generalInquiry++;
      } else if (statusStr === "call dropped") {
        stats.callDropped++;
        stats.otherGroup++;
      } else if (statusStr === "patient will contact") {
        stats.patientWillContact++;
      }

      // Source breakdown
      if (!stats.sources[sourceStr]) stats.sources[sourceStr] = 0;
      stats.sources[sourceStr]++;

      // Lead By breakdown
      if (!stats.leadBy[leadByStr]) stats.leadBy[leadByStr] = 0;
      stats.leadBy[leadByStr]++;

      // Branch breakdown (only count leads that have a branch assigned)
      const branchStr = String(row[5] || "").trim();
      if (branchStr) {
        if (!stats.branchStats[branchStr]) stats.branchStats[branchStr] = 0;
        stats.branchStats[branchStr]++;
      }

      // Timeline breakdown — track total and all confirmed+reached (New + Returning)
      if (!stats.timeline[dStr]) stats.timeline[dStr] = { total: 0, reached: 0 };
      stats.timeline[dStr].total++;
      if (isReached) {
        stats.timeline[dStr].reached++;
      }
    }
  });

  const sortedTimeline = {};
  Object.keys(stats.timeline).sort().forEach((k) => { sortedTimeline[k] = stats.timeline[k]; });
  
  stats.timeline = sortedTimeline;
  stats.confirmed = stats.confirmedReached + stats.confirmedNotReached + stats.confirmedClosed + stats.confirmedWithoutReg;
  
  stats.overallConversionRate = stats.total > 0 ? (((stats.newLeadsConverted + stats.returningLeadsReengaged) / stats.total) * 100).toFixed(2) : "0.00";
  stats.newClientConversionRate = stats.newLeadsTotal > 0 ? ((stats.newLeadsConverted / stats.newLeadsTotal) * 100).toFixed(2) : "0.00";
  stats.returningReengagementRate = stats.returningLeadsTotal > 0 ? ((stats.returningLeadsReengaged / stats.returningLeadsTotal) * 100).toFixed(2) : "0.00";

  // Calculate Top 5 Cities
  let cityList = [];
  for (let c in stats.citiesLead) {
    if (c !== "unknown") {
      cityList.push({
        name: stats.citiesLead[c].originalName,
        count: stats.citiesLead[c].count
      });
    }
  }
  
  cityList.sort((a, b) => b.count - a.count);
  stats.topLeadCities = cityList.slice(0, 5);

  return stats;
}
