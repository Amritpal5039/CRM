# Google Apps Script Web App: Code.gs Documentation

This document provides a comprehensive overview of the server-side Google Apps Script functions defined in `Code.gs`.

---

## Table of Contents
1. [doGet](#1-doget)
2. [include](#2-include)
3. [ping](#3-ping)
4. [getClientsSheet](#4-getclientssheet)
5. [getLeadsSheet](#5-getleadssheet)
6. [getSheet](#6-getsheet)
7. [runMigration](#7-runmigration)
8. [login](#8-login)
9. [deleteLead](#9-deletelead)
10. [getClientByPhone](#10-getclientbyphone)
11. [checkPhoneExists](#11-checkphoneexists)
12. [addLeadForExistingClient](#12-addleadforexistingclient)
13. [addCustomer](#13-addcustomer)
14. [formatDateToDayMonthYear](#14-formatdatetodaymonthyear)
15. [getFilteredClients](#15-getfilteredclients)
16. [getFilteredLeads](#16-getfilteredleads)
17. [getClientDetailsWithLeads](#17-getclientdetailswithleads)
18. [getUserDetails](#18-getuserdetails)
19. [addConversation](#19-addconversation)
20. [getDashboardStats](#20-getdashboardstats)

---

## 1. `doGet`

### Description
Serves the Google Apps Script Web App. Before rendering, it checks if a legacy sheet called `"Customers"` exists; if it does, it runs the data migration function `runMigration()`. It then compiles the HTML content of the main template (`Index.html`), configures viewport and scaling meta tags, sets the page title, and allows frame embedding.

* **Parameters:** None
* **Returns:** `HtmlOutput` — The evaluated HTML template for the web app UI.
* **Google Services Used:**
  * `SpreadsheetApp` (to access spreadsheet sheets and check for legacy tables)
  * `HtmlService` (to create, configure, and evaluate HTML templates)
* **Potential Errors:**
  * Throws an error if the `"Index"` HTML template file is missing.
  * Runtime scriptlet exceptions during templating evaluation (`<?!= ... ?>`).
  * Sheet read/access permission errors.

---

## 2. `include`

### Description
Helper utility used inside templating scriptlets to inline the contents of files (such as stylesheet or client-side script templates) directly into `Index.html`.

* **Parameters:**
  * `filename` (String): The file name of the page component/resource to fetch.
* **Returns:** `String` — The raw text/code content of the target file.
* **Google Services Used:**
  * `HtmlService` (to fetch the file and extract its content via `createHtmlOutputFromFile().getContent()`)
* **Potential Errors:**
  * Throws an error if the specified filename does not exist.

---

## 3. `ping`

### Description
A keep-alive function invoked periodically by client-side intervals to prevent the Google Apps Script execution session from timing out.

* **Parameters:** None
* **Returns:** `Boolean` — Always returns `true`.
* **Google Services Used:** None
* **Potential Errors:** None

---

## 4. `getClientsSheet`

### Description
Retrieves the Google Sheet object named `"Clients"`. If the sheet does not exist, it creates the sheet and appends its default column headers.

* **Parameters:** None
* **Returns:** `Sheet` — The Google Apps Script Sheet object mapping the `"Clients"` table.
* **Google Services Used:**
  * `SpreadsheetApp` (to access the active spreadsheet, search for sheets, or insert a new sheet)
* **Potential Errors:**
  * Permission denied if the active script user does not have spreadsheet edit permissions.
  * Sheet name collision/write conflict.

---

## 5. `getLeadsSheet`

### Description
Retrieves the Google Sheet object named `"Leads"`. If the sheet does not exist, it creates the sheet and appends its default column headers.

* **Parameters:** None
* **Returns:** `Sheet` — The Google Apps Script Sheet object mapping the `"Leads"` table.
* **Google Services Used:**
  * `SpreadsheetApp` (to access the active spreadsheet, search for sheets, or insert a new sheet)
* **Potential Errors:**
  * Permission denied if the active script user does not have spreadsheet edit permissions.
  * Sheet name collision/write conflict.

---

## 6. `getSheet`

### Description
A temporary compatibility fallback wrapper that returns the clients sheet object. Used during refactoring to prevent breaking code calling legacy functions.

* **Parameters:** None
* **Returns:** `Sheet` — The clients sheet object.
* **Google Services Used:**
  * Calls `getClientsSheet()`.
* **Potential Errors:** Same as `getClientsSheet()`.

---

## 7. `runMigration`

### Description
Migrates legacy single-sheet CRM data (from the `"Customers"` sheet) to the relational two-sheet database system (`"Clients"` and `"Leads"`). It maps records, groups returning clients using phone numbers, generates unique UUIDs for relationships, and appends the records. Finally, it renames the legacy sheet to `"Customers_Migrated_Backup"` to prevent duplicate runs.

* **Parameters:** None
* **Returns:** `void`
* **Google Services Used:**
  * `SpreadsheetApp` (to read data ranges, write data, insert and rename sheets)
  * `Utilities` (to generate unique IDs using `Utilities.getUuid()`)
* **Potential Errors:**
  * Data mismatch/formatting errors if old cells do not follow expected column structures.
  * Target sheet naming conflicts if `"Customers_Migrated_Backup"` already exists.
  * Timeout error if migrating a massive dataset.

---

## 8. `login`

### Description
Authenticates users against credentials stored in the `"ID"` sheet. If the `"ID"` sheet is missing, it automatically initializes it and appends a default administrative user. Checks matching rows and identifies whether the user holds an administrative role.

* **Parameters:**
  * `username` (String): The username input.
  * `password` (String): The password input.
* **Returns:** `Object` — `{ success: Boolean, name: String, isAdmin: Boolean }` or `{ success: Boolean, message: String }` if validation fails.
* **Google Services Used:**
  * `SpreadsheetApp` (to access or populate the `"ID"` sheet, and read credentials)
* **Potential Errors:**
  * Database access issues.
  * Out-of-bounds errors if credential columns are incomplete.

---

## 9. `deleteLead`

### Description
Permits administrators to delete a specific lead row from the `"Leads"` sheet by its index. Verifies authorization against user roles stored in the `"ID"` sheet before executing deletion.

* **Parameters:**
  * `rowIndex` (Number): The row index in the Leads sheet to delete.
  * `username` (String): The username of the operator requesting deletion.
* **Returns:** `Object` — `{ success: true, message: "Lead deleted successfully." }`
* **Google Services Used:**
  * `SpreadsheetApp` (to query user credentials, access the Leads sheet, and delete rows)
* **Potential Errors:**
  * `Error("Unauthorized: Only users with the Admin role can delete leads.")` if authorization fails.
  * `Error("Cannot delete header row.")` if trying to delete row index <= 1.
  * Range indexing errors if `rowIndex` is out of bounds.

---

## 10. `getClientByPhone`

### Description
Looks up and retrieves a client's record from the `"Clients"` sheet by searching matching phone numbers.

* **Parameters:**
  * `phone` (String/Number): The phone number to search.
* **Returns:** `Object` — Compiled client data fields (e.g. `rowIndex`, `clientId`, `name`, `city`, `phone`, `regId`, `enqDate`, `source`) or `null` if no match is found.
* **Google Services Used:**
  * `SpreadsheetApp` (to read data ranges)
  * `Utilities` (to format date values via `Utilities.formatDate()`)
  * `Session` (to obtain script timezone settings)
* **Potential Errors:**
  * Typecasting issues if the phone number formatting causes comparative mismatches.

---

## 11. `checkPhoneExists`

### Description
Verifies if a client with the input phone number is already registered in the system. Called dynamically by the UI.

* **Parameters:**
  * `phone` (String/Number): The phone number.
* **Returns:** `Object` — `{ exists: Boolean, client: Object }` or `{ exists: Boolean }`.
* **Google Services Used:**
  * Calls `getClientByPhone()`.
* **Potential Errors:** Same as `getClientByPhone()`.

---

## 12. `addLeadForExistingClient`

### Description
Inserts a new lead entry for an already registered client in the `"Leads"` sheet. Generates a new unique `Lead ID` and defaults parameters.

* **Parameters:**
  * `clientId` (String): The unique ID of the target client.
  * `data` (Object): Contains lead criteria (`enquiryDate`, `source`, `leadBy`).
* **Returns:** `Object` — `{ success: true, message: "New lead added to existing client successfully." }`
* **Google Services Used:**
  * `SpreadsheetApp` (to write range data)
  * `Utilities` (to generate UUIDs)
* **Potential Errors:**
  * Writing errors if spreadsheet permissions are revoked.

---

## 13. `addCustomer`

### Description
Handles adding a brand-new client along with their first lead record. Validates that the phone number doesn't already exist before writing records.

* **Parameters:**
  * `data` (Object): Map containing client and lead details (`phone`, `name`, `city`, `enquiryDate`, `source`, `leadBy`).
* **Returns:** `Object` — `{ success: true, message: String }` or `{ success: false, message: String }` on duplicate match.
* **Google Services Used:**
  * `SpreadsheetApp` (to read/write client and lead databases)
  * `Utilities` (to generate UUIDs for both clients and leads)
* **Potential Errors:**
  * Phone check/concurrency errors.

---

## 14. `formatDateToDayMonthYear`

### Description
Utility function that converts standard JS/Apps Script `Date` objects to a reader-friendly format with ordinal suffixes (e.g. `"14th june 26"`).

* **Parameters:**
  * `dateObj` (Date/String): The date object to format.
* **Returns:** `String` — Formatted date string, or the raw input value if the type is not a Date.
* **Google Services Used:**
  * `Utilities` (to format date parts via `Utilities.formatDate()`)
  * `Session` (to read script execution timezones)
* **Potential Errors:** None (returns input string fallback on failure).

---

## 15. `getFilteredClients`

### Description
Retrieves a filtered list of clients from the `"Clients"` sheet. Aggregates the total number of leads linked to each client, filters out entries with zero leads, and ignores duplicate entries.

* **Parameters:**
  * `params` (Object): Filtering criteria (`name`, `phone`).
* **Returns:** `Object` — `{ clients: Array }` where each client item has `clientId`, `name`, `city`, `phone`, `regId`, `leadCount`.
* **Google Services Used:**
  * `SpreadsheetApp` (to read both Clients and Leads sheet dataset arrays)
* **Potential Errors:**
  * Performance bottleneck or timeout if datasets grow excessively large (since ranges are loaded entirely in memory).

---

## 16. `getFilteredLeads`

### Description
A unified function for querying leads from the `"Leads"` sheet. Merges client information and provides complex filters: text queries (Name, Phone, City), date bounds (Start/End date), status filters, today-only call listings, pagination (20 items per page), and sorting orders.

* **Parameters:**
  * `params` (Object): Filtering configurations (`name`, `phone`, `city`, `isTodayOnly`, `excludeConfirmed`, `requireStatus`, `status`, `regIdFilter`, `startDate`, `endDate`, `sortOrder`, `page`).
* **Returns:** `Object` — `{ leads: Array, totalPages: Number, currentPage: Number, totalItems: Number }`
* **Google Services Used:**
  * `SpreadsheetApp` (to read dataset ranges)
  * `Utilities` (to compare timestamps and format dates)
  * `Session` (to match localized timezones)
* **Potential Errors:**
  * JSON parser errors if lead conversation history logs contain corrupted formatting.

---

## 17. `getClientDetailsWithLeads`

### Description
Retrieves metadata for a single client (including historical edit records) alongside all their corresponding leads sorted newest first.

* **Parameters:**
  * `clientId` (String): The target unique client ID.
* **Returns:** `Object` — client details structure containing metadata, edit history list, and sorted leads array.
* **Google Services Used:**
  * `SpreadsheetApp` (to lookup row values in both tables)
* **Potential Errors:**
  * `Error("Client not found.")` if no client matches the ID.
  * JSON parse errors if logs or historical cells are corrupted.

---

## 18. `getUserDetails`

### Description
Retrieves details for a single lead using its row index. Also looks up and merges the associated client's metadata.

* **Parameters:**
  * `leadRowIndex` (Number): Row index of the lead.
* **Returns:** `Object` — Merged lead and client fields.
* **Google Services Used:**
  * `SpreadsheetApp` (to read target row ranges)
* **Potential Errors:**
  * `Error("Client not found for this lead.")` if the client record associated with the lead's Client ID is missing.
  * Index out-of-bounds error if the row is invalid.

---

## 19. `addConversation`

### Description
Performs updates on a lead's record (status, branch, contact again date, arrival status) and client's record (registration ID, name, city). Appends a new conversation entry (with operator name and timestamp) to the lead's history log. If critical client fields (name, city) changed, it generates and logs a change record in the client's edit history.

* **Parameters:**
  * `leadRowIndex` (Number): Row index of the target lead.
  * `text` (String): Content of the new note.
  * `agentName` (String): The operator/agent performing the update.
  * `status` (String): Status update value.
  * `branch` (String): Branch update value.
  * `contactAgainDate` (String): Contact-again date update value.
  * `registrationId` (String): Client registration ID update value.
  * `originalStatus` (String): Original lead status.
  * `name` (String): Edited client name.
  * `city` (String): Edited client city.
  * `originalName` (String): Original client name.
  * `originalCity` (String): Original client city.
  * `clientRowIndex` (Number): Row index of the client.
  * `arrivalStatus` (String): Arrival status update value.
* **Returns:** `Object` — Status return payload including updated values, updated conversations log, and edit history.
* **Google Services Used:**
  * `SpreadsheetApp` (to write cell values)
  * `Utilities` (to format current datetime strings)
  * `Session` (to obtain timezones)
* **Potential Errors:**
  * Concurrent update anomalies or write blockages.
  * JSON parse exceptions on historical logs.

---

## 20. `getDashboardStats`

### Description
Computes aggregated metrics and datasets for the CRM charts and statistics cards (total counts, rates, conversion statistics, source breakdowns, timeline metrics, operator counts, top 5 cities) based on a specified date range.

* **Parameters:**
  * `startDateStr` (String): Lower date limit (yyyy-MM-dd).
  * `endDateStr` (String): Upper date limit (yyyy-MM-dd).
* **Returns:** `Object` — Metrics map used by Chart.js on the dashboard.
* **Google Services Used:**
  * `SpreadsheetApp` (to read clients and leads datasets)
  * `Utilities` (to parse dates)
  * `Session` (to match localized timezones)
* **Potential Errors:**
  * Handled division-by-zero occurrences safely, but invalid date cells can distort computation calculations.
