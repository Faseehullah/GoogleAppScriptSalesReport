/**************************
 * Code.gs
 **************************/
/**************************
 * CONFIGURATION
 **************************/
const Config = {
  SPREADSHEET: {
    ID: "1mypSn1d0lZtM_Aww0EWb2Ovc__HRdFsvJ06HVhYUvls", 
    RANGES: {
      // Main data is appended here
      DATA: "MAIN!A2:S",
      // The credentials range now includes 4 columns: Username, HashedPassword, AvatarURL, Email , role
      CREDENTIALS: "DATASHEETS!T2:X",
      // One row per contact (ID, ADDED BY, CUSTOMER, PERSON NAME, DESIGNATION, CONTACT, TIME STAMP)
      CONTACTS: "CONTACTS!A2:G",
      // One row per new workload (ID, ADDED BY, CUSTOMER, Competitor, Competitor Model, Daily Work Load, Estimated Per Test Cost, TIME STAMP)
      WORKLOAD: "WORKLOAD!A2:H",
      // Dropdown Ranges (updated to FLAGS sheet)
      PRODUCT: "FLAGS!A2:A350",
      ANALYZER_MODEL: "FLAGS!B2:B350",
      FLAGS: "FLAGS!C2:D350", // [Flag, Category]
      REGION: "DATASHEETS!A2:A400",
      CITY: "DATASHEETS!B2:B400",
      CUSTOMERS: "DATASHEETS!C2:C400",
      DEPARTMENT: "DATASHEETS!E2:E50", // Added Department Range
      ERROR_LOGS: "Error Logs!A2:E",
      AUDIT_LOGS: "Audit Logs!A2:D"
    }
  },
  SECURITY: {
    SESSION_DURATION_SECONDS: 600, // 10 minutes
    MAX_LOGIN_ATTEMPTS: 5,
    LOGIN_TIMEOUT_SECONDS: 600 // 5 minutes
  },
  RATE_LIMITING: {
    EXPORT_LIMIT: 10, // Max exports per user
    EXPORT_WINDOW_SECONDS: 3600 // 1 hour
  }
};

/**************************
 * SESSION MANAGEMENT
 **************************/
class SessionManager {
  constructor() {
    this.cache = CacheService.getScriptCache();
    this.duration = Config.SECURITY.SESSION_DURATION_SECONDS;
  }

  /**
   * Starts a new session for a user.
   * @param {string} username - The username.
   * @returns {string} - The session token.
   */
  startSession(username) {
    const token = Utilities.getUuid();
    this.cache.put(`session_${username}`, token, this.duration);
    return token;
  }

  /**
   * Ends a user's session.
   * @param {string} username - The username.
   */
  endSession(username) {
    this.cache.remove(`session_${username}`);
  }

  /**
   * Validates a user's session.
   * @param {string} username - The username.
   * @param {string} token - The session token.
   * @returns {boolean} - True if valid, else false.
   */
  validateSession(username, token) {
    const cachedToken = this.cache.get(`session_${username}`);
    return cachedToken === token;
  }
}

const sessionManager = new SessionManager();

// Tracks login attempts per user
const userLoginAttempts = {};
// Tracks export counts per user
const userExportCounts = {};

/**************************
 * ERROR LOGGING
 **************************/
class ErrorLogger {
  /**
   * Logs an error to the "Error Logs" sheet and the Logger.
   * @param {string} functionName - Name of the function where the error occurred.
   * @param {Error} error - The error object.
   * @param {Object} additionalInfo - Any additional information.
   */
  static log(functionName, error, additionalInfo = {}) {
    const errorLog = [
      getCurrentPKTTimeStamp(),,
      functionName,
      error.message || error,
      error.stack || '',
      JSON.stringify(additionalInfo)
    ];
    
    try {
      const sheet = SpreadsheetApp.openById(Config.SPREADSHEET.ID).getSheetByName("Error Logs");
      if (sheet) {
        sheet.appendRow(errorLog);
      }
    } catch (sheetError) {
      // If logging to sheet fails, fallback to Logger
      Logger.log(`Failed to log to sheet: ${sheetError}`);
    }
    
    // Log to Logger for real-time debugging
    Logger.log(JSON.stringify(errorLog));
  }
}

/**************************
 * AUDIT LOGGING
 **************************/
class AuditLogger {
  /**
   * Logs audit events to the "Audit Logs" sheet.
   * @param {string} action - The action performed.
   * @param {string} username - The username of the actor.
   * @param {Object} details - Additional details about the action.
   */
  static log(action, username, details = {}) {
    try {
      const auditLog = [
        getCurrentPKTTimeStamp(),,
        username,
        action,
        JSON.stringify(details)
      ];
      
      const sheet = SpreadsheetApp.openById(Config.SPREADSHEET.ID).getSheetByName("Audit Logs");
      if (sheet) {
        sheet.appendRow(auditLog);
      }
    } catch (error) {
      ErrorLogger.log('AuditLogger.log', error, { action, username, details });
    }
  }
}

/**************************
 * PAGE ROUTING
 **************************/
// We only have a single page "Index.html" to serve by default
function doGet(e) {
  try {
    // Serve the main Index page
    return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Nextek Healthcare - Sales Report')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (error) {
    ErrorLogger.log('doGet', error, { parameters: e.parameter });
    throw error;
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**************************
 * LOGIN & AUTH FUNCTIONS
 **************************/
function checkLogin(username, password) {
  try {
    // Check for too many login attempts
    const attemptsInfo = userLoginAttempts[username] || { attempts: 0, lastAttempt: 0 };
    const currentTime = Math.floor(Date.now() / 1000); // in seconds

    // If user has too many attempts in the last X seconds
    if (attemptsInfo.attempts >= Config.SECURITY.MAX_LOGIN_ATTEMPTS &&
        (currentTime - attemptsInfo.lastAttempt) < Config.SECURITY.LOGIN_TIMEOUT_SECONDS) {
      return { 
        success: false, 
        message: "Too many failed attempts. Please try again later." 
      };
    }

    // Hash the incoming password
    const hashedPassword = hashPassword(password);

    // Read credentials from the sheet
    const data = readRecord(Config.SPREADSHEET.RANGES.CREDENTIALS);
    // data[i] -> [ username, hashedPwd, avatarUrl, email ]
    const foundUser = data.find(row => row[0] === username && row[1] === hashedPassword);
    
    if (foundUser) {
      // Determine role based on the value in the role column (assume "1" for Admin, "2" for User)
      let role = "";
      if (foundUser[4] === "1") {
        role = "Admin";
      } else if (foundUser[4] === "2") {
        role = "User";
      }
      
      // Successful login
      const sessionToken = sessionManager.startSession(username);
      let avatarLink = (foundUser[2] || "").trim();
      
      // Reset login attempts for the user
      delete userLoginAttempts[username];
      AuditLogger.log('Successful Login', username, { timestamp: getCurrentPKTTimeStamp(), });
      
      return {
        success: true,
        message: "Login successful",
        token: sessionToken,
        username: username,
        avatarUrl: avatarLink,
        role: role
      };
    } else {
      // Track failed login attempt
      if (!userLoginAttempts[username]) {
        userLoginAttempts[username] = { attempts: 1, lastAttempt: currentTime };
      } else {
        userLoginAttempts[username].attempts++;
        userLoginAttempts[username].lastAttempt = currentTime;
      }
      AuditLogger.log('Failed Login Attempt', username, { timestamp: getCurrentPKTTimeStamp(), });
      return { success: false, message: "Invalid username or password" };
    }
  } catch (error) {
    ErrorLogger.log('checkLogin', error, { username });
    AuditLogger.log('Login Error', username, { error: error.message });
    return { success: false, message: "An error occurred during login" };
  }
}

/** Hashes a password using SHA-256. */
function hashPassword(password) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password)
    .map(function(byte) {
      return ('0' + (byte & 0xFF).toString(16)).slice(-2);
    })
    .join('');
}
/**************************
 * "CHANGE PASSWORD"
 **************************/
function changePassword(username, oldPassword, newPassword, token) {
  try {
    // Validate the session first
    if (!sessionManager.validateSession(username, token)) {
      throw new Error('Invalid session');
    }
    
    // Hash the old password provided by the user
    const oldHashed = hashPassword(oldPassword);
    
    // Read all credentials
    const data = readRecord(Config.SPREADSHEET.RANGES.CREDENTIALS);
    // data[i] -> [ username, hashedPwd, avatarUrl, email, role ]
    const userIndex = data.findIndex(row => row[0] === username);
    if (userIndex === -1) {
      throw new Error('User not found');
    }
    
    // Verify that the current password is correct
    if (data[userIndex][1] !== oldHashed) {
      throw new Error('Current password is incorrect');
    }
    
    // Hash the new password
    const newHashed = hashPassword(newPassword);
    
    // Update the spreadsheet with the new hashed password.
    // Assuming that the credentials range starts at row 2, and that the hashed password is in the second column of your range.
    const sheet = SpreadsheetApp.openById(Config.SPREADSHEET.ID).getSheetByName("DATASHEETS");
    const rowToUpdate = userIndex + 2; // because row 1 is header, row 2 is first record.
    const hashedPasswordColumn = 21;  // Column U (if T=20, then U=21) in your spreadsheet.
    sheet.getRange(rowToUpdate, hashedPasswordColumn).setValue(newHashed);
    
    AuditLogger.log('Password Change', username, { timestamp: getCurrentPKTTimeStamp(), });
    return { success: true, message: "Password changed successfully." };
  } catch (error) {
    ErrorLogger.log('changePassword', error, { username });
    return { success: false, message: error.message };
  }
}

/**************************
 * "FORGOT PASSWORD"
 **************************/
function forgotPassword(emailOrUsername) {
  try {
    // 1) Read all credentials
    const data = readRecord(Config.SPREADSHEET.RANGES.CREDENTIALS);
    // data[i] -> [ username, hashedPassword, avatarUrl, userEmail ]

    // 2) Find matching row by username or email
    let rowIndex = -1;
    let foundRow = null;
    for (let i = 0; i < data.length; i++) {
      const [uname, hashed, avatar, userEmail] = data[i];
      if (uname === emailOrUsername || userEmail === emailOrUsername) {
        rowIndex = i;
        foundRow = data[i];
        break;
      }
    }
    if (rowIndex === -1) {
      return { success: false, message: "User not found." };
    }

    // 3) Generate a new random password
    const newPlainPassword = generateRandomPassword(8);  // length = 8
    const newHashed = hashPassword(newPlainPassword);

    // 4) Update the credential in the sheet
    const sheet = SpreadsheetApp.openById(Config.SPREADSHEET.ID).getSheetByName("DATASHEETS");
    const startRow = 2; // T2 is row 2
    const rowToUpdate = startRow + rowIndex; 
    const hashedPasswordColumn = 21; // column U (T=20, U=21, V=22, W=23)
    sheet.getRange(rowToUpdate, hashedPasswordColumn).setValue(newHashed);

    // 5) Send the email with the new password
    const userEmail = foundRow[3]; // the Email column
    if (!userEmail) {
      return { success: false, message: "No email on file for user." };
    }
    GmailApp.sendEmail(
      userEmail,
      "Password Reset",
      `Hello,

Your password has been reset. Here is your new temporary password: ${newPlainPassword}

Please login and change it as soon as possible.

Regards,
Your Company`
    );
    AuditLogger.log('Password Reset', foundRow[0], { email: userEmail });
    return { success: true, message: "A new password has been sent to your email." };
  } catch (error) {
    ErrorLogger.log('forgotPassword', error, { emailOrUsername });
    return { success: false, message: "An error occurred while resetting the password." };
  }
}

function generateRandomPassword(length) {
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789@#$!";
  let result = "";
  for (let i = 0; i < length; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return result;
}

/**************************
 * DATA ACCESS FUNCTIONS
 **************************/
function readRecord(range) {
  try {
    let result = Sheets.Spreadsheets.Values.get(Config.SPREADSHEET.ID, range);
    return result.values || [];
  } catch (error) {
    ErrorLogger.log('readRecord', error, { range });
    return [];
  }
}

function getDropdownDataCached(name) {
  try {
    const cacheKey = `dropdown_${name}`;
    const cache = CacheService.getScriptCache();
    const cached = cache.get(cacheKey);
    
    if (cached) {
      return JSON.parse(cached);
    }
    
    const rangeMap = {
      'Product': 'PRODUCT',
      'Department': 'DEPARTMENT', 
      // Add other mappings if needed
    };
    
    const range = Config.SPREADSHEET.RANGES[rangeMap[name]];
    if (!range) {
      throw new Error(`No range configured for dropdown: ${name}`);
    }
    
    const data = readRecord(range);
    // Remove duplicates
    const uniqueData = [...new Set(data.map(row => row[0] ? row[0].trim() : '').filter(Boolean))];
    const formattedData = uniqueData.map(item => [item]);
    
    cache.put(cacheKey, JSON.stringify(formattedData), 180); // 3 mins
    return formattedData;
  } catch (error) {
    ErrorLogger.log('getDropdownDataCached', error, { name });
    return [];
  }
}

function getUniqueRegions() {
  try {
    const data = readRecord(Config.SPREADSHEET.RANGES.REGION);
    const uniqueRegions = [...new Set(data.map(row => row[0] ? row[0].trim() : '').filter(Boolean))];
    return uniqueRegions.map(region => [region]);
  } catch (error) {
    ErrorLogger.log('getUniqueRegions', error, {});
    return [];
  }
}

function getCityByRegion(region) {
  try {
    const sheet = SpreadsheetApp.openById(Config.SPREADSHEET.ID).getSheetByName("DATASHEETS");
    const regionData = sheet.getRange("A2:A400").getValues();
    const cityData = sheet.getRange("B2:B400").getValues();
    
    const uniqueCities = [...new Set(
      regionData
        .map((row, idx) => (row[0] === region ? cityData[idx][0] : null))
        .filter(city => city && city.trim() !== '')
    )];
    
    return uniqueCities.map(city => [city]);
  } catch (error) {
    ErrorLogger.log('getCityByRegion', error, { region });
    return [];
  }
}

function getCustomersByCity(city) {
  try {
    const sheet = SpreadsheetApp.openById(Config.SPREADSHEET.ID).getSheetByName("DATASHEETS");
    const cityData = sheet.getRange("B2:B400").getValues();
    const customerData = sheet.getRange("C2:C400").getValues();
    
    const uniqueCustomers = [...new Set(
      cityData
        .map((row, idx) => (row[0] === city ? customerData[idx][0] : null))
        .filter(cust => cust && cust.trim() !== '')
    )];
    
    return uniqueCustomers.map(customer => [customer]);
  } catch (error) {
    ErrorLogger.log('getCustomersByCity', error, { city });
    return [];
  }
}

function getUniqueProducts() {
  try {
    const data = readRecord(Config.SPREADSHEET.RANGES.PRODUCT);
    const uniqueProducts = [...new Set(data.map(row => row[0] ? row[0].trim() : '').filter(Boolean))];
    return uniqueProducts.map(product => [product]);
  } catch (error) {
    ErrorLogger.log('getUniqueProducts', error, {});
    return [];
  }
}

function getAnalyzersByProduct(product) {
  try {
    const sheet = SpreadsheetApp.openById(Config.SPREADSHEET.ID).getSheetByName("FLAGS");
    const productData = sheet.getRange("A2:A350").getValues();
    const analyzerData = sheet.getRange("B2:B350").getValues();
    
    const analyzers = [];
    for (let i = 0; i < productData.length; i++) {
      if (productData[i][0] && analyzerData[i][0] &&
          productData[i][0].trim().toUpperCase() === product.trim().toUpperCase()) {
        analyzers.push(analyzerData[i][0].trim());
      }
    }

    const uniqueAnalyzers = [...new Set(analyzers)];
    return uniqueAnalyzers.map(a => [a]);
  } catch (error) {
    ErrorLogger.log('getAnalyzersByProduct', error, { product });
    return [];
  }
}

function getFlagsByProductAndAnalyzer(product, analyzerModel) {
  try {
    const sheet = SpreadsheetApp.openById(Config.SPREADSHEET.ID).getSheetByName("FLAGS");
    const flagData = sheet.getRange("C2:D350").getValues(); // [Flag, Category]
    const productData = sheet.getRange("A2:A350").getValues();
    const analyzerData = sheet.getRange("B2:B350").getValues();

    const flags = [];
    for (let i = 0; i < productData.length; i++) {
      if (productData[i][0] && analyzerData[i][0] &&
          productData[i][0].trim().toUpperCase() === product.trim().toUpperCase() &&
          analyzerData[i][0].trim().toUpperCase() === analyzerModel.trim().toUpperCase()) {
        const flag = flagData[i][0] ? flagData[i][0].trim() : '';
        if (flag) {
          flags.push(flag);
        }
      }
    }

    const uniqueFlags = [...new Set(flags)];
    return uniqueFlags.map(f => [f]);
  } catch (error) {
    ErrorLogger.log('getFlagsByProductAndAnalyzer', error, { product, analyzerModel });
    return [];
  }
}

function getCategoryByProductAnalyzerFlag(product, analyzerModel, flag) {
  try {
    const sheet = SpreadsheetApp.openById(Config.SPREADSHEET.ID).getSheetByName("FLAGS");
    const data = sheet.getRange("A2:D350").getValues(); // [Product, AnalyzerModel, Flag, Category]
    
    const matchRow = data.find(row => 
      row[0] && row[0].trim().toUpperCase() === product.trim().toUpperCase() &&
      row[1] && row[1].trim().toUpperCase() === analyzerModel.trim().toUpperCase() &&
      row[2] && row[2].trim().toUpperCase() === flag.trim().toUpperCase()
    );
    return matchRow ? (matchRow[3].trim() || "") : "";
  } catch (error) {
    ErrorLogger.log("getCategoryByProductAnalyzerFlag", error, { product, analyzerModel, flag });
    return "";
  }
}

/**************************
 * WORKLOAD MANAGEMENT
 **************************/
function createMultipleWorkloadRecords(id, customer, workloads, username) {
  try {
    if (!Array.isArray(workloads) || workloads.length === 0) return;

    const sheet = SpreadsheetApp.openById(Config.SPREADSHEET.ID).getSheetByName("WORKLOAD");
    // Columns: A: ID, B: ADDED BY, C: CUSTOMER, D: Competitor, E: CompetitorModel, F: DailyWorkload, G: PerTestCost, H:Timestamp

    // Filter out workload items that are entirely empty:
const validWorkloads = workloads.filter(wl => {
  return (wl.competitor && wl.competitor.trim()) ||
         (wl.competitorModel && wl.competitorModel.trim()) ||
         (wl.dailyWorkload && wl.dailyWorkload.toString().trim()) ||
         (wl.perTestCost && wl.perTestCost.toString().trim());
});

if (!Array.isArray(validWorkloads) || validWorkloads.length === 0) return;

const rows = validWorkloads.map(wl => [
  id,
  username,
  customer || "",
  wl.competitor || "",
  wl.competitorModel || "",
  wl.dailyWorkload || "",
  wl.perTestCost || "",
  getCurrentPKTTimeStamp(),
]);


    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 8).setValues(rows);
  } catch (error) {
    ErrorLogger.log('createMultipleWorkloadRecords', error, { id, workloads });
    throw new Error('Failed to save workload records.');
  }
}

/**************************
 * CONTACT PERSON MANAGEMENT
 **************************/
function formatContactPersons(contactPersons) {
  if (!Array.isArray(contactPersons)) return '';
  
  return contactPersons.map((contact, idx) => {
    const letter = String.fromCharCode(65 + idx); // A, B, C...
    return `${letter}/${contact.name}/${contact.post}/${contact.number || 'N/A'}`;
  }).join(';');
}

function createContactsRecords(id, customerName, contactPersons, username) {
  try {
    if (!Array.isArray(contactPersons) || contactPersons.length === 0) return;

    const sheet = SpreadsheetApp.openById(Config.SPREADSHEET.ID).getSheetByName("CONTACTS");
    // Columns: A:ID, B:AddedBy, C:Customer, D:Name, E:Designation, F:Contact#, G:Timestamp

    const existingData = readRecord(Config.SPREADSHEET.RANGES.CONTACTS);
    const rowsToAppend = [];

    contactPersons.forEach(person => {
      // Check for duplication
      const duplicate = existingData.find(row => 
        row[3].toUpperCase() === person.name.toUpperCase() &&
        row[4].toUpperCase() === person.post.toUpperCase()
      );

      if (!duplicate) {
        // If no duplicate, append as new record
        rowsToAppend.push([
          id,
          username,
          customerName,
          person.name.toUpperCase(),
          person.post.toUpperCase(),
          person.number,
          getCurrentPKTTimeStamp(),
        ]);
      } else {
        // Update existing record's contact number + timestamp
        const rowIndex = existingData.indexOf(duplicate) + 2; 
        sheet.getRange(rowIndex, 6).setValue(person.number);
        sheet.getRange(rowIndex, 7).setValue(getCurrentPKTTimeStamp(),);

        AuditLogger.log('Updated Existing Contact', username, { 
          id, username, customerName, 
          contactName: person.name, 
          contactPost: person.post 
        });
      }
    });

    if (rowsToAppend.length > 0) {
      const valueRange = Sheets.newValueRange();
      valueRange.values = rowsToAppend;
      Sheets.Spreadsheets.Values.append(
        valueRange,
        Config.SPREADSHEET.ID, 
        Config.SPREADSHEET.RANGES.CONTACTS,
        { valueInputOption: "RAW" }
      );
      AuditLogger.log('Added New Contacts', username, { 
        customerName, 
        contactsAdded: rowsToAppend.map(row => row[3]) 
      });
    }

  } catch (error) {
    ErrorLogger.log('createContactsRecords', error, { id, customerName, contactPersons, username });
    throw new Error('Failed to save contact records.');
  }
}

function getExistingContacts(customerId) {
  try {
    const data = readRecord(Config.SPREADSHEET.RANGES.CONTACTS);
    
    // Map of name+post => most recent entry
    const uniqueContacts = new Map();
    
    data
      .filter(row => row[2] === customerId) // row[2] is Customer
      .forEach(row => {
        const key = `${row[3]}_${row[4]}`;
        if (!uniqueContacts.has(key) || new Date(row[6]) > new Date(uniqueContacts.get(key).timestamp)) {
          uniqueContacts.set(key, {
            name: row[3],
            post: row[4],
            number: row[5],
            timestamp: row[6]
          });
        }
      });

    return Array.from(uniqueContacts.values())
      .sort((a, b) => a.name.localeCompare(b.name))
      .map(({name, post, number}) => ({name, post, number}));
      
  } catch (error) {
    ErrorLogger.log('getExistingContacts', error, { customerId });
    return [];
  }
}

/**************************
 * FORM PROCESSING
 **************************/
function validateFormData(formObject) {
  const errors = [];
  
  const requiredFields = [
    'dateOfVisit','visitType','Region','City','Customer',
    'Department','Product','AnalyzerModel','Flags','VISIT_DESCRIPTION',
    'contactPersons'
  ];
  
  for (const field of requiredFields) {
    if (!formObject[field]) {
      errors.push(`${field} is required`);
    }
  }

  // Basic check on dateOfVisit
  if (formObject.dateOfVisit) {
    const visitDate = new Date(formObject.dateOfVisit);
    if (isNaN(visitDate.getTime()) || visitDate > new Date()) {
      errors.push('Invalid visit date');
    }
  }

  // Limit visit description length
  if (formObject.VISIT_DESCRIPTION && formObject.VISIT_DESCRIPTION.length > 1000) {
    errors.push('Visit description too long (max 1000 characters)');
  }
  
  return errors;
}

function processForm(formObject, currentUser, token) {
  try {
    // 1) Validate session
    if (!sessionManager.validateSession(currentUser, token)) {
      throw new Error('Invalid session');
    }
    
    // 2) Basic validation
    const errors = validateFormData(formObject);
    if (errors.length > 0) {
      return `Validation errors: ${errors.join(', ')}`;
    }

    // 3) If user selected "**Add New Customer**"
    if (formObject.Customer === "**Add New Customer**" && formObject.newCustomer) {
      addNewCustomerWithCity(
        formObject.Region, 
        formObject.newCustomerCity, 
        formObject.newCustomer, 
        currentUser, 
        token
      );
      formObject.Customer = formObject.newCustomer;
    }

    // 4) Validate each contact person (Server-side, optional)
    //    (You already do client-side checks, but can add more here if needed.)

    // 5) Format contact persons into single string for MAIN sheet
    const formattedContacts = formatContactPersons(formObject.contactPersons);

    // 6) Next visit date
    const nextVisitIso = formObject.nextVisit ? new Date(formObject.nextVisit).toISOString() : '';

    // 7) Unique ID
    const uniqueId = Utilities.getUuid();

    // 8) Append to MAIN sheet
    const values = [[
      uniqueId,
      formObject.dateOfVisit,
      formObject.visitType,
      formObject.Region,
      formObject.City,
      formObject.Customer,
      formObject.Department,
      currentUser,
      formObject.Product,
      formObject.AnalyzerModel,
      formObject.Flags,
      formObject.Category,
      formObject.VISIT_DESCRIPTION,
      nextVisitIso,
      formatWorkloadData(formObject.workloads),
      formattedContacts,
      formObject.IMAGE || "",
      getCurrentPKTTimeStamp(),
      ""
    ]];
    createRecord(values);

    // 9) Create contact records in CONTACTS sheet
    createContactsRecords(uniqueId, formObject.Customer, formObject.contactPersons, currentUser);

    // 10) Create workload records
    createMultipleWorkloadRecords(uniqueId, formObject.Customer, formObject.workloads, currentUser);

    AuditLogger.log('Form Submission', currentUser, { id: uniqueId, customer: formObject.Customer });
    return `Data successfully submitted by Sales Person: ${currentUser}`;
  } catch (error) {
    ErrorLogger.log('processForm', error, { formObject, currentUser });
    AuditLogger.log('Form Submission Failure', currentUser, { error: error.message });
    return "Error: " + error.message;
  }
}

function createRecord(values) {
  try {
    const valueRange = Sheets.newValueRange();
    valueRange.values = values;
    Sheets.Spreadsheets.Values.append(
      valueRange, 
      Config.SPREADSHEET.ID, 
      Config.SPREADSHEET.RANGES.DATA,
      { valueInputOption: "RAW" }
    );
  } catch (error) {
    ErrorLogger.log('createRecord', error, { values });
    throw new Error('Failed to save record.');
  }
}

/**************************
 * ADD NEW CUSTOMER
 **************************/
function addNewCustomerWithCity(region, city, customerName, username, token) {
  try {
    if (!sessionManager.validateSession(username, token)) {
      throw new Error('Invalid session');
    }

    region = sanitizeInput(region);
    city = sanitizeInput(city);
    customerName = sanitizeInput(customerName);

    const sheet = SpreadsheetApp.openById(Config.SPREADSHEET.ID).getSheetByName("DATASHEETS");
    
    let row = 2;
    while(sheet.getRange(row, 3).getValue()) row++;
    // Write to columns A, B, C
    sheet.getRange(row, 1, 1, 3).setValues([[region, city,customerName]]);

    AuditLogger.log('ADD_CUSTOMER', username, { region, city, customerName, row });
    return "New customer added successfully!";
  } catch (error) {
    ErrorLogger.log('addNewCustomerWithCity', error, { region, city, customerName, username });
    throw new Error('Failed to add new customer: ' + error.message);
  }
}

/**************************
 * EXPORT - Existing All Data
 **************************/
function getSalesPersonData(salesPersonName, role) {
  try {
    const data = readRecord(Config.SPREADSHEET.RANGES.DATA);
    if (role === "Admin") {
      return data; // Admin sees all data
    } else {
      // Regular user sees only their own data (column 8 is Sales Person)
      return data.filter(row => row[7] === salesPersonName);
    }
  } catch (error) {
    ErrorLogger.log('getSalesPersonData', error, { salesPersonName, role });
    return [];
  }
}
function exportSalesPersonDataAsCSV(salesPersonName, username, token) {
  try {
    if (!sessionManager.validateSession(username, token)) {
      throw new Error('Invalid session');
    }
    // Rate limit
    const nowSec = Math.floor(Date.now() / 1000);
    if (!userExportCounts[username]) {
      userExportCounts[username] = { count: 0, windowStart: nowSec };
    }
    const uExport = userExportCounts[username];
    if ((nowSec - uExport.windowStart) > Config.RATE_LIMITING.EXPORT_WINDOW_SECONDS) {
      uExport.count = 0;
      uExport.windowStart = nowSec;
    }
    if (uExport.count >= Config.RATE_LIMITING.EXPORT_LIMIT) {
      throw new Error('Export limit reached. Try again later.');
    }
    uExport.count++;
    userExportCounts[username] = uExport;

    const spData = getSalesPersonData(salesPersonName);
    if (spData.length === 0) {
      throw new Error('No data available for export.');
    }

    // Build CSV
    const headers = [
      "ID","Date of Visit","Visit Type","Region","City",
      "Customer","Department","Sales Person","Product",
      "Analyzer Model","Flags","Visit Description","Image URL",
      "Submission Time","Contact Persons","Next Visit Date"
    ];
    let csvContent = headers.join(",") + "\n";
    
    spData.forEach(function(row) {
      const formattedRow = row.map(item => {
        const escapedItem = (`${item}`).replace(/"/g, '""');
        return /[",]/.test(escapedItem) ? `"${escapedItem}"` : escapedItem;
      }).join(",");
      csvContent += formattedRow + "\n";
    });

    const file = DriveApp.createFile(
      Utilities.newBlob(csvContent, 'text/csv', `${salesPersonName}_Data.csv`)
    );
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const url = file.getDownloadUrl();
    ScriptApp.newTrigger('deleteTempExportFile')
      .timeBased().after(60*60*1000).create();
    return url;
  } catch (error) {
    ErrorLogger.log('exportSalesPersonDataAsCSV', error, { salesPersonName, username });
    throw new Error(error.message);
  }
}

/**************************
 * EXPORT - Date Range Filter
 **************************/
/**
 * Return the rows for user in [fromDate, toDate].
 * fromDate/toDate are in YYYY-MM-DD format from <input type="date">.
 * row[1] is your "Date of Visit".
 */
// Updated getSalesPersonDataWithinRange
function getSalesPersonDataWithinRange(salesPersonName, role, fromDate, toDate) {
  try {
    const allData = getSalesPersonData(salesPersonName, role); 
    if (!fromDate && !toDate) {
      //if no range, then return all data
      return allData;
    }
    const fromDt = fromDate ? new Date(fromDate + 'T00:00:00') : null;
    const toDt   = toDate   ? new Date(toDate + 'T23:59:59') : null;
    return allData.filter(row => {
      const dateStr = row[1];  // Column B of Visit Dates
      if (!dateStr) return false;
      const vDate = new Date(dateStr);
      if (isNaN(vDate)) return false;
      if (fromDt && vDate < fromDt) return false;
      if (toDt && vDate > toDt) return false;
      return true;
    });
  } catch (error) {
    ErrorLogger.log('getSalesPersonDataWithinRange', error, { salesPersonName, role, fromDate, toDate });
    throw new Error(error.message);
  }
}

function exportSalesPersonDataWithinRangeAsCSV(salesPersonName, fromDate, toDate, username, token,role) {
  try {
    // Validate session
    if (!sessionManager.validateSession(username, token)) {
      throw new Error('Invalid session');
    }
    
    // Rate limit
    const nowSec = Math.floor(Date.now() / 1000);
    if (!userExportCounts[username]) {
      userExportCounts[username] = { count: 0, windowStart: nowSec };
    }
    const uExport = userExportCounts[username];
    if ((nowSec - uExport.windowStart) > Config.RATE_LIMITING.EXPORT_WINDOW_SECONDS) {
      uExport.count = 0;
      uExport.windowStart = nowSec;
    }
    if (uExport.count >= Config.RATE_LIMITING.EXPORT_LIMIT) {
      throw new Error('Export limit reached. Try again later.');
    }
    uExport.count++;
    userExportCounts[username] = uExport;

    // Filter data using the current user's role
    const rows = getSalesPersonDataWithinRange(salesPersonName, role, fromDate, toDate);
    if (!rows || rows.length === 0) {
      throw new Error('No data found for the specified range.');
    }

    // Build CSV
    const headers = [
      "ID", "Date of Visit", "Visit Type", "Region", "City",
      "Customer", "Department", "Sales Person", "Product",
      "Analyzer Model", "Flags", "Category", "Visit Description", "Next Visit", "Workload",
      "CONTACT DETAILS", "Image Url", "Time Stamp"
    ];
    let csvContent = headers.join(",") + "\n";

    rows.forEach(function(row) {
      const escapedRow = row.map(item => {
        const val = (item || '').replace(/"/g, '""');
        return /[",]/.test(val) ? `"${val}"` : val;
      }).join(",");
      csvContent += escapedRow + "\n";
    });

    // Create file
    const fileName = `${salesPersonName}_Data_${fromDate || 'ALL'}_${toDate || 'ALL'}.csv`;
    const file = DriveApp.createFile(
      Utilities.newBlob(csvContent, 'text/csv', fileName)
    );
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const dlUrl = file.getDownloadUrl();
    ScriptApp.newTrigger('deleteTempExportFile')
      .timeBased().after(60*60*1000).create();
    return dlUrl;
  } catch (error) {
    ErrorLogger.log('exportSalesPersonDataWithinRangeAsCSV', error, { salesPersonName, fromDate, toDate, username });
    throw new Error(error.message);
  }
}

/**************************
 * FILE DELETION TRIGGER
 **************************/
function deleteTempExportFile() {
  try {
    const files = DriveApp.getFilesByName(/_Data\.csv$/);
    while (files.hasNext()) {
      files.next().setTrashed(true);
    }
  } catch (error) {
    ErrorLogger.log('deleteTempExportFile', error);
  }
}

/**************************
 * INPUT SANITIZATION
 **************************/
function sanitizeInput(input) {
  if (typeof input !== 'string') return input;
  return input
    .trim()
    .replace(/[<>]/g, '')    // remove < >
    .replace(/[&]/g, 'and')  // & => 'and'
    .replace(/['"]/g, '')    // remove quotes
    .substring(0, 1000);     // limit length
}

/**************************
 * Helper Functions
 **************************/
function getCurrentPKTTimeStamp() {
  return Utilities.formatDate(new Date(), "Asia/Karachi", "MM/dd/yyyy, hh:mm:ss a");
}

/**
 * Formats the workload array to a single string in MAIN sheet.
 */
function formatWorkloadData(workloads) {
  if (!Array.isArray(workloads) || workloads.length === 0) return '';
  return workloads.map((wl, idx) => {
    const letter = String.fromCharCode(65 + idx);
    return `${letter}/${sanitizeInput(wl.competitor)}/${sanitizeInput(wl.competitorModel)}/${sanitizeInput(wl.dailyWorkload)}/${sanitizeInput(wl.perTestCost)}`;
  }).join('; ');
}
