/**************************
 * CONFIGURATION
 **************************/
const Config = {
  SPREADSHEET: {
    ID: "1mypSn1d0lZtM_Aww0EWb2Ovc__HRdFsvJ06HVhYUvls",
    RANGES: {
      // Main data is appended here
      DATA: "MAIN!A2:U2",   
      // The credentials range now includes 4 columns: Username, HashedPassword, AvatarURL, Email
      CREDENTIALS: "DATASHEETS!T2:W32",
      // One row per contact (ID, CUSTOMER, PERSON NAME, DESIGNATION, CONTACT)
      CONTACTS: "CONTACTS!A2:G",
      // One row per new workload (ID, Competitor, Competitor Model, Daily Work Load, Estimated Per Test Cost)
      WORKLOAD: "WORKLOAD!A2:G2",
      // Dropdown Ranges (unchanged)
      REGION: "DATASHEETS!L2:L6",
      CITY: "DATASHEETS!J2:J400",
      DEPARTMENT: "DATASHEETS!E2:E30",
      SALESPERSON: "DATASHEETS!G2:G55",
      PRODUCT: "DATASHEETS!O2:O25",
      FLAGS: "FLAGS!A2:C50",
      ANALYZER_MODEL: "DATASHEETS!Q2:R30",
      CUSTOMERS: "DATASHEETS!A2:C400",
      CITY_REGION_MAP: "DATASHEETS!I2:J400",
      ERROR_LOGS: "Error Logs!A2:E2", // For error logging
      AUDIT_LOGS: "Audit Logs!A2:D2"
    }
  },
  SECURITY: {
    SESSION_DURATION_SECONDS: 600, // 10 minutes
    MAX_LOGIN_ATTEMPTS: 5,
    LOGIN_TIMEOUT_SECONDS: 300 // 5 minutes
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
    // Convert to Pakistan time (+5:00)
    
   
    const errorLog = [
      getPakistanDateTime().toISOString(),
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
 * PAGE ROUTING
 **************************/
function doGet(e) {
  try {
    let page = e.parameter.page;
    let template;
    
    if (page === 'extraction') {
      template = HtmlService.createTemplateFromFile('extraction');
    } else if (page === 'extractionreport') {
      template = HtmlService.createTemplateFromFile('extractionreport');
    } else {
      // Default: show the main Index
      template = HtmlService.createTemplateFromFile('Index');
    }

    return template.evaluate()
      .setTitle('Sales Report')
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
      // Successful login
      const sessionToken = sessionManager.startSession(username);
      let avatarLink = (foundUser[2] || "").trim();
      
      // Reset login attempts
      if (userLoginAttempts[username]) {
        delete userLoginAttempts[username];
      }
      AuditLogger.log('Successful Login', username, { timestamp: getPakistanDateTime().toISOString() });
      return {
        success: true,
        message: "Login successful",
        token: sessionToken,
        username: username,
        avatarUrl: avatarLink
      };
    } else {
      // Track failed attempts
      if (!userLoginAttempts[username]) {
        userLoginAttempts[username] = { attempts: 1, lastAttempt: currentTime };
      } else {
        userLoginAttempts[username].attempts += 1;
        userLoginAttempts[username].lastAttempt = currentTime;
      }
      AuditLogger.log('Failed Login Attempt', username, { timestamp: getPakistanDateTime().toISOString() });
      return { success: false, message: "Invalid username or password" };
    }
  } catch (error) {
    ErrorLogger.log('checkLogin', error, { username });
    AuditLogger.log('Login Error', username, { error: error.message });AuditLogger.log('Login Error', username, { error: error.message });
    return { success: false, message: "An error occurred during login" };
  }
}

/**
 * Hashes a password using SHA-256.
 * @param {string} password - The plain text password.
 * @returns {string} - The hashed password in hexadecimal format.
 */
function hashPassword(password) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password)
    .map(function(byte) {
      return ('0' + (byte & 0xFF).toString(16)).slice(-2);
    })
    .join('');
}


/**************************
 * "FORGOT PASSWORD" FEATURE
 * - Generates a new password, updates the sheet, and emails the user.
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
    //    The range starts at T2 => row 2 in DATASHEETS, so rowIndex=0 => T2, rowIndex=1 => T3, etc.
    //    T=column 20, U=21, V=22, W=23. The hashed password is in column U if we follow:
    //        T -> username
    //        U -> hashed password
    //        V -> avatar url
    //        W -> email
    //    Adjust if your actual columns differ.
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
    // Use GmailApp to send
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

/** Generates a random password of given length with letters, digits, and special chars */
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
      'Region': 'REGION',
      'City': 'CITY',
      'Department': 'DEPARTMENT',
      'SalesPerson': 'SALESPERSON',
      'Product': 'PRODUCT',
      'Flags': 'FLAGS',
      'AnalyzerModel': 'ANALYZER_MODEL',
      'Customer': 'CUSTOMERS'
    };
    
    const range = Config.SPREADSHEET.RANGES[rangeMap[name]];
    if (!range) {
      throw new Error(`No range configured for dropdown: ${name}`);
    }
    
    const data = readRecord(range);
    cache.put(cacheKey, JSON.stringify(data), 180); // Cache for 3 mins
    return data;
  } catch (error) {
    ErrorLogger.log('getDropdownDataCached', error, { name });
    return [];
  }
}

function getCityByRegion(region) {
  try {
    const data = readRecord(Config.SPREADSHEET.RANGES.CUSTOMERS);
    // Get unique cities for the region
    const uniqueCities = [...new Set(
      data
        .filter(row => row[0] === region)  // Filter by region
        .map(row => row[1])                // Get city column
        .filter(Boolean)                    // Remove empty values
    )];
    
    // Return in required format [[city1], [city2], ...]
    return uniqueCities.map(city => [city]);
  } catch (error) {
    ErrorLogger.log('getCityByRegion', error, { region });
    return [];
  }
}

function getCustomersByCity(city) {
  try {
    const data = readRecord(Config.SPREADSHEET.RANGES.CUSTOMERS);
    // Get unique customers for the city
    const uniqueCustomers = [...new Set(
      data
        .filter(row => row[1] === city)     // Filter by city
        .map(row => row[2])                 // Get customer name
        .filter(Boolean)                     // Remove empty values
    )];
    
    // Return in required format [[customer1], [customer2], ...]
    return uniqueCustomers.map(customer => [customer]);
  } catch (error) {
    ErrorLogger.log('getCustomersByCity', error, { city });
    return [];
  }
}



function getAnalyzersByProduct(product) {
  try {
    const data = readRecord(Config.SPREADSHEET.RANGES.ANALYZER_MODEL);
    return data
      .filter(row => row[0] === product)
      .map(row => [row[1]])
      .filter(analyzer => analyzer[0]);
  } catch (error) {
    ErrorLogger.log('getAnalyzersByProduct', error, { product });
    return [];
  }
}

function getFlagsByProduct(product) {
  try {
    // Let’s say we store them in "FLAGS!A2:C2255" => A=Product, B=Flag , C=Catergory
    const data = readRecord("FLAGS!A2:C255");
    // Filter rows where row[0] == product
    const matching = data.filter(row => row[0] && row[0].toUpperCase() === product.toUpperCase());
    // Return array-of-arrays of the flags
    const flagsList = matching.map(row => [row[1]]);
    return flagsList;
  } catch (error) {
    ErrorLogger.log('getFlagsByProduct', error, { product });
    return [];
  }
}

/**
 * Returns the single category (column C) for a given product and flag.
 * @param {string} product
 * @param {string} flag
 * @return {string} The matching category, or "" if not found.
 */
function getCategoryByFlag(product, flag) {
  try {
    const data = readRecord("FLAGS!A2:C255"); // [ Product, Flag, Category ]
    // Find the row that matches the product & flag
    const matchRow = data.find(row => 
      row[0] && row[0].toUpperCase() === product.toUpperCase() &&
      row[1] && row[1].toUpperCase() === flag.toUpperCase()
    );
    if (!matchRow) {
      // No match
      return "";
    }
    // matchRow[2] is the Category
    return matchRow[2] || "";
  } catch (error) {
    ErrorLogger.log("getCategoryByFlag", error, { product, flag });
    return "";
  }
}


/**************************
 * WORKLOAD MANAGEMENT
 **************************/
/**
 * Creates a workload record in the WORKLOAD sheet.
 * If all workload fields are empty, it does nothing.
 * @param {string} id - The unique ID for the record.
 * @param {string} customer - The customer's name.
 * @param {Array} workloads - An array of workload objects.
 */
function createMultipleWorkloadRecords(id,customer, workloads) {
  try {
    if (!Array.isArray(workloads) || workloads.length === 0) return;

    const sheet = SpreadsheetApp.openById(Config.SPREADSHEET.ID).getSheetByName("WORKLOAD");
    // Suppose columns = A: ID, B:Customer , C:Competitor, D:CompetitorModel,
    //                  E:DailyWorkload, F:PerTestCost, G:Timestamp
    

    const rows = workloads.map(wl => [
      id,
      customer || "",
      wl.competitor || "",
      wl.competitorModel || "",
      wl.dailyWorkload || "",
      wl.perTestCost || "",
      getPakistanDateTime().toISOString() // Timestamp
    ]);

    sheet.getRange(sheet.getLastRow()+1, 1, rows.length, 7).setValues(rows);
  } catch (error) {
    ErrorLogger.log('createMultipleWorkloadRecords', error, { id, workloads });
    throw new Error('Failed to save workload records.');
  }
}

/**************************
 * CONTACT PERSON MANAGEMENT 
 **************************/
function formatContactPersons(contactPersons) {
  // (OPTIONAL) If you still want to store a summary in the main record
  if (!Array.isArray(contactPersons)) return '';
  
  return contactPersons.map((contact, index) => {
    const letter = String.fromCharCode(65 + index); // A, B, C...
    return `${letter}/${contact.name}/${contact.post}/${contact.number || 'N/A'}`;
  }).join(';');
}

function parseContactPersons(contactString) {
  if (!contactString) return [];
  
  return contactString.split(';').map(contact => {
    const [letter, name, post, number] = contact.split('/');
    return {
      letter,
      name,
      post,
      number: number === 'N/A' ? '' : number
    };
  });
}

function validateContactPerson(contact) {
  const errors = [];
  if (!contact.name || contact.name.trim() === '') {
    errors.push('Contact person name is required');
  }
  if (!contact.post || contact.post.trim() === '') {
    errors.push('Contact person post is required'); 
  }
  // Only validate number if provided
  if (contact.number && contact.number.trim() !== '' && 
      !/^\+?\d{11,13}$/.test(contact.number.replace(/\s+/g, ''))) {
    errors.push('Contact number must be 11-13 digits, optionally starting with +');
  }
  return errors;
}

function createContactsRecords(id, customerName, contactPersons, username) {
  try {
    if (!Array.isArray(contactPersons) || contactPersons.length === 0) return;

    const sheet = SpreadsheetApp.openById(Config.SPREADSHEET.ID).getSheetByName("CONTACTS");
    // Columns: A: ID, B: Customer, C: Person Name, D: Designation, E: Contact Number, F: Timestamp

    const existingData = readRecord(Config.SPREADSHEET.RANGES.CONTACTS);

    const rowsToAppend = [];

    contactPersons.forEach(person => {
      // Check for duplication based on name and post
      const duplicate = existingData.find(row => 
        row[2].toUpperCase() === person.name.toUpperCase() &&
        row[3].toUpperCase() === person.post.toUpperCase()
      );

      if (!duplicate) {
        // If no duplicate, append as new record
        rowsToAppend.push([
          id, 
          customerName,
          person.name.toUpperCase(),
          person.post.toUpperCase(), 
          person.number,
          getPakistanDateTime().toISOString()
        ]);
      } else {
        // Optionally, update existing record or skip
        // For example, update contact number and timestamp
        const rowIndex = existingData.indexOf(duplicate) + 2; // Sheets are 1-indexed
        sheet.getRange(rowIndex, 5).setValue(person.number);
        sheet.getRange(rowIndex, 6).setValue(getPakistanDateTime().toISOString());

        // Log the update
        AuditLogger.log('Updated Existing Contact', username, { 
          id, 
          username,
          customerName, 
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

      // Log the additions
      AuditLogger.log('Added New Contacts', username, { 
        customerName, 
        contactsAdded: rowsToAppend.map(row => row[2]) 
      });
    }

  } catch (error) {
    ErrorLogger.log('createContactsRecords', error, { id, customerName, contactPersons });
    throw new Error('Failed to save contact records.');
  }
}

function getExistingContacts(customerId) {
    try {
        const data = readRecord(Config.SPREADSHEET.RANGES.CONTACTS);
        
        // Create a Map to store unique contacts using name+post as key
        const uniqueContacts = new Map();
        
        data
            .filter(row => row[2] === customerId) // Corrected index for Customer
            .forEach(row => {
                const key = `${row[3]}_${row[4]}`; // name_post as unique key
                // Keep most recent entry based on timestamp
                if (!uniqueContacts.has(key) || 
                    new Date(row[6]) > new Date(uniqueContacts.get(key).timestamp)) {
                    uniqueContacts.set(key, {
                        name: row[3],
                        post: row[4],
                        number: row[5],
                        timestamp: row[6]
                    });
                }
            });

        // Convert Map values to array and sort by name
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
    'dateOfVisit',
    'visitType',
    'Region',
    'City',
    'Customer',
    'Department',
    'SalesPerson',
    'Product',
    'AnalyzerModel',
    'Flags',
    'VISIT_DESCRIPTION',
    'contactPersons'  // Must have at least one contact
  ];
  
  for (const field of requiredFields) {
    if (!formObject[field]) {
      errors.push(`${field} is required`);
    }
  }
  
  if (formObject.dateOfVisit) {
    const visitDate = new Date(formObject.dateOfVisit);
    if (isNaN(visitDate.getTime()) || visitDate > new Date()) {
      errors.push('Invalid visit date');
    }
  }
  
  if (formObject.VISIT_DESCRIPTION &&
      formObject.VISIT_DESCRIPTION.length > 1000) {
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
    
    // 2) Validate basic form data
    const errors = validateFormData(formObject);
    if (errors.length > 0) {
      return `Validation errors: ${errors.join(', ')}`;
    }

    // 3) (Optional) If user selected "Add New Customer" logic
    //    This is from your existing code if you want it:
    if (formObject.Customer === "**Add New Customer**" && formObject.newCustomer) {
      addNewCustomerWithCity(formObject.Region, formObject.newCustomerCity, formObject.newCustomer, currentUser, token);
      formObject.Customer = formObject.newCustomer;
    }

    // 4) Validate each contact person
    if (formObject.contactPersons) {
      const contactErrors = formObject.contactPersons.flatMap(validateContactPerson);
      if (contactErrors.length > 0) {
        return `Contact person validation errors: ${contactErrors.join(', ')}`;
      }
    }

    // 5) Format the contact persons into a single string if we still want to store in main sheet
    const formattedContacts = formatContactPersons(formObject.contactPersons);

    // 6) Next visit date
    const nextVisit = formObject.nextVisit ? new Date(formObject.nextVisit).toISOString() : '';

    // 7) Generate a unique ID for the main record and for referencing
    const uniqueId = generateUniqueId();

    // 8) Create the main record in "MAIN!A2:U2"
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
      formObject.VISIT_DESCRIPTION,
      formObject.IMAGE || "",
      getPakistanDateTime().toLocaleString(),
      formattedContacts,  // store the contact summary
      nextVisit
    ]];
    createRecord(values);

    // 9) Create separate contact records in CONTACTS sheet
    createContactsRecords(
      uniqueId,
      formObject.Customer,
      formObject.contactPersons,
      currentUser
    );

    // 10) Create workload record (if competitor info is provided)
    createMultipleWorkloadRecords(uniqueId, formObject.Customer,formObject.workloads);
    AuditLogger.log('Form Submission', currentUser, { id: uniqueId, customer: formObject.Customer });
    return "Data successfully submitted by Sales Person: " + currentUser;
  } catch (error) {
    ErrorLogger.log('processForm', error, { formObject, currentUser });
    AuditLogger.log('Form Submission Failure', currentUser, { error: error.message });
    return "Error: " + error.message;
  }
}

function createRecord(values) {
  try {
    let valueRange = Sheets.newValueRange();
    valueRange.values = values;
    Sheets.Spreadsheets.Values.append(
      valueRange, 
      Config.SPREADSHEET.ID, 
      Config.SPREADSHEET.RANGES.DATA, 
      { valueInputOption: "RAW" }
    );
  } catch (error) {
    ErrorLogger.log('createRecord', error, { values });
    throw new Error('Failed to save record. Please try again or contact support.');
  }
}

function generateUniqueId() {
  return Utilities.getUuid();
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
    
    // Find first empty row in customer range
    let customerRow = 2;
    while(sheet.getRange(customerRow, 1).getValue()) customerRow++;

    // Add data to first empty row
    sheet.getRange(customerRow, 1, 1, 3).setValues([[region, customerName, city]]);

    AuditLogger.log('ADD_CUSTOMER', username, { 
      region, city, customerName,
      customerRow
    });

    return "New customer added successfully!";
  } catch (error) {
    ErrorLogger.log('addNewCustomerWithCity', error, { region, city, customerName, username });
    throw new Error('Failed to add new customer: ' + error.message);
  }
}
/**************************
 * EXPORT FUNCTIONS
 **************************/
function getSalesPersonData(salesPersonName) {
  try {
    const data = readRecord(Config.SPREADSHEET.RANGES.DATA);
    return data.filter(row => row[7] === salesPersonName);
  } catch (error) {
    ErrorLogger.log('getSalesPersonData', error, { salesPersonName });
    return [];
  }
}

function exportSalesPersonDataAsCSV(salesPersonName, username, token) {
  try {
    // Validate session
    if (!sessionManager.validateSession(username, token)) {
      throw new Error('Invalid session');
    }
    
    // Rate limiting
    const currentTime = Math.floor(Date.now() / 1000);
    if (!userExportCounts[username]) {
      userExportCounts[username] = { count: 0, windowStart: currentTime };
    }
    
    const userExport = userExportCounts[username];
    
    if ((currentTime - userExport.windowStart) > Config.RATE_LIMITING.EXPORT_WINDOW_SECONDS) {
      // Reset
      userExport.count = 0;
      userExport.windowStart = currentTime;
    }
    
    if (userExport.count >= Config.RATE_LIMITING.EXPORT_LIMIT) {
      throw new Error('Export limit reached. Please try again later.');
    }
    
    userExport.count += 1;
    userExportCounts[username] = userExport;
    
    const spData = getSalesPersonData(salesPersonName);
    if (spData.length === 0) {
      throw new Error('No data available for export.');
    }

    // Build CSV content
    const headers = [
      "ID", "Date of Visit", "Visit Type", "Region", "City", 
      "Customer", "Department", "Sales Person", "Product", 
      "Analyzer Model", "Flags", "Visit Description", "Image URL", 
      "Submission Time", "Contact Persons", "Next Visit Date"
    ];
    let csvContent = headers.join(",") + "\n";
    
    spData.forEach(function(row) {
      const formattedRow = row.map(item => {
        const escapedItem = (`${item}`).replace(/"/g, '""');
        return /[",]/.test(escapedItem) ? `"${escapedItem}"` : escapedItem;
      }).join(",");
      csvContent += formattedRow + "\n";
    });

    // Convert CSV to Blob
    const blob = Utilities.newBlob(csvContent, 'text/csv', `${salesPersonName}_Data.csv`);
    const file = DriveApp.createFile(blob);
    const downloadUrl = file.getDownloadUrl();
    
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // Schedule deletion after 1 hour
    ScriptApp.newTrigger('deleteTempExportFile')
             .timeBased()
             .after(60 * 60 * 1000)
             .create();

    AuditLogger.log('Data Export', username, { salesPersonName, fileName: `${salesPersonName}_Data.csv`, downloadUrl });
    return downloadUrl;
  } catch (error) {
    ErrorLogger.log('exportSalesPersonDataAsCSV', error, { salesPersonName, username });
    return "";
  }
}

/**
 * Deletes the temporary export file after a certain period.
 */
function deleteTempExportFile() {
  try {
    const files = DriveApp.getFilesByName(/_Data\.csv$/);
    while (files.hasNext()) {
      const file = files.next();
      file.setTrashed(true);
    }
  } catch (error) {
    ErrorLogger.log('deleteTempExportFile', error);
  }
}

/**************************
 * AUDIT LOGGING
 **************************/
class AuditLogger {
  static async log(action, username, details = {}) {
       
    try {
      const auditLog = [
        getPakistanDateTime().toISOString(),
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
 * INPUT SANITIZATION
 **************************/
function sanitizeInput(input) {
  if (typeof input !== 'string') return input;
  return input
    .trim()
    .replace(/[<>]/g, '') // Remove < and >
    .replace(/[&]/g, 'and') // Replace &
    .replace(/['"]/g, '') // Remove quotes
    .substring(0, 1000); // Limit length
}
/**************************
 * helper function pak time
 **************************/

function getPakistanDateTime() {
  // Create date in UTC
  const date = new Date();
  
  // Convert to Pakistan time (UTC+5)
  const pkTime = new Date(date.getTime() + (5 * 60 * 60 * 1000));
  
  return pkTime;
}

/**************************
 * TEST FUNCTIONS
 **************************/
function testErrorLogging() {
  try {
    throw new Error("This is a test error for Error Logging.");
  } catch (e) {
    ErrorLogger.log('testErrorLogging', e, { testKey: 'testValue' });
  }
}

function testCreateRecord() {
  const testValues = [[
    'test-id',
    '2025-01-01',
    'NEW SALES CALL',
    'Region1',
    'City1',
    'Customer1',
    'Department1',
    'Fasi ullah',
    'Product1',
    'Analyzer1',
    'Flag1',
    'Description of visit',
    'https://example.com/image.png',
    new Date().toLocaleString(),
    'A/John Doe/Manager/1234567890',
    '2025-01-15'
  ]];
  createRecord(testValues);
}
