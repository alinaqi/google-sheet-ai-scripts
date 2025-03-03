/**
 * Contact Enrichment and Outreach System for Google Sheets
 * 
 * This script:
 * 1. Reads contact data from a Google Sheet (Name and Email in separate columns)
 * 2. Uses Perplexity API to find LinkedIn information about each contact
 * 3. Finds potential connections at the same organization
 * 4. Uses Anthropic API to generate personalized outreach suggestions for zenloop
 * 5. Writes all this information back to the Google Sheet
 * 6. Logs all operations and errors to a dedicated log sheet
 */

// Your API keys - replace with actual keys
const PERPLEXITY_API_KEY = "your_perplexity_key";
const ANTHROPIC_API_KEY = "your_anthropic_key";

// Constants
const LOG_SHEET_NAME = "Process Logs";
const MAX_RETRIES = 3;
const RETRY_DELAY_MS = 2000;

/**
 * Creates a menu item in Google Sheets UI
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Contact Enrichment')
    .addItem('Enrich Selected Contacts', 'enrichSelectedContacts')
    .addSeparator()
    .addItem('View Process Logs', 'viewLogs')
    .addItem('Clear Process Logs', 'clearLogs')
    .addToUi();
}

/**
 * Sets up the log sheet
 */
function setupLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  
  // Create log sheet if it doesn't exist
  if (!logSheet) {
    logSheet = ss.insertSheet(LOG_SHEET_NAME);
    const headers = ["Timestamp", "Level", "Message"];
    logSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    logSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    logSheet.setFrozenRows(1);
    
    // Auto-resize columns
    for (let i = 1; i <= headers.length; i++) {
      logSheet.autoResizeColumn(i);
    }
    
    // Add color formatting for log levels
    const levelRange = logSheet.getRange("B:B");
    const rule1 = SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains("ERROR")
      .setBackground("#F4CCCC")  // Light red
      .setRanges([levelRange])
      .build();
    const rule2 = SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains("WARNING")
      .setBackground("#FCE5CD")  // Light orange
      .setRanges([levelRange])
      .build();
    const rule3 = SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains("INFO")
      .setBackground("#D9EAD3")  // Light green
      .setRanges([levelRange])
      .build();
    const rules = logSheet.getConditionalFormatRules();
    rules.push(rule1, rule2, rule3);
    logSheet.setConditionalFormatRules(rules);
  }
  
  return logSheet;
}

/**
 * Logs a message to the log sheet
 */
function logMessage(level, message) {
  // Also log to Apps Script logger
  Logger.log(`[${level}] ${message}`);
  
  try {
    const logSheet = setupLogSheet();
    const timestamp = new Date().toISOString();
    logSheet.appendRow([timestamp, level, message]);
  } catch (error) {
    // If logging fails, just use the Apps Script logger as fallback
    Logger.log(`Error writing to log sheet: ${error.toString()}`);
  }
}

/**
 * Views the process logs
 */
function viewLogs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  
  if (!logSheet) {
    SpreadsheetApp.getUi().alert('No log sheet found. Please run a process first.');
    return;
  }
  
  // Activate the log sheet
  logSheet.activate();
}

/**
 * Clears the process logs
 */
function clearLogs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  
  if (!logSheet) {
    SpreadsheetApp.getUi().alert('No log sheet found.');
    return;
  }
  
  // Clear log data but keep headers
  const lastRow = logSheet.getLastRow();
  if (lastRow > 1) {
    logSheet.deleteRows(2, lastRow - 1);
  }
  
  SpreadsheetApp.getUi().alert('Logs have been cleared.');
}

/**
 * Main function to enrich selected contacts
 */
function enrichSelectedContacts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Get all data to find column indices
  const allData = sheet.getDataRange().getValues();
  if (allData.length <= 1) {
    logMessage("ERROR", "Sheet is empty or only has headers");
    SpreadsheetApp.getUi().alert('Sheet is empty or only has headers. Please add contact data.');
    return;
  }
  
  // Find column indices from headers
  const headers = allData[0];
  const nameColIndex = headers.indexOf("Name");
  // Note: We're not using the "Email" column for email address extraction,
  // since the email is contained within the Name cell
  const loginColIndex = headers.indexOf("No of logins");
  const linkedinColIndex = headers.indexOf("LinkedIn Information");
  const connectionsColIndex = headers.indexOf("Potential Connections") !== -1 ? 
                              headers.indexOf("Potential Connections") : 
                              headers.indexOf("Potential Conn");
  const outreachColIndex = headers.indexOf("Outreach Suggestions") !== -1 ? 
                          headers.indexOf("Outreach Suggestions") : 
                          headers.indexOf("Outreach Sugg");
  const statusColIndex = headers.indexOf("Process Status");
  
  // Validate column indices
  if (nameColIndex === -1) {
    logMessage("ERROR", `Required 'Name' column not found.`);
    SpreadsheetApp.getUi().alert('Required "Name" column not found.');
    return;
  }
  
  // Log column indices for debugging
  logMessage("INFO", `Column indices - Name: ${nameColIndex}, LinkedIn: ${linkedinColIndex}, ` +
             `Connections: ${connectionsColIndex}, Outreach: ${outreachColIndex}, Status: ${statusColIndex}`);
  
  // Get the selected rows
  const selection = sheet.getActiveRange();
  const startRow = selection.getRow();
  const numRows = selection.getNumRows();
  
  // Validate selection
  if (startRow < 2) { // Assuming first row is headers
    logMessage("WARNING", "Selection includes header row. Will skip header row in processing.");
  }
  
  logMessage("INFO", `Selected rows ${startRow} to ${startRow + numRows - 1} for processing`);
  
  // Process each selected row
  let successCount = 0;
  let errorCount = 0;
  
  for (let i = 0; i < numRows; i++) {
    const currentRow = startRow + i;
    
    // Skip header row
    if (currentRow === 1) {
      logMessage("INFO", "Skipping header row");
      continue;
    }
    
    // Get name and email from the row
    const rowData = sheet.getRange(currentRow, 1, 1, headers.length).getValues()[0];
    
    // The "Name" column contains both name and email on separate lines
    const nameCell = String(rowData[nameColIndex] || "");
    logMessage("INFO", `Name cell content from row ${currentRow}: "${nameCell}"`);
    
    // Try to extract name and email from the name cell
    const nameLines = nameCell.split(/\r?\n/).map(line => line.trim()).filter(line => line);
    
    if (nameLines.length < 2) {
      const message = `Skipping row ${currentRow}: Name cell doesn't contain both name and email (separated by line break)`;
      logMessage("WARNING", message);
      if (statusColIndex !== -1) {
        sheet.getRange(currentRow, statusColIndex + 1).setValue("SKIPPED: Format error");
        sheet.getRange(currentRow, statusColIndex + 1).setBackground("#D9D9D9"); // Light gray
      }
      continue;
    }
    
    const name = nameLines[0];
    const email = nameLines[1];
    
    // Skip if name or email is empty
    if (!name || !email) {
      const message = `Skipping row ${currentRow}: Missing name or email`;
      logMessage("WARNING", message);
      if (statusColIndex !== -1) {
        sheet.getRange(currentRow, statusColIndex + 1).setValue("SKIPPED: Missing data");
        sheet.getRange(currentRow, statusColIndex + 1).setBackground("#D9D9D9"); // Light gray
      }
      continue;
    }
    
    // Log the contact we're processing
    logMessage("INFO", `Processing contact at row ${currentRow}: ${name} (${email})`);
    
    // Extract domain from email to identify company
    try {
      const domain = email.split('@')[1];
      if (!domain) {
        throw new Error("Invalid email format");
      }
      
      const company = domain.split('.')[0];
      
      // Update status to show processing
      if (statusColIndex !== -1) {
        sheet.getRange(currentRow, statusColIndex + 1).setValue("PROCESSING");
        sheet.getRange(currentRow, statusColIndex + 1).setBackground("#FCE5CD"); // Light orange
      }
      
      // Process this contact
      try {
        // Get LinkedIn information
        logMessage("INFO", `Fetching LinkedIn info for ${name} at ${company}`);
        const linkedinInfo = getLinkedInInfo(name, company);
        if (linkedinColIndex !== -1) {
          sheet.getRange(currentRow, linkedinColIndex + 1).setValue(linkedinInfo);
        }
        
        // Get potential connections at the same company
        logMessage("INFO", `Finding potential connections at ${company}`);
        const potentialConnections = getPotentialConnections(company, name);
        if (connectionsColIndex !== -1) {
          sheet.getRange(currentRow, connectionsColIndex + 1).setValue(potentialConnections);
        }
        
        // Generate outreach suggestions
        logMessage("INFO", `Generating outreach suggestions for ${name}`);
        const outreachSuggestions = generateOutreachSuggestions(name, linkedinInfo, potentialConnections, company);
        if (outreachColIndex !== -1) {
          sheet.getRange(currentRow, outreachColIndex + 1).setValue(outreachSuggestions);
        }
        
        // Update status to show completion
        if (statusColIndex !== -1) {
          sheet.getRange(currentRow, statusColIndex + 1).setValue("COMPLETED");
          sheet.getRange(currentRow, statusColIndex + 1).setBackground("#D9EAD3"); // Light green
        }
        
        successCount++;
        
        // Add a small delay to avoid rate limits
        Utilities.sleep(RETRY_DELAY_MS);
      } catch (error) {
        const errorMessage = `Error processing ${name}: ${error.toString()}`;
        logMessage("ERROR", errorMessage);
        
        if (linkedinColIndex !== -1) {
          sheet.getRange(currentRow, linkedinColIndex + 1).setValue("Error: " + error.toString());
        }
        
        if (statusColIndex !== -1) {
          sheet.getRange(currentRow, statusColIndex + 1).setValue("ERROR: Processing failed");
          sheet.getRange(currentRow, statusColIndex + 1).setBackground("#F4CCCC"); // Light red
        }
        
        errorCount++;
      }
    } catch (error) {
      const errorMessage = `Error with email format for ${name}: ${error.toString()}`;
      logMessage("ERROR", errorMessage);
      
      if (statusColIndex !== -1) {
        sheet.getRange(currentRow, statusColIndex + 1).setValue("ERROR: Invalid email format");
        sheet.getRange(currentRow, statusColIndex + 1).setBackground("#F4CCCC"); // Light red
      }
      
      errorCount++;
    }
  }
  
  const completionMessage = `Contact enrichment completed! Successful: ${successCount}, Errors: ${errorCount}`;
  logMessage("INFO", completionMessage);
  SpreadsheetApp.getUi().alert(completionMessage);
}

/**
 * Query Perplexity API to get LinkedIn information about a person
 */
function getLinkedInInfo(name, company) {
  const query = `Find detailed professional information about ${name} who works at ${company}. Include their current role, years at the company, previous experience, education, skills, and any notable projects or achievements. Focus on information that would be available on LinkedIn or similar professional profiles.`;
  
  return callAPIWithRetries("Perplexity", () => queryPerplexityAPI(query));
}

/**
 * Query Perplexity API to find potential connections at the same company
 */
function getPotentialConnections(company, excludeName) {
  const query = `Find 3-5 key decision-makers or team leaders at ${company} who might be connected to customer experience, digital transformation, or feedback management (excluding ${excludeName}). For each person, provide their name, role, and brief background that makes them relevant for a CX AI platform like zenloop.`;
  
  return callAPIWithRetries("Perplexity", () => queryPerplexityAPI(query));
}

/**
 * Generate personalized outreach suggestions using Anthropic API
 */
function generateOutreachSuggestions(name, linkedinInfo, potentialConnections, company) {
  const prompt = `
You are a B2B sales expert specialized in AI and customer experience platforms. 
I need to reach out to ${name} at ${company}. 

Here's what I know about them:
${linkedinInfo}

Other potential contacts at the company include:
${potentialConnections}

Based on this information, provide concise, personalized outreach suggestions for introducing zenloop, an AI-based customer experience platform. Include:

1. A compelling subject line for an email
2. A brief introduction that shows I've done my homework
3. A value proposition specifically tailored to their role and company
4. A clear, low-pressure call-to-action

Keep the suggestions action-oriented and focused on how zenloop can solve specific problems they might be facing with customer feedback.
`;

  return callAPIWithRetries("Anthropic", () => queryAnthropicAPI(prompt));
}

/**
 * Call API with retry mechanism
 */
function callAPIWithRetries(apiName, apiCallFunction) {
  let retries = 0;
  
  while (retries < MAX_RETRIES) {
    try {
      const result = apiCallFunction();
      return result;
    } catch (error) {
      retries++;
      logMessage("WARNING", `${apiName} API call failed (attempt ${retries}/${MAX_RETRIES}): ${error.toString()}`);
      
      // If we've reached max retries, throw the error
      if (retries >= MAX_RETRIES) {
        throw new Error(`${apiName} API call failed after ${MAX_RETRIES} attempts: ${error.toString()}`);
      }
      
      // Exponential backoff
      const backoffDelay = RETRY_DELAY_MS * Math.pow(2, retries - 1);
      logMessage("INFO", `Retrying in ${backoffDelay/1000} seconds...`);
      Utilities.sleep(backoffDelay);
    }
  }
}

/**
 * Call the Perplexity API
 */
function queryPerplexityAPI(query) {
  const url = "https://api.perplexity.ai/chat/completions";
  
  // Format the payload to match the OpenAI-compatible format
  const payload = {
    "model": "sonar-pro",
    "messages": [
      {
        "role": "system",
        "content": "You are a helpful AI assistant that provides accurate, detailed information. Focus on professional details and factual information."
      },
      {
        "role": "user",
        "content": query
      }
    ],
    "max_tokens": 2000
  };
  
  const options = {
    "method": "post",
    "headers": {
      "Authorization": `Bearer ${PERPLEXITY_API_KEY}`,
      "Content-Type": "application/json"
    },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };
  
  logMessage("INFO", `Calling Perplexity API with query length: ${query.length} characters`);
  const startTime = new Date().getTime();
  
  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  const endTime = new Date().getTime();
  
  logMessage("INFO", `Perplexity API responded in ${(endTime - startTime)/1000} seconds with code ${responseCode}`);
  
  if (responseCode !== 200) {
    const errorText = response.getContentText();
    logMessage("ERROR", `Perplexity API error: ${responseCode}, ${errorText}`);
    throw new Error(`Perplexity API error: ${responseCode}, ${errorText.substring(0, 200)}...`);
  }
  
  try {
    const responseJson = JSON.parse(response.getContentText());
    
    // Extract just the message content
    if (responseJson.choices && responseJson.choices.length > 0) {
      const content = responseJson.choices[0].message.content;
      return content;
    } else {
      throw new Error("No valid response content from Perplexity API");
    }
  } catch (error) {
    logMessage("ERROR", `Error parsing Perplexity API response: ${error.toString()}`);
    throw new Error(`Error parsing Perplexity API response: ${error.toString()}`);
  }
}

/**
 * Call the Anthropic Claude API
 */
function queryAnthropicAPI(prompt) {
  const url = "https://api.anthropic.com/v1/messages";
  
  const payload = {
    "model": "claude-3-7-sonnet-20250219",
    "max_tokens": 800,
    "messages": [
      {
        "role": "user",
        "content": prompt
      }
    ]
  };
  
  const options = {
    "method": "post",
    "headers": {
      "x-api-key": ANTHROPIC_API_KEY,
      "anthropic-version": "2023-06-01",
      "Content-Type": "application/json"
    },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };
  
  logMessage("INFO", `Calling Anthropic API with prompt length: ${prompt.length} characters`);
  const startTime = new Date().getTime();
  
  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  const endTime = new Date().getTime();
  
  logMessage("INFO", `Anthropic API responded in ${(endTime - startTime)/1000} seconds with code ${responseCode}`);
  
  if (responseCode !== 200) {
    const errorText = response.getContentText();
    logMessage("ERROR", `Anthropic API error: ${responseCode}, ${errorText}`);
    throw new Error(`Anthropic API error: ${responseCode}, ${errorText.substring(0, 200)}...`);
  }
  
  try {
    const responseJson = JSON.parse(response.getContentText());
    
    // Extract just the message content
    if (responseJson.content && responseJson.content.length > 0) {
      const text = responseJson.content[0].text;
      return text;
    } else {
      throw new Error("No valid response content from Anthropic API");
    }
  } catch (error) {
    logMessage("ERROR", `Error parsing Anthropic API response: ${error.toString()}`);
    throw new Error(`Error parsing Anthropic API response: ${error.toString()}`);
  }
}
