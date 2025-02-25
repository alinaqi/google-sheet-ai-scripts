// Configuration
const PERPLEXITY_API_KEY = 'your_key';
const PERPLEXITY_API_URL = 'your_key';
const ANTHROPIC_API_KEY = 'your_key';
const ANTHROPIC_API_URL = 'https://api.anthropic.com/v1/messages';

// Add OpenAI configuration
const OPENAI_API_KEY = 'your_key';
const OPENAI_API_URL = 'https://api.openai.com/v1/chat/completions';

// Column indices (1-based)
const COLUMNS = {
  COMPANY_NAME: 1,
  WEBSITE: 2,
  BUSINESS_OVERVIEW: 3,
  TARGET_AUDIENCE: 4,
  PRODUCTS: 5,
  PRICING: 6
};

// Create a log sheet if it doesn't exist
function ensureLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName('ProcessLog');
  if (!logSheet) {
    logSheet = ss.insertSheet('ProcessLog');
    logSheet.appendRow(['Timestamp', 'Process', 'Status', 'Details']);
    logSheet.setFrozenRows(1);
  }
  return logSheet;
}

// Log function
function logToSheet(process, status, details) {
  const logSheet = ensureLogSheet();
  const timestamp = new Date().toLocaleString();
  logSheet.appendRow([timestamp, process, status, details]);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Company Analysis')
    .addItem('Populate Company Information', 'populateCompanyInfo')
    .addItem('Analyze Collaboration Opportunities', 'analyzeCollaborations')
    .addItem('Analyze Collaboration Probability', 'analyzeCollaborationProbability')
    .addSeparator()
    .addItem('Reset Analysis Progress', 'resetAnalysisProgress')
    .addItem('⚠️ Stop Running Scripts', 'stopScripts')
    .addSeparator()
    .addItem('Clear Process Log', 'clearLog')
    .addToUi();
}

function stopScripts() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    '⚠️ Stop Scripts',
    'To stop running scripts:\n\n' +
    '1. Click the "Stop" button (■) in the toolbar\n' +
    '2. Or press Ctrl + Alt + Shift + K (Windows)\n' +
    '3. Or press Cmd + Option + Shift + K (Mac)\n\n' +
    'After stopping, please wait a few seconds before running new scripts.',
    ui.ButtonSet.OK
  );
}

function clearLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('ProcessLog');
  if (logSheet) {
    logSheet.clear();
    logSheet.appendRow(['Timestamp', 'Process', 'Status', 'Details']);
  }
}

function showProgress(message) {
  SpreadsheetApp.getActive().toast(message, '🔄 Progress Update');
}

function populateCompanyInfo() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('List of companies');
  const lastRow = sheet.getLastRow();
  
  // Get all existing data at once for better performance
  const allData = sheet.getRange(2, 1, lastRow - 1, COLUMNS.PRICING).getValues();
  
  logToSheet('Company Info Population', 'Started', `Checking ${lastRow - 1} companies`);
  showProgress(`Starting to check company information for ${lastRow - 1} companies...`);
  
  let successCount = 0;
  let errorCount = 0;
  let skippedCount = 0;
  let retryCount = 0;
  const maxRetries = 2;  // Maximum number of retries per company
  
  // Start from row 2 to skip header
  for (let row = 2; row <= lastRow; row++) {
    const rowIndex = row - 2; // Index in allData array
    const companyName = allData[rowIndex][COLUMNS.COMPANY_NAME - 1];
    const companyWebsite = allData[rowIndex][COLUMNS.WEBSITE - 1];
    
    if (!companyName) {
      skippedCount++;
      continue;
    }
    
    // Check if company already has information
    const hasBusinessOverview = !isEmpty(allData[rowIndex][COLUMNS.BUSINESS_OVERVIEW - 1]);
    const hasTargetAudience = !isEmpty(allData[rowIndex][COLUMNS.TARGET_AUDIENCE - 1]);
    const hasProducts = !isEmpty(allData[rowIndex][COLUMNS.PRODUCTS - 1]);
    const hasPricing = !isEmpty(allData[rowIndex][COLUMNS.PRICING - 1]);
    
    if (hasBusinessOverview && hasTargetAudience && hasProducts && hasPricing) {
      skippedCount++;
      showProgress(`Skipping ${companyName} (already has information) (${row-1}/${lastRow-1}, Skipped: ${skippedCount})`);
      logToSheet('Company Processing', 'Skipped', `${companyName} already has complete information`);
      continue;
    }
    
    showProgress(`Processing ${companyName} (${row-1}/${lastRow-1}, Skipped: ${skippedCount})`);
    logToSheet('Company Processing', 'In Progress', `Starting ${companyName}`);
    
    let attempts = 0;
    let success = false;
    
    while (attempts < maxRetries && !success) {
      try {
        attempts++;
        if (attempts > 1) {
          logToSheet('Retry', 'Info', `Attempt ${attempts} for ${companyName}`);
          Utilities.sleep(2000);  // Wait longer between retries
        }
        
        const companyInfo = fetchCompanyInfo(companyName, companyWebsite);
        
        // Validate the response
        if (!companyInfo || typeof companyInfo !== 'object') {
          throw new Error('Invalid response format');
        }
        
        updateCompanyRow(sheet, row, companyInfo);
        successCount++;
        success = true;
        logToSheet('Company Processing', 'Success', `Completed ${companyName}`);
        
      } catch (error) {
        logToSheet('Company Processing', 'Error', `Attempt ${attempts} failed for ${companyName}: ${error.message}`);
        
        if (attempts === maxRetries) {
          errorCount++;
          // Mark the row with error
          sheet.getRange(row, COLUMNS.BUSINESS_OVERVIEW)
               .setValue(`Error processing: ${error.message}. Please try again.`)
               .setBackground('#ffcdd2');  // Light red background
        } else {
          retryCount++;
        }
      }
    }
    
    // Add delay to avoid rate limits
    Utilities.sleep(1000);
  }
  
  const finalMessage = `Completed! Success: ${successCount}, Skipped: ${skippedCount}, Errors: ${errorCount}, Retries: ${retryCount}`;
  showProgress(finalMessage);
  logToSheet('Company Info Population', 'Completed', finalMessage);
}

function fetchCompanyInfo(companyName, companyWebsite) {
  const messages = [
    {
      "role": "system",
      "content": "You are a TOP MARKET RESEARCHER. Return ONLY a JSON object with the following structure, no other text: { \"businessOverview\": \"...\", \"targetAudience\": \"...\", \"products\": \"...\", \"pricing\": \"...\" }"
    },
    {
      "role": "user",
      "content": `Research and provide information about ${companyName} (website: ${companyWebsite}) in the specified JSON format. Include ONLY the JSON object, no other text.`
    }
  ];

  const options = {
    'method': 'post',
    'headers': {
      'Authorization': `Bearer ${PERPLEXITY_API_KEY}`,
      'Content-Type': 'application/json'
    },
    'payload': JSON.stringify({
      'model': 'sonar-pro',
      'messages': messages
    }),
    'muteHttpExceptions': true
  };

  try {
    const response = UrlFetchApp.fetch(PERPLEXITY_API_URL, options);
    const jsonResponse = JSON.parse(response.getContentText());
    
    // Log the raw response for debugging
    console.log(`Raw response for ${companyName}:`, JSON.stringify(jsonResponse));
    logToSheet('API Response', 'Debug', `Response for ${companyName}: ${JSON.stringify(jsonResponse.choices[0].message.content)}`);

    try {
      // Try to parse the content as JSON
      const companyInfo = JSON.parse(jsonResponse.choices[0].message.content);
      
      // Validate the required fields exist
      const requiredFields = ['businessOverview', 'targetAudience', 'products', 'pricing'];
      const missingFields = requiredFields.filter(field => !companyInfo[field]);
      
      if (missingFields.length > 0) {
        throw new Error(`Missing required fields: ${missingFields.join(', ')}`);
      }
      
      return companyInfo;
    } catch (parseError) {
      // If JSON parsing fails, try to extract JSON from the text
      const content = jsonResponse.choices[0].message.content;
      const jsonMatch = content.match(/\{[\s\S]*\}/);
      
      if (jsonMatch) {
        const extractedJson = JSON.parse(jsonMatch[0]);
        return extractedJson;
      }
      
      logToSheet('Parsing Error', 'Error', `Failed to parse JSON for ${companyName}: ${parseError.message}`);
      throw new Error(`Failed to parse response: ${parseError.message}`);
    }
  } catch (error) {
    logToSheet('API Error', 'Error', `API call failed for ${companyName}: ${error.message}`);
    throw new Error(`API call failed: ${error.message}`);
  }
}

function updateCompanyRow(sheet, row, companyInfo) {
  sheet.getRange(row, COLUMNS.BUSINESS_OVERVIEW).setValue(companyInfo.businessOverview);
  sheet.getRange(row, COLUMNS.TARGET_AUDIENCE).setValue(companyInfo.targetAudience);
  sheet.getRange(row, COLUMNS.PRODUCTS).setValue(companyInfo.products);
  sheet.getRange(row, COLUMNS.PRICING).setValue(companyInfo.pricing);
}

// Helper function to clean up text that might interfere with JSON parsing
function cleanJsonString(str) {
  // Remove any potential markdown code block syntax
  str = str.replace(/```json/g, '').replace(/```/g, '');
  // Remove any leading/trailing whitespace
  str = str.trim();
  return str;
}

function analyzeCollaborations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const matrixSheet = ss.getSheetByName('CollaborationMatrix');
  const companyListSheet = ss.getSheetByName('List of companies');
  
  logToSheet('Collaboration Analysis', 'Started', 'Beginning collaboration analysis');
  showProgress('Starting collaboration analysis...');
  
  // Get company information map
  showProgress('Loading company information...');
  const companyInfo = getCompanyInformation(companyListSheet);
  
  // Debug log to show loaded companies
  const loadedCompanies = Object.keys(companyInfo);
  logToSheet('Data Loading', 'Debug', `Loaded ${loadedCompanies.length} companies: ${loadedCompanies.join(', ')}`);
  
  // Get the range of companies from matrix
  const matrixCompanies = matrixSheet.getRange(1, 2, 1, matrixSheet.getLastColumn() - 1).getValues()[0];
  logToSheet('Data Loading', 'Debug', `Matrix contains companies: ${matrixCompanies.join(', ')}`);
  
  // Verify all matrix companies exist in loaded data
  const missingCompanies = matrixCompanies.filter(company => company && !companyInfo[company]);
  if (missingCompanies.length > 0) {
    logToSheet('Data Loading', 'Warning', `Companies in matrix but missing from company list: ${missingCompanies.join(', ')}`);
  }

  // Get the range of companies
  const numCompanies = matrixSheet.getLastRow() - 1; // Subtract 1 for header
  const totalPairs = (numCompanies * (numCompanies - 1)) / 2;
  let processedPairs = 0;
  let skippedPairs = 0;
  let successCount = 0;
  let errorCount = 0;
  
  // Get all existing values at once for better performance
  const existingValues = matrixSheet.getRange(2, 2, numCompanies, numCompanies).getValues();
  
  logToSheet('Collaboration Analysis', 'In Progress', `Total pairs to check: ${totalPairs}`);
  
  // Process each cell in the matrix
  for (let row = 2; row <= numCompanies + 1; row++) {
    for (let col = 2; col <= numCompanies + 1; col++) {
      // Skip if companies are the same
      if (row === col) {
        continue;
      }
      
      const company1 = matrixSheet.getRange(1, col).getValue();
      const company2 = matrixSheet.getRange(row, 1).getValue();
      
      // Check if cell already has content (using the pre-fetched values)
      const existingValue = existingValues[row - 2][col - 2];
      if (!isEmpty(existingValue)) {
        skippedPairs++;
        processedPairs++;
        showProgress(`Skipping ${company1} & ${company2} (already analyzed) (${processedPairs}/${totalPairs})`);
        continue;
      }
      
      processedPairs++;
      showProgress(`Analyzing ${company1} & ${company2} (${processedPairs}/${totalPairs}, Skipped: ${skippedPairs})`);
      
      try {
        const collaboration = analyzeCompanyPair(
          company1, 
          company2, 
          companyInfo[company1], 
          companyInfo[company2]
        );
        
        // Update both cells (matrix is symmetric)
        matrixSheet.getRange(row, col).setValue(collaboration);
        matrixSheet.getRange(col, row).setValue(collaboration);
        
        successCount++;
        logToSheet('Pair Analysis', 'Success', `Completed ${company1} & ${company2}`);
        
        // Add delay to avoid rate limits
        Utilities.sleep(1000);
      } catch (error) {
        errorCount++;
        logToSheet('Pair Analysis', 'Error', `Failed ${company1} & ${company2}: ${error.message}`);
        continue;
      }
    }
  }
  
  const finalMessage = `Analysis Complete! Successful pairs: ${successCount}, Skipped pairs: ${skippedPairs}, Errors: ${errorCount}`;
  showProgress(finalMessage);
  logToSheet('Collaboration Analysis', 'Completed', finalMessage);
}

function analyzeCompanyPair(company1, company2, info1, info2) {
  // Validate company information
  if (!info1 || !info2) {
    logToSheet('Collaboration Analysis', 'Error', `Missing company information for ${company1} or ${company2}`);
    return `Unable to analyze collaboration: Missing company information for ${!info1 ? company1 : company2}`;
  }

  // Ensure all required fields exist with default values if missing
  info1 = {
    domain: info1.domain || 'Not provided',
    businessOverview: info1.businessOverview || 'Not provided',
    targetAudience: info1.targetAudience || 'Not provided',
    products: info1.products || 'Not provided'
  };

  info2 = {
    domain: info2.domain || 'Not provided',
    businessOverview: info2.businessOverview || 'Not provided',
    targetAudience: info2.targetAudience || 'Not provided',
    products: info2.products || 'Not provided'
  };

  try {
    const options = {
      'method': 'post',
      'headers': {
        'x-api-key': ANTHROPIC_API_KEY,
        'content-type': 'application/json',
        'anthropic-version': '2023-06-01'
      },
      'payload': JSON.stringify({
        'model': 'claude-3-5-sonnet-20241022',
        'max_tokens': 8096,
        'temperature': 0.7,
        'system': 'You are a business strategy consultant specializing in identifying collaboration opportunities between companies. Analyze the provided company information and suggest specific, actionable collaboration opportunities. Keep the response under 100 words and focus on the most impactful opportunity.',
        'messages': [
          {
            'role': 'user',
            'content': [
              {
                'type': 'text',
                'text': `Analyze collaboration opportunities between these companies. Focus on their products and where there can be easy wins such as cross and upselling opportunities:
                 
                 Company 1: ${company1}
                 Domain: ${info1.domain}
                 Business Overview: ${info1.businessOverview}
                 Target Audience: ${info1.targetAudience}
                 Products: ${info1.products}
                 
                 Company 2: ${company2}
                 Domain: ${info2.domain}
                 Business Overview: ${info2.businessOverview}
                 Target Audience: ${info2.targetAudience}
                 Products: ${info2.products}
                 
                 What are the most promising collaboration opportunities between these companies? Define a roadmap and a potential pitch.`
              }
            ]
          }
        ]
      }),
      'muteHttpExceptions': true
    };

    const response = UrlFetchApp.fetch(ANTHROPIC_API_URL, options);
    const jsonResponse = JSON.parse(response.getContentText());
    
    // Log the raw response for debugging
    logToSheet('API Response', 'Debug', `Raw response for ${company1} & ${company2}: ${JSON.stringify(jsonResponse)}`);
    
    // Handle the response according to the new format
    if (!jsonResponse || !jsonResponse.content || !Array.isArray(jsonResponse.content)) {
      throw new Error(`Invalid API response format: ${JSON.stringify(jsonResponse)}`);
    }
    
    // Extract text from the response
    const textContent = jsonResponse.content.find(item => item.type === 'text');
    if (!textContent || !textContent.text) {
      throw new Error('No text content found in response');
    }
    
    return textContent.text;
  } catch (error) {
    const errorMessage = `API call failed for ${company1} & ${company2}: ${error.message}`;
    logToSheet('API Error', 'Error', errorMessage);
    return `Error analyzing collaboration: ${error.message}. Please check the ProcessLog sheet for details.`;
  }
}

// Helper function to get company information from the sheet
function getCompanyInformation(sheet) {
  if (!sheet) {
    logToSheet('Data Loading', 'Error', 'Sheet not found');
    throw new Error('Sheet not found');
  }

  const lastRow = sheet.getLastRow();
  const companyInfo = {};
  
  // Get all data at once to improve performance
  const data = sheet.getRange(2, 1, lastRow - 1, COLUMNS.PRICING).getValues();
  
  data.forEach((row, index) => {
    const companyName = row[COLUMNS.COMPANY_NAME - 1];
    if (!companyName) return;
    
    // Log missing data for debugging
    const missingFields = [];
    if (!row[COLUMNS.WEBSITE - 1]) missingFields.push('Website');
    if (!row[COLUMNS.BUSINESS_OVERVIEW - 1]) missingFields.push('Business Overview');
    if (!row[COLUMNS.TARGET_AUDIENCE - 1]) missingFields.push('Target Audience');
    if (!row[COLUMNS.PRODUCTS - 1]) missingFields.push('Products');
    if (!row[COLUMNS.PRICING - 1]) missingFields.push('Pricing');
    
    if (missingFields.length > 0) {
      logToSheet('Data Loading', 'Warning', `Company ${companyName} is missing: ${missingFields.join(', ')}`);
    }
    
    companyInfo[companyName] = {
      domain: row[COLUMNS.WEBSITE - 1] || '',
      businessOverview: row[COLUMNS.BUSINESS_OVERVIEW - 1] || '',
      targetAudience: row[COLUMNS.TARGET_AUDIENCE - 1] || '',
      products: row[COLUMNS.PRODUCTS - 1] || '',
      pricing: row[COLUMNS.PRICING - 1] || ''
    };
  });
  
  if (Object.keys(companyInfo).length === 0) {
    logToSheet('Data Loading', 'Warning', 'No company information found in sheet');
  }
  
  return companyInfo;
}

// Helper function to check if a cell is empty (removing duplicate)
function isEmpty(value) {
  return value === null || value === undefined || value === '';
}

function analyzeCollaborationProbability() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const matrixSheet = ss.getSheetByName('CollaborationMatrix');
  const companyListSheet = ss.getSheetByName('List of companies');
  
  // Get progress from Properties Service
  const scriptProperties = PropertiesService.getScriptProperties();
  let startRow = parseInt(scriptProperties.getProperty('lastProcessedRow')) || 17;  // Start from row 17 if no saved progress
  let startCol = parseInt(scriptProperties.getProperty('lastProcessedCol')) || 17;  // Start from column Q (17) if no saved progress
  
  logToSheet('Probability Analysis', 'Started', `Beginning collaboration probability analysis from cell: Column ${startCol}, Row ${startRow}`);
  showProgress(`Starting collaboration probability analysis from position: Column ${startCol}, Row ${startRow}...`);
  
  // Get company information map
  showProgress('Loading company information...');
  const companyInfo = getCompanyInformation(companyListSheet);
  
  // Get the range of companies
  const numCompanies = matrixSheet.getLastRow() - 1;
  const totalPairs = (numCompanies * (numCompanies - 1)) / 2;
  let processedPairs = 0;
  let skippedPairs = 0;
  let successCount = 0;
  let errorCount = 0;
  
  // Get all existing values at once for better performance
  const existingValues = matrixSheet.getRange(2, 2, numCompanies, numCompanies).getValues();
  
  logToSheet('Probability Analysis', 'In Progress', `Total pairs to check: ${totalPairs}`);
  
  // Process each cell in the matrix
  for (let row = startRow; row <= numCompanies + 1; row++) {
    for (let col = (row === startRow ? startCol : 2); col <= numCompanies + 1; col++) {
      // Save current progress
      scriptProperties.setProperties({
        'lastProcessedRow': row.toString(),
        'lastProcessedCol': col.toString()
      });
      
      // Skip if companies are the same
      if (row === col) {
        continue;
      }
      
      const company1 = matrixSheet.getRange(1, col).getValue();
      const company2 = matrixSheet.getRange(row, 1).getValue();
      
      // Check if cell is empty
      const existingValue = existingValues[row - 2][col - 2];
      if (isEmpty(existingValue)) {
        skippedPairs++;
        processedPairs++;
        showProgress(`Skipping ${company1} & ${company2} (no collaboration data) (${processedPairs}/${totalPairs})`);
        continue;
      }
      
      processedPairs++;
      showProgress(`Analyzing probability for ${company1} & ${company2} (${processedPairs}/${totalPairs}, Skipped: ${skippedPairs})`);
      
      try {
        const probability = analyzeCompanyPairProbability(
          company1, 
          company2, 
          companyInfo[company1], 
          companyInfo[company2],
          existingValue
        );
        
        // Color the cells based on probability
        const color = getProbabilityColor(probability.score);
        const cell = matrixSheet.getRange(row, col);
        const symmetricCell = matrixSheet.getRange(col, row);
        
        // Update both cells (matrix is symmetric)
        cell.setBackground(color)
            .setNote(`Probability Score: ${probability.score}%\n\nReasoning: ${probability.reasoning}`);
        symmetricCell.setBackground(color)
                    .setNote(`Probability Score: ${probability.score}%\n\nReasoning: ${probability.reasoning}`);
        
        successCount++;
        logToSheet('Probability Analysis', 'Success', `Completed ${company1} & ${company2} - Score: ${probability.score}%`);
        
        // Add delay to avoid rate limits
        Utilities.sleep(1000);
        
        // Check if we're approaching the time limit (5 minutes)
        if (processedPairs % 10 === 0) {  // Check every 10 pairs
          if (new Date().getTime() - START_TIME > 4.5 * 60 * 1000) {  // 4.5 minutes
            const pauseMessage = `Script paused due to time limit. Resume from: Column ${col}, Row ${row}`;
            logToSheet('Probability Analysis', 'Paused', pauseMessage);
            showProgress(pauseMessage);
            return;  // Exit the function
          }
        }
      } catch (error) {
        errorCount++;
        logToSheet('Probability Analysis', 'Error', `Failed ${company1} & ${company2}: ${error.message}`);
        continue;
      }
    }
  }
  
  // Clear progress when complete
  scriptProperties.deleteProperty('lastProcessedRow');
  scriptProperties.deleteProperty('lastProcessedCol');
  
  const finalMessage = `Probability Analysis Complete! Successful: ${successCount}, Skipped: ${skippedPairs}, Errors: ${errorCount}`;
  showProgress(finalMessage);
  logToSheet('Probability Analysis', 'Completed', finalMessage);
}

// Add this at the top of your file with other constants
const START_TIME = new Date().getTime();

// Add a function to reset progress if needed
function resetAnalysisProgress() {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteProperty('lastProcessedRow');
  scriptProperties.deleteProperty('lastProcessedCol');
  SpreadsheetApp.getActive().toast('Analysis progress has been reset', '✅ Reset Complete');
}

function analyzeCompanyPairProbability(company1, company2, info1, info2, existingCollaboration) {
  // Validate company information
  if (!info1 || !info2) {
    throw new Error(`Missing company information for ${!info1 ? company1 : company2}`);
  }

  try {
    const options = {
      'method': 'post',
      'headers': {
        'Authorization': `Bearer ${OPENAI_API_KEY}`,
        'Content-Type': 'application/json'
      },
      'payload': JSON.stringify({
        'model': 'o3-mini',
        'messages': [
          {
            'role': 'system',
            'content': 'You are a business and marketing strategy expert. Analyze collaboration potential between companies and provide: 1) A probability score (0-100), and 2) A brief explanation of the score. Keep responses concise.'
          },
          {
            'role': 'user',
            'content': `Analyze the probability of successful collaboration between these companies:

Company 1: ${company1}
Domain: ${info1.domain}
Business Overview: ${info1.businessOverview}
Target Audience: ${info1.targetAudience}
Products: ${info1.products}

Company 2: ${company2}
Domain: ${info2.domain}
Business Overview: ${info2.businessOverview}
Target Audience: ${info2.targetAudience}
Products: ${info2.products}

Proposed Collaboration:
${existingCollaboration}

Provide:
1. A probability score (0-100) for collaboration success
2. A brief explanation (max 100 words) of the score

Consider:
- Market alignment
- Product complementarity
- Target audience overlap
- Technical feasibility
- Potential conflicts
- Market timing

Format your response as:
Score: [number]
Reasoning: [explanation]`
          }
        ]
      }),
      'muteHttpExceptions': true
    };

    const response = UrlFetchApp.fetch(OPENAI_API_URL, options);
    const jsonResponse = JSON.parse(response.getContentText());
    
    // Log the raw response for debugging
    logToSheet('API Response', 'Debug', `Raw OpenAI response for ${company1} & ${company2}: ${JSON.stringify(jsonResponse)}`);
    
    if (!jsonResponse.choices || !jsonResponse.choices[0] || !jsonResponse.choices[0].message) {
      throw new Error('Invalid API response format');
    }
    
    const content = jsonResponse.choices[0].message.content;
    
    // Extract score and reasoning from the text response
    const scoreMatch = content.match(/Score:\s*(\d+)/i);
    const reasoningMatch = content.match(/Reasoning:\s*([^\n]+)/i);
    
    if (!scoreMatch || !reasoningMatch) {
      throw new Error('Could not extract score or reasoning from response');
    }
    
    const score = Math.min(100, Math.max(0, Number(scoreMatch[1])));
    const reasoning = reasoningMatch[1].trim();
    
    return {
      score: score,
      reasoning: reasoning
    };
  } catch (error) {
    logToSheet('API Error', 'Error', `Failed to analyze ${company1} & ${company2}: ${error.message}\nResponse content: ${jsonResponse?.choices?.[0]?.message?.content || 'No content'}`);
    throw new Error(`Failed to analyze probability: ${error.message}`);
  }
}

function getProbabilityColor(score) {
  if (score >= 70) {
    return '#b7e1cd'; // Light green
  } else if (score >= 40) {
    return '#fff2cc'; // Light yellow
  } else {
    return '#f4c7c3'; // Light red
  }
}
