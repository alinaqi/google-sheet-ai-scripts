// Add OpenAI configuration
const OPENAI_API_KEY = 'your_key';
const PERPLEXITY_API_KEY = 'your_key';
const OPENAI_MODEL = 'gpt-4o-mini';
const BATCH_SIZE = 5; // Number of profiles to process in each batch

function logToSheet(message) {
  try {
    console.log(message); // Log to Apps Script console
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      console.error('Could not get active spreadsheet');
      return;
    }
    
    let logSheet = ss.getSheetByName('Processing Logs');
    if (!logSheet) {
      console.log('Creating new log sheet...');
      logSheet = ss.insertSheet('Processing Logs');
      logSheet.appendRow(['Timestamp', 'Message']);
      console.log('Log sheet created');
    }
    
    const timestamp = new Date().toISOString();
    logSheet.appendRow([timestamp, message]);
    
  } catch (error) {
    console.error('Error in logToSheet:', error.message, '\nStack:', error.stack);
    // Try one more time to log the error to the sheet
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const logSheet = ss.getSheetByName('Processing Logs');
      if (logSheet) {
        logSheet.appendRow([new Date().toISOString(), `Logging error: ${error.message}`]);
      }
    } catch (e) {
      console.error('Failed to log error to sheet:', e.message);
    }
  }
}

function processBatch() {
  try {
    logToSheet('=============== STARTING NEW BATCH ===============');
    logToSheet('Starting batch processing...');
    
    const sheet = SpreadsheetApp.getActiveSheet();
    logToSheet(`Active sheet name: ${sheet.getName()}`);
    
    const lastRow = sheet.getLastRow();
    logToSheet(`Last row in sheet: ${lastRow}`);
    
    // Set column headers if not present
    const headers = ['Name', 'LinkedIn URL', 'Title', 'About Profile', 'Email Subject', 'Email Content', 'Batch Number', 'Status'];
    logToSheet('Checking headers...');
    
    const firstRow = sheet.getRange('A1:H1').getValues()[0];
    logToSheet(`Current headers: ${JSON.stringify(firstRow)}`);
    
    if (firstRow[0] === '') {
      logToSheet('Headers missing, adding them now...');
      sheet.getRange('A1:H1').setValues([headers]);
      logToSheet('Added headers to sheet');
    }
    
    // Find the next batch to process
    let currentBatchNumber = 1;
    logToSheet('Starting batch number calculation...');
    
    // Handle case when sheet only has headers
    if (lastRow > 1) {
      logToSheet(`Getting batch numbers from rows 2 to ${lastRow}`);
      const batchRange = sheet.getRange(`G2:G${lastRow}`);
      const batchValues = batchRange.getValues();
      logToSheet(`Raw batch values: ${JSON.stringify(batchValues)}`);
      
      const batchNumbers = batchValues
        .flat()
        .map(val => {
          logToSheet(`Processing batch value: ${val}, type: ${typeof val}`);
          // Convert to number, handling null, undefined, and empty strings
          if (!val || val === '') {
            logToSheet('Empty or null batch value, returning 0');
            return 0;
          }
          const num = Number(val);
          if (isNaN(num)) {
            logToSheet(`Invalid number value: ${val}, returning 0`);
            return 0;
          }
          logToSheet(`Valid batch number found: ${num}`);
          return num;
        });
      
      logToSheet(`Processed batch numbers: ${JSON.stringify(batchNumbers)}`);
      
      if (batchNumbers.length > 0) {
        const maxBatchNumber = Math.max(...batchNumbers);
        logToSheet(`Maximum batch number found: ${maxBatchNumber}`);
        currentBatchNumber = maxBatchNumber + 1;
      } else {
        logToSheet('No valid batch numbers found, using default: 1');
      }
    } else {
      logToSheet('Sheet only has headers, using default batch number: 1');
    }
    
    logToSheet(`Final calculated batch number: ${currentBatchNumber}`);
    
    // First process existing profiles that don't have emails
    logToSheet('Starting to process existing profiles...');
    let profilesProcessed = processExistingProfiles(sheet, currentBatchNumber);
    logToSheet(`Processed ${profilesProcessed} existing profiles`);
    
    // Then discover and process new profiles if needed
    if (profilesProcessed < BATCH_SIZE) {
      logToSheet(`Need to discover more profiles, ${BATCH_SIZE - profilesProcessed} slots remaining`);
      discoverNewProfiles(sheet, currentBatchNumber, BATCH_SIZE - profilesProcessed);
    } else {
      logToSheet('Batch is full, skipping profile discovery');
    }
    
    logToSheet('=============== BATCH PROCESSING COMPLETE ===============');
    SpreadsheetApp.getActive().toast('Batch processing complete!', 'Status');
    
  } catch (error) {
    const errorMessage = `Fatal error in processBatch: ${error.message}\nStack trace: ${error.stack}`;
    logToSheet('=============== ERROR ===============');
    logToSheet(errorMessage);
    logToSheet('=============== END ERROR ===============');
    SpreadsheetApp.getActive().toast(errorMessage, 'Error');
  }
}

function processExistingProfiles(sheet, batchNumber) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    logToSheet('No profiles to process - sheet is empty except for headers');
    return 0;
  }
  
  let profilesProcessed = 0;
  logToSheet(`Starting to process existing profiles for batch ${batchNumber}`);
  
  // Get all rows without emails but with names and LinkedIn URLs
  for (let row = 2; row <= lastRow && profilesProcessed < BATCH_SIZE; row++) {
    const range = sheet.getRange(`A${row}:H${row}`);
    const values = range.getValues()[0];
    const [name, linkedinUrl, title, about, subject, content, batch, status] = values;
    
    logToSheet(`Row ${row}: Name=${name}, LinkedIn=${linkedinUrl}, Content=${!!content}, Batch=${batch}`);
    
    // Skip if already processed or missing required data
    if (!name || !linkedinUrl) {
      logToSheet(`Skipping row ${row}: Missing name or LinkedIn URL`);
      continue;
    }
    
    if (content || batch) {
      logToSheet(`Skipping row ${row}: Already processed`);
      continue;
    }
    
    logToSheet(`Processing existing profile: ${name}`);
    
    try {
      const profileData = analyzeProfile(name, linkedinUrl);
      
      // Update sheet with results
      range.setValues([[
        name,
        linkedinUrl,
        profileData.title,
        profileData.aboutProfile,
        profileData.emailSubject,
        profileData.emailContent,
        batchNumber,
        'Completed'
      ]]);
      
      profilesProcessed++;
      logToSheet(`Successfully processed profile for ${name}`);
      
      // Wait between API calls
      Utilities.sleep(2000);
      
    } catch (error) {
      logToSheet(`Error processing ${name}'s profile: ${error.message}`);
      range.getCell(1, 8).setValue('Error: ' + error.message);
    }
  }
  
  return profilesProcessed;
}

function discoverNewProfiles(sheet, batchNumber, remainingSlots) {
  logToSheet(`Discovering new profiles, slots remaining: ${remainingSlots}`);
  
  // Get all processed profiles for discovery
  const lastRow = sheet.getLastRow();
  const processedProfiles = sheet.getRange(`A2:B${lastRow}`).getValues()
    .filter(row => row[0] && row[1])
    .slice(-5); // Use last 5 processed profiles for discovery
  
  const analyzedUrls = new Set(
    sheet.getRange(`B2:B${lastRow}`).getValues()
      .map(row => row[0].toString().toLowerCase())
      .filter(url => url !== '')
  );
  
  for (const [name, linkedinUrl] of processedProfiles) {
    if (remainingSlots <= 0) break;
    
    try {
      logToSheet(`Discovering related profiles for ${name}`);
      const relatedProfiles = discoverRelatedProfiles(name, linkedinUrl);
      
      if (relatedProfiles && relatedProfiles.length > 0) {
        const newProfiles = relatedProfiles.filter(profile => 
          !analyzedUrls.has(profile.linkedinUrl.toLowerCase())
        ).slice(0, remainingSlots);
        
        if (newProfiles.length > 0) {
          const lastRowAfterAnalysis = sheet.getLastRow();
          const newRows = newProfiles.map(profile => [
            profile.name,
            profile.linkedinUrl,
            '', // title
            '', // about
            '', // subject
            '', // content
            batchNumber,
            'Pending'
          ]);
          
          sheet.getRange(lastRowAfterAnalysis + 1, 1, newRows.length, 8).setValues(newRows);
          
          remainingSlots -= newProfiles.length;
          logToSheet(`Added ${newProfiles.length} new profiles from ${name}'s network`);
          
          // Process the newly added profiles
          for (const profile of newProfiles) {
            const row = sheet.getLastRow();
            const profileData = analyzeProfile(profile.name, profile.linkedinUrl);
            
            sheet.getRange(row, 1, 1, 8).setValues([[
              profile.name,
              profile.linkedinUrl,
              profileData.title,
              profileData.aboutProfile,
              profileData.emailSubject,
              profileData.emailContent,
              batchNumber,
              'Completed'
            ]]);
            
            logToSheet(`Processed new profile: ${profile.name}`);
            Utilities.sleep(2000);
          }
        }
      }
    } catch (error) {
      logToSheet(`Error discovering profiles from ${name}: ${error.message}`);
    }
  }
}

function analyzeProfile(name, linkedinUrl) {
  logToSheet(`Analyzing profile for ${name}`);
  
  try {
    // Generate personalized content with OpenAI
    const analysisPrompt = `
      Research this LinkedIn profile: ${name} (${linkedinUrl})
      
      Create a personalized outreach about Protaige, an AI-driven marketing automation platform. Key benefits:
      - Complete brand voice and story capture/management
      - Persona creation and management
      - End-to-end campaign creation (strategy to content)

      Keep the email funny and quirky.
      
      YOU MUST RESPOND WITH VALID JSON ONLY. Do not include any explanatory text.
      The JSON must have exactly these fields:
      {
        "title": "their current title",
        "aboutProfile": "key insights about their background (max 100 words)",
        "emailSubject": "compelling personalized subject line",
        "emailContent": "professional personalized email (2-3 paragraphs)"
      }
    `;
    
    const profileData = callOpenAI(analysisPrompt);
    logToSheet('Successfully analyzed profile');
    return profileData;
  } catch (error) {
    logToSheet(`Analysis failed: ${error.message}`);
    throw error;
  }
}

function discoverRelatedProfiles(name, linkedinUrl) {
  logToSheet(`Discovering related profiles for ${name}`);
  
  const discoveryPrompt = `
    Research this LinkedIn profile: ${name} (${linkedinUrl})
    Find 5 similar profiles in their network who might be interested in AI marketing automation. Focus on marketing directors, CMOs etc. 
    
    Format the response exactly like this:
    PROFILES_START
    name: Full Name
    linkedin: Profile URL
    PROFILES_END
  `;
  
  try {
    const research = callPerplexityAPI(discoveryPrompt);
    
    // Extract related profiles
    const profilesMatch = research.match(/PROFILES_START([\s\S]*?)PROFILES_END/);
    if (profilesMatch) {
      const profilesText = profilesMatch[1];
      const profileEntries = profilesText.split(/(?=name:)/);
      
      const relatedProfiles = profileEntries
        .map(entry => {
          const nameMatch = entry.match(/name:\s*([^\n]+)/);
          const linkedinMatch = entry.match(/linkedin:\s*([^\n]+)/);
          
          if (nameMatch && linkedinMatch) {
            return {
              name: nameMatch[1].trim(),
              linkedinUrl: linkedinMatch[1].trim()
            };
          }
          return null;
        })
        .filter(profile => profile !== null);
      
      logToSheet(`Found ${relatedProfiles.length} related profiles`);
      return relatedProfiles;
    }
    
    return [];
  } catch (error) {
    logToSheet(`Profile discovery failed: ${error.message}`);
    return [];
  }
}

function callPerplexityAPI(prompt) {
  logToSheet('Calling Perplexity API...');
  
  const url = 'https://api.perplexity.ai/chat/completions';
  const payload = {
    model: 'llama-3.1-sonar-large-128k-chat',
    messages: [
      {
        role: 'system',
        content: 'You are a thorough professional profile researcher. Focus on relevant experience and connections.'
      },
      {
        role: 'user',
        content: prompt
      }
    ]
  };

  const options = {
    method: 'post',
    headers: {
      'Authorization': `Bearer ${PERPLEXITY_API_KEY}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    if (responseCode !== 200) {
      throw new Error(`API error (${responseCode}): ${responseText}`);
    }
    
    const jsonResponse = JSON.parse(responseText);
    return jsonResponse.choices[0].message.content;
  } catch (error) {
    logToSheet(`Perplexity API error: ${error.message}`);
    throw error;
  }
}

function callOpenAI(prompt) {
  logToSheet('Calling OpenAI API...');
  
  const url = 'https://api.openai.com/v1/chat/completions';
  const payload = {
    model: 'gpt-4o',
    response_format: { "type": "json_object" },
    messages: [
      {
        role: 'system',
        content: 'You are an expert at creating personalized B2B outreach content. Focus on value proposition and relevant experience.'
      },
      {
        role: 'user',
        content: prompt
      }
    ]
  };

  const options = {
    method: 'post',
    headers: {
      'Authorization': `Bearer ${OPENAI_API_KEY}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    if (responseCode !== 200) {
      throw new Error(`API error (${responseCode}): ${responseText}`);
    }
    
    const jsonResponse = JSON.parse(responseText);
    let content = jsonResponse.choices[0].message.content;
    content = content.trim();
    
    try {
      const parsedContent = JSON.parse(content);
      
      // Validate required fields
      const requiredFields = ['title', 'aboutProfile', 'emailSubject', 'emailContent'];
      const missingFields = requiredFields.filter(field => !parsedContent[field]);
      
      if (missingFields.length > 0) {
        throw new Error(`Missing required fields: ${missingFields.join(', ')}`);
      }
      
      return parsedContent;
    } catch (parseError) {
      logToSheet(`Error parsing OpenAI response as JSON: ${content}`);
      throw parseError;
    }
  } catch (error) {
    logToSheet(`OpenAI API error: ${error.message}`);
    throw error;
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('LinkedIn Analysis')
    .addItem('Process Next Batch', 'processBatch')
    .addToUi();
}
