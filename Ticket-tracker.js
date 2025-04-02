// Zendesk Ticket Tracker for Google Sheets
// This script fetches Zendesk tickets with a specific tag and displays them in a Google Sheet

// Your Zendesk credentials - you'll need to set these
const ZENDESK_DOMAIN = 'reallycoolsupport.zendesk.com'; // Replace with your Zendesk domain
const ZENDESK_EMAIL = 'Your@email.com'; // Replace with your Zendesk  email
const ZENDESK_API_TOKEN = '12324353421aSfdgfsgrwasd'; // Replace with your API token

// Main function to fetch tickets and update the sheet
function updateTicketsInSheet() {
  try {
    // Get the active sheet
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Read the tag from cell A1
    const tag = sheet.getRange('A1').getValue();
    
    if (!tag || tag.trim() === '') {
      sheet.getRange('A3').setValue('Error: Please enter a tag in cell A1');
      return;
    }
    
    // Set up headers if they don't exist
    setupHeaders(sheet);
    
    // Clear existing ticket data (but keep the headers and tag)
    clearExistingData(sheet);
    
    // Fetch tickets with the specified tag
    const tickets = fetchTicketsWithTag(tag);
    
    if (!tickets || tickets.length === 0) {
      sheet.getRange('A3').setValue(`No tickets found with tag: ${tag}`);
      return;
    }
    
    // Update the sheet with the ticket data
    displayTickets(sheet, tickets);
    
    // Update last refreshed timestamp
    const now = new Date();
    sheet.getRange('C1').setValue(`Last updated: ${now.toLocaleString()}`);
    
  } catch (error) {
    console.error('Error updating tickets:', error);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    sheet.getRange('A3').setValue(`Error: ${error.message}`);
  }
}

// Function to set up headers in the sheet
function setupHeaders(sheet) {
  const headers = ['Ticket ID', 'Subject', 'Status', 'Priority', 'Requester', 'Assigned To', 'Created Date', 'Updated Date', 'Link'];
  
  // Check if headers already exist
  const existingHeaders = sheet.getRange(2, 1, 1, headers.length).getValues()[0];
  
  // If headers don't match, update them
  if (existingHeaders.join('') !== headers.join('')) {
    sheet.getRange(2, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(2, 1, 1, headers.length).setFontWeight('bold');
  }
}

// Function to clear existing ticket data
function clearExistingData(sheet) {
  // Get the last row with data
  const lastRow = sheet.getLastRow();
  
  // If we have data beyond the header row, clear it
  if (lastRow > 2) {
    sheet.getRange(3, 1, lastRow - 2, sheet.getLastColumn()).clearContent();
  }
}

// Function to fetch tickets with a specific tag from Zendesk
function fetchTicketsWithTag(tag) {
  // Create the search URL with pagination (100 tickets per page)
  const searchUrl = `https://${ZENDESK_DOMAIN}/api/v2/search.json?query=tags:${tag}+type:ticket&per_page=100`;
  
  // Set up the request options
  const options = {
    method: 'get',
    headers: {
      'Authorization': 'Basic ' + Utilities.base64Encode(`${ZENDESK_EMAIL}/token:${ZENDESK_API_TOKEN}`)
    },
    muteHttpExceptions: true
  };
  
  // Make the request to Zendesk API
  const response = UrlFetchApp.fetch(searchUrl, options);
  const responseCode = response.getResponseCode();
  
  // Check if the request was successful
  if (responseCode !== 200) {
    console.error(`API Error: ${responseCode}`);
    throw new Error(`Zendesk API Error: ${responseCode}`);
  }
  
  // Parse the response
  const responseData = JSON.parse(response.getContentText());
  
  return responseData.results;
}

// Function to display tickets in the sheet
function displayTickets(sheet, tickets) {
  // Prepare data for batch update
  const ticketData = tickets.map(ticket => [
    ticket.id,
    ticket.subject,
    ticket.status,
    ticket.priority || 'None',
    ticket.requester_id, // Ideally should resolve to name, would need another API call
    ticket.assignee_id || 'Unassigned', // Same as above
    new Date(ticket.created_at).toLocaleString(),
    new Date(ticket.updated_at).toLocaleString(),
    `https://minutemediasupport.zendesk.com/agent/ticket/${ticket.id}`
  ]);
  
  // Update the sheet with all ticket data at once
  if (ticketData.length > 0) {
    sheet.getRange(3, 1, ticketData.length, ticketData[0].length).setValues(ticketData);
  }
  
  // Add a note about the number of tickets found
  sheet.getRange('A1').setValue(sheet.getRange('A1').getValue());
  sheet.getRange('B1').setValue(`Found ${tickets.length} tickets`);
}

// Function to create a trigger to run the script every 10 minutes
function createTimeTrigger() {
  // Delete any existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'updateTicketsInSheet') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // Create a new trigger to run every 10 minutes
  ScriptApp.newTrigger('updateTicketsInSheet')
    .timeBased()
    .everyMinutes(10)
    .create();
    
  // Show confirmation to the user
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.getRange('D1').setValue('Auto-refresh: Every 10 minutes');
}

// Function to add a menu to the spreadsheet
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Zendesk Tickets')
    .addItem('Update Tickets Now', 'updateTicketsInSheet')
    .addItem('Set Auto-refresh (10 min)', 'createTimeTrigger')
    .addToUi();
}

