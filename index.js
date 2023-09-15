const xlsx = require('xlsx');
const axios = require('axios');
const axiosRetry = require('axios-retry');
const dotenv = require('dotenv');

dotenv.config();

const ACCESS_TOKEN = process.env.API_KEY;

console.log('API_KEY: ', ACCESS_TOKEN);

// Configure axios-retry
axiosRetry(axios, { retries: 3, retryDelay: axiosRetry.exponentialDelay });

const workbook = xlsx.readFile('./ZaeK_Schleswig_Holstein_Parsing_14_09_2023_01_34.xlsx');
const sheet_name_list = workbook.SheetNames;
let jsonData = xlsx.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);

// Remove duplicates based on the email property
const contactsData = Array.from(new Set(jsonData.map((a) => a['E-Mail']))).map((email) => {
  return jsonData.find((a) => a['E-Mail'] === email);
});
// Remove duplicates based on the company name property
const dealsData = Array.from(new Set(jsonData.map((a) => a['Company Name']))).map((company) => {
  return jsonData.find((a) => a['Company Name'] === company);
});

// Log duplicates based on the email property
const logDuplicateEmails = () => {
  let seenEmails = {};
  jsonData.forEach((a) => {
    if (a['E-Mail'] in seenEmails) {
      if (seenEmails[a['E-Mail']] === 1) {
        // Only log the first time a duplicate is found
        console.log('Duplicate Email: ', a['E-Mail']);
      }
      seenEmails[a['E-Mail']]++;
    } else {
      seenEmails[a['E-Mail']] = 1;
    }
  });
};

// Log duplicates based on the company name property
const logDuplicateCompanies = () => {
  let seenCompanies = {};
  jsonData.forEach((a) => {
    if (a['Company Name'] in seenCompanies) {
      if (seenCompanies[a['Company Name']] === 1) {
        // Only log the first time a duplicate is found
        console.log('Duplicate Company Name: ', a['Company Name']);
      }
      seenCompanies[a['Company Name']]++;
    } else {
      seenCompanies[a['Company Name']] = 1;
    }
  });
};

// Call the new functions
// logDuplicateEmails();
// logDuplicateCompanies();

// console.log('Contacts: ', contactsData);
// console.log('Deals: ', dealsData);

// Function to validate email address
// Function to validate email address
function validateEmail(email) {
  const re = /^[\w\.-]+@[^\s@]+\.[^\s@]{2,}$/;
  if (!re.test(email)) return false;
  const domain = email.split('@')[1];
  if (
    domain.startsWith('.') ||
    domain.endsWith('.') ||
    domain.includes('..') ||
    domain.includes('.-') ||
    domain.includes('-.')
  ) {
    return false;
  }
  return true;
}

async function createContactsBatch(accessToken, contacts) {
  const headers = {
    Authorization: `Bearer ${accessToken}`,
    'Content-Type': 'application/json',
  };

  const apiUrl = `https://api.hubapi.com/crm/v3/objects/contacts`;

  for (let contact of contacts) {
    try {
      const response = await axios.post(
        apiUrl,
        {
          properties: {
            company: contact['Company Name'],
            email: contact['E-Mail'],
            firstname: contact['Company Name'],
            website: contact['Website'],
            zip: contact['Zip Code'],
            house_number: parseInt(contact['House number']),
            city: contact['City'],
            street: contact['Street'],
            fax: contact['Fax Number'],
            phone: contact['Phone Number'],
          },
        },
        {
          headers: headers,
        }
      );

      console.log(`Created contact with ID: ${response.data.id}`);
    } catch (error) {
      if (error.response && error.response.status === 409) {
        console.error('A contact with the same ID already exists. Skipping...');
        continue; // Skip this contact and move on to the next one
      } else if (
        error.response &&
        error.response.data &&
        typeof error.response.data?.message === 'string' &&
        error.response.data?.message?.includes('INVALID_EMAIL')
      ) {
        // Handle invalid email error
        console.error('Invalid email for contact:', contact['E-Mail'], 'Skipping...');
        continue; // Skip this contact with invalid email and move on to the next one
      } else {
        console.error('Error creating contacts:', error);
        throw error; // Re-throw the error for axios-retry to catch and retry
      }
    }

    // Delay next request for rate limiting (100 requests per 10 seconds)
    await new Promise((resolve) => setTimeout(resolve, 100));
  }
}

async function createDealsBatch(accessToken, deals) {
  const headers = {
    Authorization: `Bearer ${accessToken}`,
    'Content-Type': 'application/json',
  };

  const apiUrl = `https://api.hubapi.com/crm/v3/objects/deals/batch/create`;

  try {
    const response = await axios.post(
      apiUrl,
      {
        inputs: deals.map((deal) => ({
          properties: {
            dealname: deal['Company Name'], // Using company name as deal name
            pipeline: 'default',
          },
        })),
      },
      {
        headers: headers,
      }
    );

    return response.data;
  } catch (error) {
    console.error('Error creating deals batch:', error);
  }
}

async function runCreateDeals() {
  // Split jsonData into batches of 100 (HubSpot's maximum batch size)
  const dealBatches = [];
  for (let i = 0; i < dealsData.length; i += 100) {
    dealBatches.push(dealsData.slice(i, i + 100));
  }

  // Create deals
  for (let i = 0; i < dealBatches.length; i++) {
    console.log(`Creating deals batch ${i + 1}...`);
    const deals = dealBatches[i].filter((contact) => validateEmail(contact['E-Mail'])); // Only valid emails;
    await createDealsBatch(ACCESS_TOKEN, deals);

    // Handle rate limiting (100 requests per 10 seconds)
    await new Promise((resolve) => setTimeout(resolve, 100));
  }
}

const runCreateContacts = async () => {
  // Split jsonData into batches of 100 (HubSpot's maximum batch size)
  const contactBatches = [];
  for (let i = 0; i < contactsData.length; i += 100) {
    contactBatches.push(contactsData.slice(i, i + 100));
  }

  // Create contacts
  for (let i = 0; i < contactBatches.length; i++) {
    console.log(`Creating contacts batch ${i + 1}...`);
    const contacts = contactBatches[i].filter((contact) => validateEmail(contact['E-Mail'])); // Only valid emails
    await createContactsBatch(ACCESS_TOKEN, contacts);

    // Handle rate limiting (100 requests per 10 seconds)
    await new Promise((resolve) => setTimeout(resolve, 100));
  }
};
runCreateContacts();
// runCreateDeals();
