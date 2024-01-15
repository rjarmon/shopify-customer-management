/*------------------------
Libraries
------------------------*/

// Import necessary libraries
require('dotenv').config();
const express = require('express');
const multer = require('multer');
const axios = require('axios');
const FormData = require('form-data');
const graph = require('@microsoft/microsoft-graph-client');
const fs = require('fs');

// Additional utility libraries
const path = require('path');
const { TokenCredentialAuthenticationProvider } = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');
const { ClientSecretCredential } = require('@azure/identity');

/*------------------------
Additional variables
------------------------*/

// Regular expression for URL validation
const urlRegex = /^(ftp|http|https):\/\/[^ "]+$/;

/*------------------------
Express Setup
------------------------*/

// Define server port
const SERVER_PORT = 3000;

// Create an instance of Express
const app = express();

// Configure multer for file uploads
const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

// Parse incoming request bodies
const bodyParser = require('body-parser');
app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());

/*------------------------
Microsoft Graph Client Setup
------------------------*/

// Initialize Microsoft Graph client
const { AZURE_APP_TENANT_ID, AZURE_APP_CLIENT_ID, AZURE_APP_CLIENT_SECRET_VALUE, ADMIN_ACCESS_TOKEN } = process.env;

// Check if required environment variables are present
if (!AZURE_APP_TENANT_ID || !AZURE_APP_CLIENT_ID || !AZURE_APP_CLIENT_SECRET_VALUE || !ADMIN_ACCESS_TOKEN) {
  console.error('One or more required environment variables are missing.');
  process.exit(1);
}

// Set up authentication credentials
const credential = new ClientSecretCredential(AZURE_APP_TENANT_ID, AZURE_APP_CLIENT_ID, AZURE_APP_CLIENT_SECRET_VALUE);
const authProvider = new TokenCredentialAuthenticationProvider(credential, { scopes: ['https://graph.microsoft.com/.default'] });
const msGraphClient = graph.Client.initWithMiddleware({
  debugLogging: true,
  authProvider: authProvider
});

/*------------------------
Shopify API Configuration
------------------------*/

const adminAccessToken = process.env.ADMIN_ACCESS_TOKEN;
const shopifyAPIUrl = 'https://[REDACTED].myshopify.com/admin/api/2023-04/graphql.json';

/*------------------------
Utility Functions
------------------------*/

/**
 * Retrieves the MIME type based on the file extension.
 * @param {string} extension - The file extension.
 * @returns {string|null} - The MIME type or null for unknown file types.
 */
function getMimeType(extension) {
  switch (extension.toLowerCase()) {
    case '.png':
      return 'image/png';
    case '.jpg':
    case '.jpeg':
      return 'image/jpeg';
    case '.pdf':
      return 'application/pdf';
    default:
      return null; // Return null for unknown file type
  }
}

/*------------------------
File Upload Endpoint
------------------------*/

app.post('/upload', upload.single('tax_exempt_form'), async (req, res) => {
  try {
    // Extract customer information and uploaded file from the request
    const { body: { customerId, customerCompany }, file } = req;

    // Set a cookie for user feedback
    res.cookie('uploadMessage', 'Your file was uploaded successfully!', { maxAge: 3000 });

    /*------------------------
    Create staged upload.
    ---
    Shopify sets up temporary file targets in AWS s3 buckets so we can host file data (images, videos, etc).
    If you already have a public url for your image file then you can skip this step and pass the url directly to the create file endpoint.
    But in many cases, you'll want to first stage the upload on s3. Cases include generating a specific name for your image, uploading the image from a private server, etc.
    ------------------------*/

    // Query
    const stagedUploadsQuery = `mutation stagedUploadsCreate($input: [StagedUploadInput!]!) {
      stagedUploadsCreate(input: $input) {
        stagedTargets {
          resourceUrl
          url
          parameters {
            name
            value
          }
        }
        userErrors {
          field
          message
        }
      }
    }`;

    // Get the file extension using path.extname
    const fileExtension = path.extname(file.originalname);

    // Determine the MIME type based on the file extension
    const mimeType = getMimeType(fileExtension);

    // Check if the MIME type is valid
    if (!mimeType) {
      const errorMessage = 'Invalid file type: Please submit a JPG, PNG, or PDF.';
      // Pass the error message to the customer on the webpage or handle it as needed
      // return res.status(400).send(errorMessage);
      // Pass the error message to the webpage using an alert
      return res.send(`<script>alert('${errorMessage}'); window.location.href = 'https://https://[REDACTED].myshopify.com/account';</script>`);
    }

    // Update the stagedUploadsVariables
    const stagedUploadsVariables = {
      input: {
        filename: `${customerCompany} Tax Exempt Form${fileExtension}`,
        httpMethod: 'POST',
        mimeType: mimeType,
        resource: 'FILE',
      }
    }

    // Result
    const stagedUploadsQueryResult = await axios.post(
      shopifyAPIUrl,
      {
        query: stagedUploadsQuery,
        variables: stagedUploadsVariables,
      },
      {
        headers: {
          'X-Shopify-Access-Token': ADMIN_ACCESS_TOKEN,
          'Content-Security-Policy': 'frame-ancestors https://[REDACTED].myshopify.com https://admin.shopify.com',
        },
      }
    );

    // Save the target info.
    const target = stagedUploadsQueryResult.data.data.stagedUploadsCreate.stagedTargets[0];
    const params = target.parameters; // Parameters contain all the sensitive info we'll need to interact with the AWS bucket.
    const url = target.url; // This is the URL used to post data to AWS. It's a generic s3 URL that, when combined with the params, sends data to the correct place.
    const resourceUrl = target.resourceUrl; // This is the specific URL that will contain the image data after the file has been uploaded to the AWS staged target.

    /*------------------------
    Post to temp target.
    ---
    A temp target is a URL hosted on Shopify's AWS servers.
    ------------------------*/
    // Generate a form, add the necessary params, and append the file.
    // Must use the FormData library to create form data via the server.
    const form = new FormData();

    // Add each of the params received from Shopify to the form. This will ensure the AJAX request has the proper permissions and s3 location data.
    params.forEach(({ name, value }) => {
      form.append(name, value);
    });

    // Add the file to the form.
    form.append('file', Buffer.from(file.buffer));

    // Post the file data to Shopify's AWS s3 bucket. After posting, the resource URL can be used to create the file in Shopify.
    await axios.post(url, form, {
      headers: {
        ...form.getHeaders(), // Pass the headers generated by FormData library. It will contain content-type: multipart/formdata. It's necessary to specify this when posting to AWS.
        // 'Content-Length': fileSize + 5000, // AWS requires content length to be included in the headers. This may not be automatically passed, so it needs to be specified. 5000 is added to ensure the upload works, or else there will be an error saying the data isn't formatted properly.
      },
    }).catch(error => {
      console.error('Error downloading file: ', error);
      return res.status(500).send('Internal Server Error');
    });

    /*------------------------
    Create the file.
    Now that the file is prepared and accessible on the staged target, use the resource URL from AWS to create the file.
    ------------------------*/
    // Query
    const createFileQuery = `mutation fileCreate($files: [FileCreateInput!]!) {
      fileCreate(files: $files) {
        files {
          fileStatus
          createdAt
          ... on GenericFile {
            id
            url
          }
          ... on MediaImage {
            id
            image {
              id
              originalSrc: url
            }
          }
        }
        userErrors {
          field
          message
        }
      }
    }`;

    // Variables
    const createFileVariables = {
      files:
      {
        contentType: mimeType === 'application/pdf' ? 'FILE' : 'IMAGE', // Change contentType to 'FILE' if the file is a PDF, otherwise 'IMAGE'
        originalSource: resourceUrl, // Pass the resource URL generated above as the original source. Shopify will do the work of parsing that URL and adding it to files.
      },
    };

    // Finally post the file to Shopify. It should appear in Settings > Files.
    const createFileQueryResult = await axios.post(
      shopifyAPIUrl,
      {
        query: createFileQuery,
        variables: createFileVariables,
      },
      {
        headers: {
          'X-Shopify-Access-Token': ADMIN_ACCESS_TOKEN,
        },
      }
    );

    const fileId = createFileQueryResult.data.data.fileCreate.files[0].id;

    // Update the customer's metafield with the link to their file
    const setMetafieldQuery = `mutation metafieldsSet($metafields: [MetafieldsSetInput!]!) {
      metafieldsSet(metafields: $metafields) {
        metafields {
          key
          namespace
          value
          createdAt
          updatedAt
        }
        userErrors {
          field
          message
          code
        }
      }
    }`;

    // Variables
    const setMetafieldVariables = {
      metafields: [
        {
          key: 'tax_exempt_form',
          namespace: 'tax_exempt_forms',
          ownerId: `gid://shopify/Customer/${customerId}`,
          type: 'file_reference',
          value: fileId
        }
      ]
    };

    // Query
    const setMetafieldQueryResult = await axios.post(
      shopifyAPIUrl,
      {
        query: setMetafieldQuery,
        variables: setMetafieldVariables,
      },
      {
        headers: {
          'X-Shopify-Access-Token': ADMIN_ACCESS_TOKEN,
        },
      }
    ).catch(error => {
      console.error('Error creating metafields: ', error)
      return res.status(500).send('Internal Server Error');
    });

    // Log success or failure
    if (setMetafieldQueryResult) {
      const taxExemptDocumentUploadEmail = {
        message: {
          subject: `New tax exempt document uploaded for ${customerCompany}`,
          body: {
            contentType: 'html',
            content: `${customerCompany} has uploaded a new tax exempt document\n<a href="https://[REDACTED].myshopify.com/admin/customers/${customerId}">Click here to review the document and approve the customer</a>.`
          },
          toRecipients: [
            {
              emailAddress: {
                address: "[REDACTED]"
              }
            }
          ],
          "from": {
            "emailAddress": {
              "address": "[REDACTED]"
            }
          }
        },
        saveToSentItems: 'true'
      };
    
      await msGraphClient.api('/users/[REDACTED]/sendMail').post(taxExemptDocumentUploadEmail);

    } else {
      console.error('There was an error uploading the document')
      return res.status(500).send('Internal Server Error');
    }

    // Redirect the user to a success page
    res.redirect('https://[REDACTED].myshopify.com/account');
  } catch (error) {
    console.error('Error processing file upload:', error);
    return res.status(500).send('Internal Server Error');
  }
});

/*------------------------
Customer Registration Endpoint
------------------------*/

app.post('/register', async (req, res) => {
  try {
    // Destructuring request properties
    const { body: { companyName, firstName, lastName, email, companyWebsite, phoneNumber } } = req;

    // Check and format the company website URL
    let website = companyWebsite;
    if (!companyWebsite.startsWith("http://") && !companyWebsite.startsWith("https://")) {
      website = "http://" + companyWebsite;
    }

    // Function to convert phone number to E.164 format
    function convertToE164(phoneNumber) {
      // Remove all non-digit characters from the phone number
      const digitsOnly = phoneNumber.replace(/\D/g, '');
    
      // Check if the number starts with a country code
      if (digitsOnly.startsWith('1') && digitsOnly.length === 11) {
        // If it starts with 1 and has 11 digits, assume it's already in E.164 format
        return `+${digitsOnly}`;
      } else if (digitsOnly.length === 10) {
        // If it has 10 digits, assume it's a US number and prepend the country code
        return `+1${digitsOnly}`;
      } else {
        // If it has a different number of digits, or doesn't match any of the above conditions, return null or handle the error as needed
        return null;
      }
    }

    // Convert phone number to E.164 format
    const phone = convertToE164({ phoneNumber });

    // Create the user in Shopify using the GraphQL Admin API
    const createCustomerQuery = `mutation createCustomerMetafields($input: CustomerInput!) {
      customerCreate(input: $input) {
        customer {
          id
          email
          phone
          taxExempt
          acceptsMarketing
          firstName
          lastName
          addresses {
            company
            firstName
            lastName
            phone
          }
          smsMarketingConsent {
            marketingState
            marketingOptInLevel
          }
          metafields(first: 3) {
            edges {
              node {
                key
                namespace
                type
                value
              }
            }
          }
        }
        userErrors {
          field
          message
        }
      }
    }`;

    // Set the customer creation variables
    const createCustomerVariables = {
      input: {
        email: email,
        phone: phone,
        firstName: firstName,
        lastName: lastName,
        addresses: [
          {
            company: companyName,
            firstName: firstName,
            lastName: lastName,
            phone: phone
          }
        ],
        metafields: [
          {
            key: 'company',
            namespace: 'customer-info',
            type: 'single_line_text_field',
            value: companyName
          },
          {
            key: 'phone_number',
            namespace: 'customer-info',
            type: 'single_line_text_field',
            value: phone
          }
        ]
      }
    };

    if (website && urlRegex.test(website)) {
      createCustomerVariables.input.metafields.push({
        key: 'company_website',
        namespace: 'customer-info',
        type: 'url',
        value: website
      });
    }

    // Perform the customer creation query
    const createCustomerQueryResult = await axios.post(
      shopifyAPIUrl,
      {
        query: createCustomerQuery,
        variables: createCustomerVariables
      },
      {
        headers: {
          'X-Shopify-Access-Token': ADMIN_ACCESS_TOKEN,
          'Content-Security-Policy': 'frame-ancestors https://[REDACTED].myshopify.com https://admin.shopify.com',
        },
      }
    );

    const customer = createCustomerQueryResult.data.data.customerCreate.customer;
    const customerId = customer ? customer.id : null;
    
    if (customerId) {
      const generateAccountActivationUrlQuery = `mutation customerGenerateAccountActivationUrl($customerId: ID!) {
        customerGenerateAccountActivationUrl(customerId: $customerId) {
          accountActivationUrl
          userErrors {
            field
            message
          }
        }
      }`;
    
      const generateAccountActivationUrlVariables = {
        customerId: customerId,
      };
    
      const generateAccountActivationQueryResult = await axios.post(
        shopifyAPIUrl,
        {
          query: generateAccountActivationUrlQuery,
          variables: generateAccountActivationUrlVariables,
        },
        {
          headers: {
            'X-Shopify-Access-Token': adminAccessToken,
          },
        }
      );
    
      const activationUrl =
        generateAccountActivationQueryResult.data.data.customerGenerateAccountActivationUrl.accountActivationUrl;

      const accountActivationEmail = {
        message: {
          subject: 'Account Activation',
          body: {
            contentType: 'Text',
            content: `Click the following link to activate your account:\n${activationUrl}`
          },
          toRecipients: [
            {
              emailAddress: {
                address: email
              }
            }
          ],
          "from": {
            "emailAddress": {
              "address": "[REDACTED]"
            }
          }
        },
        saveToSentItems: 'true'
      };
        
      const newCustomerEmail = {
        message: {
          subject: `New customer: ${companyName}`,
          body: {
            contentType: 'html',
            content: `${companyName} has created an account!\n<a href="https://[REDACTED].myshopify.com/admin/customers/${customerId}">Click here to view the customer</a>.`
          },
          toRecipients: [
            {
              emailAddress: {
                address: "rjarmon@gmail.com"
              }
            }
          ],
          "from": {
            "emailAddress": {
              "address": "[REDACTED]"
            }
          }
        },
        saveToSentItems: 'true'
      };
    
      
      try {
        // Send the activation email
        await msGraphClient.api('/users/[REDACTED]/sendMail').post(accountActivationEmail);
      
        // If the activation email is sent successfully, send the new customer email
        await msGraphClient.api('/users/[REDACTED]/sendMail').post(newCustomerEmail);
      
        console.log('Emails sent successfully.');
      } catch (error) {
        console.error('Error sending activation email:', error);
      }

    } else {
    const userErrors = createCustomerQueryResult.data.data.customerCreate.userErrors;
    console.error('Error creating customer:', userErrors);
    }

    // Redirect the user to a success page
    res.redirect('https://[REDACTED].myshopify.com/account');
  } catch (error) {
    console.error('Error processing customer registration:', error);
    return res.status(500).send('Internal Server Error');
  }
});

/*------------------------
Server Listening
------------------------*/

app.listen(SERVER_PORT, () => {
  console.log(`Server is running on port ${SERVER_PORT}`);
});
