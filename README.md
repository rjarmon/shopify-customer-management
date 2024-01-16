# Shopify Customer Management App
This application utilizes node.js and express.js to make GraphQL requests with Shopify and Microsoft/Azure to manage customer data and send email notifications to appropriate parties.

## Customer Creation
When a new customer registration form is filled out, the app will create the customer, creating additional metafields for the customer for additional data via Shopify's GraphQL API.
The app will then continue to use Shopify's GraphQL API to generate a customer account activation URL and send it to the new customer with Microsft's GraphQL API.
The app will also send an email, again using Microsoft's GraphQL API, to the appropriate Shopify store's staff members to notify them that a new customer account has been created. This email will contain a link to that customer's account in the store's admin backend.

## File Upload
When a customer uploades a file via a form on their profile page, the application creates a staged upload via Shopify's GraphQL API based on the type of file that was uploaded.
The app continues to make API requests to Shopify to create the file in the stores admin backend from the staged upload.
Continuing to make API requests, the app gets the id for the file that was created and connects it to the customer's account by setting it as a `file_reference` metafield for that customer.
The app then makes a Microsoft GraphQL API request to send an email notification to the appropriate Shopify store's staff members to notify them that a customer has uploaded their tax exempt documention. This email will contain a link to the customer's account in the store's admin backend so they can review the document and choose whether or not to approve the account for full store access.
