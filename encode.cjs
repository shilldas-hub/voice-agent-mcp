const fs = require('fs');
// Read the file
const file = fs.readFileSync('service_account.json');
// Convert to Base64
console.log(file.toString('base64'));