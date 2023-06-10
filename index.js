require('dotenv').config();
const { GoogleSpreadsheet } = require('google-spreadsheet');
const schedule = require('node-schedule');
const loveQuotes = require('./quotes'); // load love quotes from another file

// Initialize the Google Spreadsheet by its ID (from the environment variables)
const doc = new GoogleSpreadsheet(process.env.GOOGLE_SPREADSHEET_ID);

// Function to fetch a random quote from the array
function getRandomQuote() {
	return loveQuotes[Math.floor(Math.random() * loveQuotes.length)];
}

async function accessSpreadsheet() {
	await doc.useServiceAccountAuth({
		client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
		private_key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
	});
	await doc.loadInfo();

	console.log(`Welcome, Space Cowboy!`);
	console.log(`Working on the table - ${doc.title}`);

	// Access the original sheet
	const originalSheet = doc.sheetsByTitle[process.env.GOOGLE_SPREADSHEET_TITLE];
	const originalSheetData = [];
	await originalSheet.loadCells();

	// Get all cell data from original sheet
	for (let i = 0; i < originalSheet.rowCount; i++) {
		const cellValue = originalSheet.getCell(i, 0).value;
		if (cellValue !== null) {
			originalSheetData.push(cellValue);
		}
	}

	// Save the original sheet name
	const originalSheetName = originalSheet.title;

	// Create a new sheet with a temporary unique name
	const newSheet = await doc.addSheet({ title: 'TempSheet' + Date.now() });
	await newSheet.loadCells();

	// Grab a random quote
	const randomQuote = getRandomQuote();

	// Add the new quote to the first row
	newSheet.getCell(0, 0).value = randomQuote;

	// Add the original sheet data starting from the second row
	for (let i = 0; i < originalSheetData.length; i++) {
		newSheet.getCell(i + 1, 0).value = originalSheetData[i];
	}

	// Save the changes
	await newSheet.saveUpdatedCells();

	// Delete the original sheet
	await originalSheet.delete();

	// Rename the new sheet back to the original name
	await newSheet.updateProperties({ title: originalSheetName });

	console.log(
		`A new quote has been added to the top of the table - ${doc.title}`
	);
}

// Schedule the job to run 3 times a day
schedule.scheduleJob('0 8,14,20 * * *', accessSpreadsheet);

// Run the function immediately on start
accessSpreadsheet().catch(console.error);
