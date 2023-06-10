require('dotenv').config();

const express = require('express');
const app = express();
const port = process.env.PORT || 3000;

const { GoogleSpreadsheet } = require('google-spreadsheet');
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

app.get('/', async (req, res) => {
	try {
		res.status(200).json({ message: 'Welcome, Space Cowboy!' });
	} catch (error) {
		console.error(error);
		res.status(500).json({ error: 'An error occurred.' });
	}
});

app.get('/make-sign', async (req, res) => {
	try {
		await accessSpreadsheet();
		res.status(200).json({ message: 'Quote added successfully!' });
	} catch (error) {
		console.error(error);
		res
			.status(500)
			.json({ error: 'An error occurred while adding the quote.' });
	}
});

app.listen(port, () => {
	console.log(`Server is running on port ${port}`);
});
