function NewScrapCompletedQuests() {
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
	const sheetData = sheet.getDataRange().getValues();
	const getQuestNameRegex = new RegExp(/(Completion Status )(\[)(.*?)(\])/m);

	let QuestsNames = [];

	sheetData.forEach(async (item, position) => {
		if (position === 0) {
			QuestsNames.push({
				index: letterValue('h'),
				name: getQuestNameRegex.exec(item[letterValue('h')])[3]
			});
			QuestsNames.push({
				index: letterValue('i'),
				name: getQuestNameRegex.exec(item[letterValue('i')])[3]
			});
			QuestsNames.push({
				index: letterValue('j'),
				name: getQuestNameRegex.exec(item[letterValue('j')])[3]
			});
			QuestsNames.push({
				index: letterValue('k'),
				name: getQuestNameRegex.exec(item[letterValue('k')])[3]
			});
			QuestsNames.push({
				index: letterValue('l'),
				name: getQuestNameRegex.exec(item[letterValue('l')])[3]
			});
		} else {
			const url = item[letterValue('g')];

			const websiteContent = await UrlFetchApp.fetch(url).getContentText();

			QuestsNames.forEach(quest => {
				const isCompleted = websiteContent.includes(quest.name);
				sheet.getRange(position + 1, quest.index + 1).setValue(isCompleted);
			});
		}
	});
}

function SendEmails() {
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
	const sheetData = sheet.getDataRange().getValues();
	const getQuestNameRegex = new RegExp(/(Completion Status )(\[)(.*?)(\])/m);

	let QuestsNames = [];

	sheetData.forEach(async (item, position) => {
		if (position === 0) {
			QuestsNames.push({
				index: letterValue('h'),
				name: getQuestNameRegex.exec(item[letterValue('h')])[3]
			});
			QuestsNames.push({
				index: letterValue('i'),
				name: getQuestNameRegex.exec(item[letterValue('i')])[3]
			});
			QuestsNames.push({
				index: letterValue('j'),
				name: getQuestNameRegex.exec(item[letterValue('j')])[3]
			});
			QuestsNames.push({
				index: letterValue('k'),
				name: getQuestNameRegex.exec(item[letterValue('k')])[3]
			});
			QuestsNames.push({
				index: letterValue('l'),
				name: getQuestNameRegex.exec(item[letterValue('l')])[3]
			});
		} else {
			let message = `You have finished ${
				item[letterValue('c')]
			} Quests in Total!<br>Quests Breakdown: <br><br>`;

			QuestsNames.forEach(quest => {
				const questName = quest.name;
				const isCompleted = item[quest.index];
				message += `${questName}: ${isCompleted ? 'Complete' : 'Not Completed'}<br>`;
			});
			MailApp.sendEmail({
				to: 'lpanjwani@outlook.com',
				subject: 'Quests Tracking Updates | DSC @ MDX',
				htmlBody: message
			});
		}
	});
}

function letterValue(str) {
	var anum = {
		a: 1,
		b: 2,
		c: 3,
		d: 4,
		e: 5,
		f: 6,
		g: 7,
		h: 8,
		i: 9,
		j: 10,
		k: 11,
		l: 12,
		m: 13,
		n: 14,
		o: 15,
		p: 16,
		q: 17,
		r: 18,
		s: 19,
		t: 20,
		u: 21,
		v: 22,
		w: 23,
		x: 24,
		y: 25,
		z: 26
	};
	return anum[str] - 1;
}

function onOpen() {
	const ui = SpreadsheetApp.getUi();

	ui.createMenu('Quests Functions')
		.addItem('Retrieve Completed Quests', 'NewScrapCompletedQuests')
		.addToUi();
}
