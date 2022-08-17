const CLIENT_ID = '221706100089-1qtkob4oco195e5vn9dh3ihkmb25claq.apps.googleusercontent.com';
const API_KEY = 'AIzaSyBkPNOHg34iD8YRpxzXubJxEAqVO32dWfc';

// Discovery doc URL for APIs used by the quickstart
const DISCOVERY_DOC = 'https://sheets.googleapis.com/$discovery/rest?version=v4';

// Authorization scopes required by the API; multiple scopes can be
// included, separated by spaces.
const SCOPES = 'https://www.googleapis.com/auth/spreadsheets';

let tokenClient;
let gapiInited = false;
let gisInited = false;

$(document).ready(function () {
});

/**
 * Callback after api.js is loaded.
 */
function gapiLoaded() {
	gapi.load('client', intializeGapiClient);
}

function gisLoaded() {
	tokenClient = google.accounts.oauth2.initTokenClient({
		client_id: CLIENT_ID,
		scope: SCOPES,
		callback: '', // defined later
	});
	gisInited = true;
	maybeEnableButtons();
}

/**
 * Callback after the API client is loaded. Loads the
 * discovery doc to initialize the API.
 */
async function intializeGapiClient() {
	await gapi.client.init({
		apiKey: API_KEY,
		discoveryDocs: [DISCOVERY_DOC],
	});
	gapiInited = true;
	maybeEnableButtons();
}

function maybeEnableButtons() {
	if (gapiInited && gisInited) {
		document.getElementById('btnLoad').disabled = false;
		document.getElementById('btnPopulate').disabled = false;
	}
}

function onBtnLoadClick() {
	tokenClient.callback = async (resp) => {
		if (resp.error !== undefined) {
			throw (resp);
		}
		await loadContent();
	};

	if (gapi.client.getToken() === null) {
		// Prompt the user to select a Google Account and ask for consent to share their data
		// when establishing a new session.
		tokenClient.requestAccessToken({ prompt: 'consent' });
	} else {
		// Skip display of account chooser and consent dialog for an existing session.
		tokenClient.requestAccessToken({ prompt: '' });
	}
}

async function loadContent() {
	let response;
	try {
		// Fetch first 10 files
		response = await gapi.client.sheets.spreadsheets.values.get({
			spreadsheetId: $("#gsheet_id").val(),
			range: $("#sheet_name").val() + '!' + $("#column_range_start").val() + $("#first_row").val() + ':' + $("#column_range_end").val(),
		});
	} catch (err) {
		alert(err.message);
		console.log(err);
		return;
	}
	const range = response.result;
	if (!range || !range.values || range.values.length == 0) {
		const colspan = $("table thead th").length;
		$("table tbody").html('<tr><td colspan=' + colspan + '>No values found.</td></tr>');
		return;
	}
	
	let html = '';
	for(let i = 0; i < range.values.length; i++) {
		let row = range.values[i];
		html += '<tr><td>&nbsp;</td>';
		for(let j = 0; j < 7; j++) {	//	row.length
			if(row.length <= j || row[j] == '')
				html += '<td>&nbsp;</td>';
			else
				html += '<td>' + row[j] + '</td>';
		}
		html += '<td data-prop="mobile_phone"></td></tr>';
	}
	$("table tbody").html(html);
	//const output = range.values.reduce(
	//	(str, row) => `${str}${row[0]}, ${row[4]}\n`,
	//	'Name, Major:\n');
	//document.getElementById('content').innerText = output;
}

function onBtnPopulateClick() {
	if(!$("#valid_mobile").is(":checked") && !$("#valid_email").is(":checked")) {
		alert('Please choose at least one validation option.');
		return;
	}

	//	Extend columns
	var params = {
		// The ID of the spreadsheet to update.
		spreadsheetId: $("#gsheet_id").val(),

		// The A1 notation of a range to search for a logical table of data.
		// Values will be appended after the last row of the table.
		range: $("#sheet_name").val() + "!DZ2",

		// How the input data should be interpreted.
		valueInputOption: 'RAW',

		// How the input data should be inserted.
		insertDataOption: 'OVERWRITE',
	};

	var valueRangeBody = {
		"majorDimension": "COLUMNS",
    "values": [[""]]
	};

	var request = gapi.client.sheets.spreadsheets.values.append(params, valueRangeBody);
	request.then(async function(response) {
		// TODO: Change code below to process the `response` object:
		//console.log(response.result);

		let header = $("table thead tr");
		let bHeaderWritten = false;
		for(let i = 0; i < $("table tbody tr").length; i++) {
			let row = $("table tbody tr")[i];
			row.children[0].innerHTML = 'Reading';
			
			//	Get Parameters
			let linkedin_url = row.children[5].innerHTML;
			if(linkedin_url == '&nbsp;')	linkedin_url = '';
			let email = row.children[6].innerHTML;
			if(email == '&nbsp;')	email = '';
			let cond = '', tmpKey = '';
			if(linkedin_url == '' || linkedin_url.indexOf('/in/') < 0) {
				cond = 'email=' + email;
				tmpKey = email;
			}
			else {
				cond = 'profile=linkedin.com/in/' + linkedin_url.substring(linkedin_url.indexOf('/in/') + 4);
				tmpKey = linkedin_url.substring(linkedin_url.indexOf('/in/') + 4);
			}

			//	Check LocalStorage
			let res = '';
			res = localStorage.getItem(tmpKey);

			let process = function(key, index, json) {
				localStorage.setItem(key, JSON.stringify(json));

				if(!bHeaderWritten) {
					values = ['Mobile Phone'];
					for(var key in json) {
						if(key == 'mobile_phone')
							continue;
						$('table thead tr').append($('<th>').attr('data-prop', key).html(key));
						$('table tbody tr').append($('<td>').attr('data-prop', key));

						values.push(key);
					}

					updateSpreadSheetValues($("#gsheet_id").val(), $("#sheet_name").val(), "H1:DZ", values);

					bHeaderWritten = true;
				}

				let row = $("table tbody tr")[index];
				row.children[0].innerHTML = 'Loaded';
				values = [json['mobile_phone']];
				for(var key in json) {
					let content = '';
					if(json[key]) {
						if(typeof(json[key]) == 'object')
							content = JSON.stringify(json[key]);
						else
							content = json[key].toString();
					}
					$(row).find(`td[data-prop="${key}"]`).html('<div>' + content + '</div>');
					
					if(key == 'mobile_phone')
						continue;
					values.push(content);
				}

				//	Update Values on the Google Sheet
				updateSpreadSheetValues($("#gsheet_id").val(), $("#sheet_name").val(), "H" + (index + 2) + ":DZ", values);
			}

			if(res == null || res == '') {
				var required = '';
				if($("#valid_mobile").is(":checked") && $("#valid_email").is(":checked")) {
					required = 'required=mobile_phone AND emails';
				}
				else if($("#valid_mobile").is(":checked")) {
					required = 'required=mobile_phone';
				}
				else {
					required = 'required=emails';
				}
				//	Read from PDL API
				$.ajax({
					url: `https://api.peopledatalabs.com/v5/person/enrich?api_key=${$("#pdl_api_key").val()}&pretty=False&${required}&${cond}`
					, success: function(res) {
						if(res.status == '200') {
							process(tmpKey, i, res.data);
						}
						else {
							$("table tbody tr")[i].children[0].innerHTML = 'Not Found';
						}
					}
				});
			}
			else {
				process(tmpKey, i, JSON.parse(res));
			}
		}
	}, function(reason) {
		console.error('error: ' + reason.result.error.message);
	});
}

function updateSpreadSheetValues(sheet_id, sheet_name, range, values) {
	var params = {
		// The ID of the spreadsheet to update.
		spreadsheetId: sheet_id,
	};

	var batchUpdateValuesRequestBody = {
		// How the input data should be interpreted.
		valueInputOption: 'RAW',

		// The new values to apply to the spreadsheet.
		data: [
			{
				"majorDimension": "DIMENSION_UNSPECIFIED",
				"range": sheet_name + '!' + range,
				"values": [
					values
				]
			}
		],
	};

	var request = gapi.client.sheets.spreadsheets.values.batchUpdate(params, batchUpdateValuesRequestBody);
	request.then(function(response) {
		// TODO: Change code below to process the `response` object:
//		console.log(response.result);
	}, function(reason) {
		console.error('error: ' + reason.result.error.message);
	});
}