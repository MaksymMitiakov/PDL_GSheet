using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using Newtonsoft.Json.Linq;

namespace GSheets
{
	public partial class Form1 : Form
	{
		/* Global instance of the scopes required by this tool.
		 If modifying these scopes, delete your previously saved token.json/ folder. */
		static string[] Scopes = { SheetsService.Scope.Spreadsheets };
		static string ApplicationName = "Google Sheets API .NET";
		static string PDL_API_KEY = "60fbfebedcbccffe237f0509aef2f8b656e7e47ad4bcc05e3aca0afbcd1fa5a4";

		public Form1()
		{
			InitializeComponent();
		}

		private void Form1_Load(object sender, EventArgs e)
		{
		}

		private void btnLoad_Click(object sender, EventArgs e)
		{
			try
			{
				// Create Google Sheets API service.
				var service = new SheetsService(new BaseClientService.Initializer
				{
					HttpClientInitializer = GetCredential(),
					ApplicationName = ApplicationName
				});

				// Define request parameters.
				String spreadsheetId = txtSpreadSheetId.Text;
				String range = txtSheetName.Text + "!" + txtColumnStart.Text + txtFirstRow.Text + ":" + txtColumnEnd.Text;
				SpreadsheetsResource.ValuesResource.GetRequest request =
					service.Spreadsheets.Values.Get(spreadsheetId, range);

				ValueRange response = request.Execute();
				IList<IList<Object>> values = response.Values;
				dataGrid.RowCount = 0;
				if (values == null || values.Count == 0)
				{
					Console.WriteLine("No data found.");
					return;
				}
				dataGrid.RowCount = values.Count;
				//dataGrid.ColumnCount = values[0].Count;
				int i = 0;
				foreach (var row in values)
				{
					for (int j = 0; j < row.Count; j++)
					{
						dataGrid.Rows[i].Cells[j].Value = row[j];
					}
					i++;
				}
			}
			catch (FileNotFoundException ex)
			{
				MessageBox.Show(ex.Message);
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
		}

		private async void btnPopulate_Click(object sender, EventArgs e)
		{
			SheetsService sheetsService = new SheetsService(new BaseClientService.Initializer
			{
				HttpClientInitializer = GetCredential(),
				ApplicationName = ApplicationName,
			});

			// The ID of the spreadsheet to update.
			string spreadsheetId = txtSpreadSheetId.Text;

			//	Extend the columns
			SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum valueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;

			// TODO: Assign values to desired properties of `requestBody`. All existing
			// properties will be replaced:
			Google.Apis.Sheets.v4.Data.ValueRange requestBody = new Google.Apis.Sheets.v4.Data.ValueRange();
			requestBody.MajorDimension = "COLUMNS";
			var oblist = new List<object>() { "" };
			requestBody.Values = new List<IList<object>> { oblist };

			SpreadsheetsResource.ValuesResource.AppendRequest request = sheetsService.Spreadsheets.Values.Append(requestBody, spreadsheetId, txtSheetName.Text + "!DZ2");
			request.ValueInputOption = valueInputOption;
			request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.OVERWRITE;

			// To execute asynchronously in an async method, replace `request.Execute()` as shown:
			Google.Apis.Sheets.v4.Data.AppendValuesResponse response1 = request.Execute();

			//	Populate
			var client = new HttpClient();

			bool bHeaderWritten = false;
			for (int i = 0; i < dataGrid.RowCount; i++)
			{
				var row = dataGrid.Rows[i];
				row.HeaderCell.Value = "Reading";
				string sLinkedinUrl = row.Cells[4].Value == null ? "" : row.Cells[4].Value.ToString();
				string sEmail = row.Cells[5].Value == null ? "" : row.Cells[5].Value.ToString();
				string sCond = "", sTmpFileName = "";
				if(sLinkedinUrl == "" || sLinkedinUrl.IndexOf("/in/") < 0)
				{
					sCond = "email=" + sEmail;
					sTmpFileName = sEmail;
				}
				else
				{
					sCond = "profile=linkedin.com/in/" + sLinkedinUrl.Substring(sLinkedinUrl.IndexOf("/in/") + 4);
					sTmpFileName = sLinkedinUrl.Substring(sLinkedinUrl.IndexOf("/in/") + 4);
				}

				string sRes = "";
				if (File.Exists("TmpFiles//" + sTmpFileName))
				{
					sRes = File.ReadAllText("TmpFiles//" + sTmpFileName);
				}
				else
				{
					// Get the response.
					HttpResponseMessage response = await client.GetAsync("https://api.peopledatalabs.com/v5/person/enrich?api_key=" + PDL_API_KEY + "&pretty=False&required=mobile_phone&" + sCond);

					// Get the response content.
					HttpContent responseContent = response.Content;

					// Get the stream of the content.
					using (var reader = new StreamReader(await responseContent.ReadAsStreamAsync()))
					{
						// Write the output.
						sRes = await reader.ReadToEndAsync();

						File.WriteAllText("TmpFiles//" + sTmpFileName, sRes);
					}
				}

				JObject jRes = JObject.Parse(sRes);
				if (jRes.GetValue("status").ToString() == "200")
				{
					row.HeaderCell.Value = "Loaded";
					JObject jData = (JObject)jRes.GetValue("data");
					List<string> keys = jData.Properties().Select(p => p.Name).ToList();
					foreach(string key in keys)
					{
						int nColumnIndex = 0;
						if (dataGrid.Columns.Contains(key)) {
							for(int j = 0; j < dataGrid.Columns.Count; j++)
							{
								if (dataGrid.Columns[j].Name == key)
								{
									nColumnIndex = j;
									break;
								}
							}
						}
						else
						{
							nColumnIndex = dataGrid.Columns.Count;
							dataGrid.Columns.Add(key, key);
						}
						if(jData.GetValue(key).Type == JTokenType.String)
							row.Cells[nColumnIndex].Value = jData.GetValue(key).ToString();
						else
							row.Cells[nColumnIndex].Value = jData.GetValue(key).ToString(Newtonsoft.Json.Formatting.None);

						//UpdateSpreadSheetValue(txtSpreadSheetId.Text, txtSheetName.Text, "Z" + (i + 2), row.Cells[nColumnIndex].Value.ToString());
					}

					List<object> arrValues = new List<object>();
					if(!bHeaderWritten)
					{
						for (int j = 7; j < dataGrid.Columns.Count; j++)
						{
							arrValues.Add(dataGrid.Columns[j].Name.ToString());
						}
						UpdateSpreadSheetValues(txtSpreadSheetId.Text, txtSheetName.Text, "H1:DZ", arrValues);

						arrValues = new List<object>();
						bHeaderWritten = true;
					}

					for(int j = 7; j < dataGrid.Columns.Count; j++)
					{
						arrValues.Add(row.Cells[j].Value.ToString());
					}
					UpdateSpreadSheetValues(txtSpreadSheetId.Text, txtSheetName.Text, "H" + (i + 2).ToString() + ":DZ", arrValues);
				}
				else
					row.HeaderCell.Value = "Not Found";
			}

			
		}

		private UserCredential GetCredential()
		{
			UserCredential credential;
			// Load client secrets.
			using (var stream =
				   new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
			{
				/* The file token stores the user's access and refresh tokens, and is created
				 automatically when the authorization flow completes for the first time. */
				string credPath = "token";
				credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
					GoogleClientSecrets.FromStream(stream).Secrets,
					Scopes,
					"user",
					CancellationToken.None,
					new FileDataStore(credPath, true)).Result;
				Console.WriteLine("Credential file saved to: " + credPath);
			}
			return credential;
		}

		private void UpdateSpreadSheetValue(string sSpreadSheetId, string sSheet, string sCell, string sValue)
		{
			SheetsService sheetsService = new SheetsService(new BaseClientService.Initializer
			{
				HttpClientInitializer = GetCredential(),
				ApplicationName = ApplicationName,
			});

			// The ID of the spreadsheet to update.
			string spreadsheetId = sSpreadSheetId;

			// The A1 notation of the values to update.
			string range = sSheet + "!" + sCell;

			try
			{
				// How the input data should be interpreted.
				SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum valueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;

				// TODO: Assign values to desired properties of `requestBody`. All existing
				// properties will be replaced:
				Google.Apis.Sheets.v4.Data.ValueRange requestBody = new Google.Apis.Sheets.v4.Data.ValueRange();
				requestBody.MajorDimension = "COLUMNS";
				var oblist = new List<object>() { sValue };
				requestBody.Values = new List<IList<object>> { oblist };

				SpreadsheetsResource.ValuesResource.UpdateRequest request = sheetsService.Spreadsheets.Values.Update(requestBody, spreadsheetId, range);
				request.ValueInputOption = valueInputOption;

				// To execute asynchronously in an async method, replace `request.Execute()` as shown:
				Google.Apis.Sheets.v4.Data.UpdateValuesResponse response = request.Execute();
				// Data.UpdateValuesResponse response = await request.ExecuteAsync();

				// TODO: Change code below to process the `response` object:
				//Console.WriteLine(JsonConvert.SerializeObject(response));
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
		}

		private void UpdateSpreadSheetValues(string sSpreadSheetId, string sSheet, string sRange, List<object> arrValues)
		{
			SheetsService sheetsService = new SheetsService(new BaseClientService.Initializer
			{
				HttpClientInitializer = GetCredential(),
				ApplicationName = ApplicationName,
			});

			// The ID of the spreadsheet to update.
			string spreadsheetId = sSpreadSheetId;

			// The A1 notation of the values to update.
			string range = sSheet + "!" + sRange;

			try
			{
				// How the input data should be interpreted.
				string valueInputOption = "RAW";

				//	The new values to apply to the spreadsheet
				List<ValueRange> data = new List<ValueRange>();

				BatchUpdateValuesRequest requestBody = new BatchUpdateValuesRequest();
				requestBody.ValueInputOption = valueInputOption;
				requestBody.Data = data;

				ValueRange valueRange = new ValueRange();
				valueRange.Range = range;
				valueRange.Values = new List<IList<object>> { arrValues };
				data.Add(valueRange);

				SpreadsheetsResource.ValuesResource.BatchUpdateRequest request = sheetsService.Spreadsheets.Values.BatchUpdate(requestBody, spreadsheetId);

				// To execute asynchronously in an async method, replace `request.Execute()` as shown:
				BatchUpdateValuesResponse response = request.Execute();
				// Data.UpdateValuesResponse response = await request.ExecuteAsync();

				// TODO: Change code below to process the `response` object:
				//Console.WriteLine(JsonConvert.SerializeObject(response));
			}
			catch (Google.GoogleApiException gEx)
			{
				// How the input data should be interpreted.
				SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum valueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;

				// TODO: Assign values to desired properties of `requestBody`. All existing
				// properties will be replaced:
				Google.Apis.Sheets.v4.Data.ValueRange requestBody = new Google.Apis.Sheets.v4.Data.ValueRange();
				requestBody.MajorDimension = "COLUMNS";
				var oblist = new List<object>() { "" };
				requestBody.Values = new List<IList<object>> { oblist };

				SpreadsheetsResource.ValuesResource.AppendRequest request = sheetsService.Spreadsheets.Values.Append(requestBody, spreadsheetId, sSheet + "!DZ");
				request.ValueInputOption = valueInputOption;
				request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.OVERWRITE;

				// To execute asynchronously in an async method, replace `request.Execute()` as shown:
				Google.Apis.Sheets.v4.Data.AppendValuesResponse response1 = request.Execute();
				// Data.UpdateValuesResponse response = await request.ExecuteAsync();

				// TODO: Change code below to process the `response` object:
				//Console.WriteLine(JsonConvert.SerializeObject(response));s
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
		}
	}
}