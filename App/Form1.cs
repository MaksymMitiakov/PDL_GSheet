using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using Newtonsoft.Json.Linq;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;

namespace GSheets
{
	public partial class Form1 : Form
	{
		/* Global instance of the scopes required by this tool.
		 If modifying these scopes, delete your previously saved token.json/ folder. */
		static string[] Scopes = { SheetsService.Scope.Spreadsheets };
		static string ApplicationName = "Google Sheets API .NET";
		static string PDL_API_KEY = "7235ac0c222dbdd7db4dd24fc7b152c395d8a55ede1c662862cae2ee99776b88";

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
					dataGrid.Rows[i].Cells[0].Value = i + 1;
					for (int j = 0; j < row.Count; j++)
					{
						dataGrid.Rows[i].Cells[j + 1].Value = row[j];
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
			try
			{
				PDL_API_KEY = txtPDLKey.Text;

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

				SpreadsheetsResource.ValuesResource.AppendRequest request = sheetsService.Spreadsheets.Values.Append(requestBody, spreadsheetId, txtSheetName.Text + "!CF2");
				request.ValueInputOption = valueInputOption;
				request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.OVERWRITE;

				// To execute asynchronously in an async method, replace `request.Execute()` as shown:
				Google.Apis.Sheets.v4.Data.AppendValuesResponse response1 = request.Execute();

				//	Populate
				JArray requests = new JArray();

				for (int i = 0; i < dataGrid.RowCount; i++)
				{
					var row = dataGrid.Rows[i];
					row.HeaderCell.Value = "Reading";
					string sLinkedinUrl = row.Cells[5].Value == null ? "" : row.Cells[5].Value.ToString();
					string sEmail = row.Cells[6].Value == null ? "" : row.Cells[6].Value.ToString();
					string sTmpFileName = "";
					JObject one = new JObject();
					if (sLinkedinUrl == "" || sLinkedinUrl.IndexOf("/in/") < 0)
					{
						JObject tmp = new JObject();
						tmp.Add("email", sEmail);
						one.Add("params", tmp);
						sTmpFileName = sEmail;
					}
					else
					{
						JObject tmp = new JObject();
						tmp.Add("profile", "linkedin.com/in/" + sLinkedinUrl.Substring(sLinkedinUrl.IndexOf("/in/") + 4));
						one.Add("params", tmp);
						sTmpFileName = sLinkedinUrl.Substring(sLinkedinUrl.IndexOf("/in/") + 4);
					}

					string sRes = "";
					if (File.Exists("TmpFiles//" + sTmpFileName))
					{
						sRes = File.ReadAllText("TmpFiles//" + sTmpFileName);
						JObject jRes = JObject.Parse(sRes);
						ProcessOneRow(row, jRes);
					}
					else
					{
						requests.Add(one);
						if (requests.Count >= 100)
						{
							await ProcessRequests(requests);

							requests.Clear();
						}
					}
					/*
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

						//for(int j = 7; j < dataGrid.Columns.Count; j++)
						//{
						//	arrValues.Add(row.Cells[j].Value.ToString());
						//}
						//UpdateSpreadSheetValues(txtSpreadSheetId.Text, txtSheetName.Text, "H" + (i + 2).ToString() + ":DZ", arrValues);
					}
					else
						row.HeaderCell.Value = "Not Found";*/
				}

				if (requests.Count > 0)
					await ProcessRequests(requests);

				BulkUpdateGSheet();

				MessageBox.Show("Operation completed");
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
		}

		private void BulkUpdateGSheet()
		{
			int startIndex = 0;
			int cntPerRequest = 2000;
			while (true)
			{
				List<List<object>> arrValues = new List<List<object>>();
				List<object> keys = new List<object>();
				List<object> values = new List<object>();
				int i, j;
				for (i = 8; i < dataGrid.Columns.Count; i++)
				{
					keys.Add(dataGrid.Columns[i].Name);
					values.Add(dataGrid.Columns[i].Name);
				}
				if (startIndex == 0)
				{
					arrValues.Add(values);
				}

				for (i = startIndex; i < dataGrid.Rows.Count && i < startIndex + cntPerRequest; i++) {
					var row = dataGrid.Rows[i];
					values = new List<object>();
					for (j = 8; j < dataGrid.Columns.Count; j++)
					{
						if (row.Cells[j].Value == null)
						{
							values.Add("");
						}
						else if (row.Cells[j].Value?.ToString().Length < 50000) {
							values.Add(row.Cells[j].Value?.ToString());
						}
						else
						{
							values.Add(row.Cells[j].Value?.ToString().Substring(0, 10000));
						}
					}
					arrValues.Add(values);
				}

				//	Update Values on the Google Sheet
				if (startIndex == 0)
					UpdateSpreadSheetValues(txtSpreadSheetId.Text, txtSheetName.Text, "H1:CF" + (i + 1), arrValues);
				else
					UpdateSpreadSheetValues(txtSpreadSheetId.Text, txtSheetName.Text, "H" + (startIndex + 2) + ":CF" + (i + 1), arrValues);

				startIndex += cntPerRequest;
				if (startIndex > dataGrid.Rows.Count)
					break;
			}
		}

		private async Task ProcessRequests(JArray requests)
		{
			string sRequired = "";
			if (chkValidEmail.Checked && chkValidMobilePhone.Checked)
			{
				sRequired = "email AND mobile_phone";
			}
			else if (chkValidEmail.Checked)
			{
				sRequired = "email";
			}
			else if (chkValidMobilePhone.Checked)
			{
				sRequired = "mobile_phone";
			}
			else
			{
				MessageBox.Show("Please choose at least one validation option.");
				return;
			}

			JObject param = new JObject();
			param.Add("required", sRequired);
			param.Remove("requests");
			param.Add("requests", requests);

			var client = new HttpClient();
			StringContent httpContent = new StringContent(param.ToString(), System.Text.Encoding.UTF8, "application/json");
			client.DefaultRequestHeaders.Add("x-api-key", PDL_API_KEY);

			// Get the response.
			HttpResponseMessage response = await client.PostAsync("https://api.peopledatalabs.com/v5/person/bulk?api_key=" + PDL_API_KEY, httpContent);

			// Get the response content.
			HttpContent responseContent = response.Content;

			// Get the stream of the content.
			using (var reader = new StreamReader(await responseContent.ReadAsStreamAsync()))
			{
				// Write the output.
				string sRes = await reader.ReadToEndAsync();

				JArray jRes = new JArray();
				try
				{
					jRes = JArray.Parse(sRes);
				}
				catch (Exception ex)
				{
					MessageBox.Show(sRes);
					return;
				}
				for (int j = 0; j < jRes.Count; j++)
				{
					JObject jReq = (JObject)requests[j];
					JObject jParam = (JObject)jReq.GetValue("params");
					DataGridViewRow tmpRow = null;
					if (jParam.ContainsKey("email"))
					{
						string sEmail = jParam.GetValue("email").ToString();
						for (int k = 0; k < dataGrid.Rows.Count; k++)
						{
							if (dataGrid.Rows[k].Cells[6].Value?.ToString() == sEmail)
							{
								tmpRow = dataGrid.Rows[k];
								if (((JObject)jRes[j]).GetValue("status").ToString() == "200")
								{
									try
									{
										File.WriteAllText("TmpFiles//" + jParam.GetValue("email").ToString(), ((JObject)jRes[j]).ToString());
									}
									catch (Exception) { }
								}
								break;
							}
						}
					}
					else
					{
						string sLinkedinUrl = jParam.GetValue("profile").ToString();
						for (int k = 0; k < dataGrid.Rows.Count; k++)
						{
							if (dataGrid.Rows[k].Cells[5].Value?.ToString().IndexOf(sLinkedinUrl) >= 0)
							{
								tmpRow = dataGrid.Rows[k];
								if (((JObject)jRes[j]).GetValue("status").ToString() == "200")
								{
									try
									{
										File.WriteAllText("TmpFiles//" + sLinkedinUrl.Substring(sLinkedinUrl.IndexOf("/in/") + 4), ((JObject)jRes[j]).ToString());
									}
									catch (Exception) { }
								}
								break;
							}
						}
					}
					if (tmpRow != null)
						ProcessOneRow(tmpRow, (JObject)jRes[j]);
				}
			}
		}

		private void ProcessOneRow(DataGridViewRow row, JObject jObj)
		{
			if (jObj.GetValue("status").ToString() == "200")
			{
				row.HeaderCell.Value = "Loaded";
				JObject jData = (JObject)jObj.GetValue("data");
				List<string> keys = jData.Properties().Select(p => p.Name).ToList();
				foreach (string key in keys)
				{
					int nColumnIndex = 0;
					if (dataGrid.Columns.Contains(key))
					{
						for (int j = 0; j < dataGrid.Columns.Count; j++)
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
					if (jData.GetValue(key).Type == JTokenType.String)
						row.Cells[nColumnIndex].Value = jData.GetValue(key).ToString();
					else
						row.Cells[nColumnIndex].Value = jData.GetValue(key).ToString(Newtonsoft.Json.Formatting.None);

					//UpdateSpreadSheetValue(txtSpreadSheetId.Text, txtSheetName.Text, "Z" + (i + 2), row.Cells[nColumnIndex].Value.ToString());
				}
			}
			else
			{
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

		private void UpdateSpreadSheetValues(string sSpreadSheetId, string sSheet, string sRange, List<List<object>> arrValues)
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
				IList<IList<object>> values = new List<IList<object>>();
				foreach(List<object> list in arrValues)
				{
					values.Add(list);
				}
				valueRange.Values = (IList<IList<object>>)values;
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
				MessageBox.Show(gEx.Message);
				return;
				// How the input data should be interpreted.
				SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum valueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;

				// TODO: Assign values to desired properties of `requestBody`. All existing
				// properties will be replaced:
				Google.Apis.Sheets.v4.Data.ValueRange requestBody = new Google.Apis.Sheets.v4.Data.ValueRange();
				requestBody.MajorDimension = "COLUMNS";
				var oblist = new List<object>() { "" };
				requestBody.Values = new List<IList<object>> { oblist };

				SpreadsheetsResource.ValuesResource.AppendRequest request = sheetsService.Spreadsheets.Values.Append(requestBody, spreadsheetId, sSheet + "!CF");
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