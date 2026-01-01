using ExcelDataReader;
using iText.IO.Font.Constants;
using iText.Kernel.Font;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using MseExcelAnalysis.Models;
using Serilog;
using System.Globalization;

class Program
{
	static async Task Main(string[] args)
	{
		ConfigureLogging();

		try
		{
			Log.Information("Application started");
			if (args.Length < 2) 
			{
				Log.Information("Please provide folder path for analysis and Operation Type");
				return;
			}

			var mseProjects = new List<MseProject>
			{
				new MseProject { ProjectName = "MBAG", OrderAmount = 120587.5m, CompletedWorkingDays = 0.0m },
				new MseProject { ProjectName = "DTAG", OrderAmount = 274469.25m, CompletedWorkingDays = 0.0m },
				new MseProject { ProjectName = "FUSO", OrderAmount = 42000.0m, CompletedWorkingDays = 0.0m },
				new MseProject { ProjectName = "MSE", OrderAmount = 0.0m, CompletedWorkingDays = 0.0m },
				new MseProject { ProjectName = "QMSR", OrderAmount = 25300.0m, CompletedWorkingDays = 0.0m },
				new MseProject { ProjectName = "QMSR WARTUNG", OrderAmount = 14300.0m, CompletedWorkingDays = 0.0m },
				new MseProject { ProjectName = "VESUV", OrderAmount = 124542.0m, CompletedWorkingDays = 0.0m },
				new MseProject { ProjectName = "VESUV WARTUNG", OrderAmount = 24000.0m, CompletedWorkingDays = 0.0m },
				new MseProject { ProjectName = "RAD-DB WARTUNG", OrderAmount = 18700.0m, CompletedWorkingDays = 0.0m },
				new MseProject { ProjectName = "RAD-DB", OrderAmount = 75900.0m, CompletedWorkingDays = 0.0m },
				new MseProject { ProjectName = "RAD-DB (AMG)", OrderAmount = 75900.0m, CompletedWorkingDays = 0.0m },
				new MseProject { ProjectName = "REIFEN-DB", OrderAmount = 56100.0m, CompletedWorkingDays = 0.0m },
				new MseProject { ProjectName = "REIFEN-DB WARTUNG", OrderAmount = 15400.0m, CompletedWorkingDays = 0.0m },
				new MseProject { ProjectName = "REIFEN-DB AMG", OrderAmount = 20000.0m, CompletedWorkingDays = 0.0m },
				new MseProject { ProjectName = "AV-DB", OrderAmount = 19074.0m, CompletedWorkingDays = 0.0m },
				new MseProject { ProjectName = "AV-DB WARTUNG", OrderAmount = 19250.0m, CompletedWorkingDays = 0.0m },
				new MseProject { ProjectName = "LV-DB", OrderAmount = 30294.0m, CompletedWorkingDays = 0.0m },
				new MseProject { ProjectName = "LV-DB WARTUNG", OrderAmount = 19250.0m, CompletedWorkingDays = 0.0m },
				new MseProject { ProjectName = "SOLARSCHMEIDE", OrderAmount = 34800.0m, CompletedWorkingDays = 0.0m },
				new MseProject { ProjectName = "GESCHÄFTSFÜHRUNG", OrderAmount = 0.0m, CompletedWorkingDays = 0.0m },
				new MseProject { ProjectName = "PROJEKT-MANAGEMENT", OrderAmount = 0.0m, CompletedWorkingDays = 0.0m },
				new MseProject { ProjectName = "Fahrzeugverwaltung", OrderAmount = 0.0m, CompletedWorkingDays = 0.0m },
				new MseProject { ProjectName = "Tagesgeschäft", OrderAmount = 0.0m, CompletedWorkingDays = 0.0m },
				new MseProject { ProjectName = "Daily-Scrum-Meeting", OrderAmount = 0.0m, CompletedWorkingDays = 0.0m },
				new MseProject { ProjectName = "KRANK", OrderAmount = 0.0m, CompletedWorkingDays = 0.0m },
				new MseProject { ProjectName = "OTHERS", OrderAmount = 0.0m, CompletedWorkingDays = 0.0m }
			};

			var analysisFiles = Directory.GetFiles(args[0]).Where(af => af.EndsWith(".xlsm"));

			List<SheetResult> combineSheetResult = new List<SheetResult>();

			// Get analysis data for each timesheet excel file
			foreach (var analysisFile in analysisFiles) 
			{
				FileInfo inputExcelFileInfo = new(analysisFile.Trim());
				

				Log.Information($"Starting file analysis - {inputExcelFileInfo.Name}");

				var result = int.TryParse(args[1].Trim(), out var analysisType);

				if (analysisType == (int)AnalysisType.Attendance)
				{
					Log.Information("Analysis type is - Attendance");
					var data = await ReadAttendanceExcelStreamingParallelAsync(inputExcelFileInfo);

					// Combine the personal analysis result to prepare the report
					foreach (var sheetResult in data)
					{
						combineSheetResult.Add(sheetResult);
					}
				}
			}

			var allFlatRecords = combineSheetResult.SelectMany(res => res.Records.Where(r => !r.Projekt.ToLower().Equals("total")));

			foreach (var flatRecord in allFlatRecords)
			{
				var projectFromTimesheet = flatRecord.Projekt.Trim().ToLowerInvariant();
				var projectToSearch = "";

				if(projectFromTimesheet == "trucks" || projectFromTimesheet == "emt dtag" || projectFromTimesheet == "trcuks")
				{
					projectToSearch = "DTAG";
				}
				else if (projectFromTimesheet == "cars")
				{
					projectToSearch = "MBAG";
				}
				else if (projectFromTimesheet == "fuso")
				{
					projectToSearch = "FUSO";
				}
				else
				{
					projectToSearch = projectFromTimesheet;
				}

				var projectSearchSubString = projectToSearch.Trim().ToLowerInvariant().Substring(0, 3);

				var hasProjectEnding = projectFromTimesheet.Contains("amg")
					|| projectFromTimesheet.Contains("(amg)")
					|| projectFromTimesheet.Contains("@amg")
					|| projectFromTimesheet.Contains("@amd")
					|| projectFromTimesheet.Contains("wartung") 
					|| projectFromTimesheet.Contains("amd");

				if (hasProjectEnding)
				{
					var projectEndSearchSubString = projectToSearch.Trim().ToLowerInvariant().Substring(projectToSearch.Length-4, 4 );
					if(projectEndSearchSubString.EndsWith("@amg")
						|| projectEndSearchSubString.EndsWith("@amd")
						|| projectEndSearchSubString.EndsWith("amg)"))
					{
						projectEndSearchSubString = "amg";
					}

					var project = mseProjects.Where(mp => mp.ProjectName.Trim().ToLowerInvariant().StartsWith(projectSearchSubString)
					&& mp.ProjectName.Trim().ToLowerInvariant().EndsWith(projectEndSearchSubString))
					.FirstOrDefault();
					if (project != null) 
					{ 
						project.CompletedWorkingDays += flatRecord.Tagen;
					}
					else
					{
						var otherProject = mseProjects.Where(mp => mp.ProjectName.Trim().ToLowerInvariant().Equals("others")).FirstOrDefault();
						if (otherProject != null)
						{
							otherProject.CompletedWorkingDays += flatRecord.Tagen;
						}
					}
				}
				else
				{
					var project = mseProjects.Where(mp => mp.ProjectName.Trim().ToLowerInvariant().StartsWith(projectSearchSubString)
					&& (!mp.ProjectName.Trim().ToLowerInvariant().EndsWith("wartung"))).FirstOrDefault();
					if (project != null)
					{
						project.CompletedWorkingDays += flatRecord.Tagen;
					}
					else
					{
						var otherProject = mseProjects.Where(mp => mp.ProjectName.Trim().ToLowerInvariant().Equals("others")).FirstOrDefault();
						if (otherProject != null)
						{
							otherProject.CompletedWorkingDays += flatRecord.Tagen;
						}
					}
				}

			}

			var teamName = args[0].Split('\\').Last();

			string resultPdfPath = System.IO.Path.Combine(args[0], teamName + "TimesheetAnalysis.pdf");
			if (File.Exists(resultPdfPath)) 
			{
				File.Delete(resultPdfPath);
			}

			await WritePdfAsync(combineSheetResult, resultPdfPath, teamName);
			Log.Information("PDF successfully created");

			// Create project summary for a team
			var calculated = mseProjects.Where(mp => mp.CompletedWorkingDays > 0).ToList().OrderBy(mp => mp.ProjectName).ToList();
			
			string summaryPdfPath = System.IO.Path.Combine(args[0], teamName + "ProjectSummary.pdf");
			if (File.Exists(summaryPdfPath))
			{
				File.Delete(summaryPdfPath);
			}

			await WriteSummaryPdfAsync(calculated, summaryPdfPath, teamName);
			Log.Information("Summary PDF successfully created");

		}
		catch (Exception ex)
		{
			Log.Fatal(ex, "Unhandled application error");
		}
		finally
		{
			Log.CloseAndFlush();
		}
	}

	// =========================
	// Logging
	// =========================
	static void ConfigureLogging()
	{
		Log.Logger = new LoggerConfiguration()
			.MinimumLevel.Information()
			.WriteTo.Console()
			.WriteTo.File("logs/MseExcelAnalysis.log", rollingInterval: RollingInterval.Day)
			.CreateLogger();
	}

	// =========================
	// Excel Streaming (Async)
	// =========================
	static async Task<List<SheetResult>> ReadAttendanceExcelStreamingParallelAsync(FileInfo filePath)
	{
		return await Task.Run(() =>
		{
			var employeeName = filePath.Name.Split('.').First();

			var results = new List<SheetResult>();
			var syncLock = new object();

			System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

			using var stream = File.Open(filePath.FullName, FileMode.Open, FileAccess.Read);
			using var reader = ExcelReaderFactory.CreateReader(stream);

			var sheets = new List<(string Name, List<object[]> Rows)>();

			// First pass: read sheets sequentially (stream-safe)
			do
			{
				var rows = new List<object[]>();
				if (reader.Name == "2025")
				{
					while (reader.Read())
					{
						var row = new object[reader.FieldCount];
						reader.GetValues(row);
						rows.Add(row);
					}

					sheets.Add((reader.Name, rows));
				}
			}
			while (reader.NextResult());

			var headerMap = new Dictionary<string, int>();
			var projektStundenMap = new Dictionary<string, decimal>();
			var records = new List<AttendanceRecord>();

			var sheet = sheets[0];

			List<string> headerToTake = new List<string> { "projekt", "stunden" };
			var headerRow = 0;

			for (int r = 0; r < sheet.Rows.Count; r++)
			{
				var row = sheet.Rows[r];
				var firstColumn = row[0] == null? "" : row[0].ToString();

				if(string.IsNullOrWhiteSpace(firstColumn) || firstColumn.Trim().ToLower()!="datum")
				{
					continue;
				}

				headerRow = r;
				break;
			}

			for (int i = 0; i < sheet.Rows[headerRow].Length; i++)
			{
				var header = sheet.Rows[headerRow][i]?.ToString();
				if (!string.IsNullOrWhiteSpace(header) && headerToTake.Contains(header.ToLower()))
				{
					headerMap[header] = i;
				}
			}

			for (int r = headerRow + 1; r < sheet.Rows.Count; r++)
			{
				var row = sheet.Rows[r];
				if (row[0] == null) 
				{ 
					continue; 
				}

				string projekt = row[headerMap["Projekt"]]?.ToString();

				if (string.IsNullOrWhiteSpace(projekt) 
				|| projekt.Trim().ToLower() == "urlaub" 
				|| projekt.Trim().ToLower() == "x"
				|| projekt.Trim().ToLower() == "feiertag")
				{
					continue;
				}

				List<string> nonProductiveProjects = new List<string> { "krank" };
				string stundenStr = row[headerMap["Stunden"]]?.ToString();

				if (nonProductiveProjects.Contains(projekt.Trim().ToLowerInvariant()))
				{
					stundenStr = "8";
				}
				var result = decimal.TryParse(stundenStr, out var stunden);

				if (!projektStundenMap.ContainsKey(projekt.Trim()))
				{
					projektStundenMap[projekt.Trim()] = stunden;
				}
				else
				{
					projektStundenMap[projekt.Trim()] += stunden;
				}
			}

			foreach (var item in projektStundenMap)
			{
				records.Add(new AttendanceRecord(item.Key, item.Value, item.Value/8));
			}

			var sumStunden = projektStundenMap.Sum(it => it.Value);
			records.Add(new AttendanceRecord("Total", sumStunden, sumStunden / 8));

			lock (syncLock)
			{
				results.Add(new SheetResult(employeeName + "-" + sheet.Name, records));
				Log.Information("Processed {Sheet} ({Count} rows)", sheet.Name, records.Count);
			}

			return results;
		});
	}


	// =========================
	// PDF Writing + Chart (Async)
	// =========================
	static async Task WritePdfAsync(List<SheetResult> sheets,string outputPath, string teamName)
	{
		await Task.Run(() =>
		{
			using var writer = new PdfWriter(outputPath);
			using var pdf = new PdfDocument(writer);
			using var document = new Document(pdf, PageSize.A4);

			// Create fonts
			PdfFont boldFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
			PdfFont regularFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);

			//var title = fileNameWithoutExtension.ToUpper() + " - Working Reports";

			foreach (var sheet in sheets)
			{
				var title = teamName.ToUpper() + " - Working Reports";

				// Sheet header
				document.Add(new Paragraph(title + " ("+ sheet.SheetName+" )")
					.SetFont(boldFont)
					.SetFontSize(16));

				// Table
				Table table = new Table(3).UseAllAvailableWidth();
				table.AddHeaderCell(new Cell().Add(new Paragraph("Projekt").SetFont(boldFont)));
				table.AddHeaderCell(new Cell().Add(new Paragraph("Stunden").SetFont(boldFont)));
				table.AddHeaderCell(new Cell().Add(new Paragraph("Tagen").SetFont(boldFont)));
				
				var records = sheet.Records.Where(p => p.Projekt.Trim().ToLowerInvariant() != "total").OrderBy(r => r.Projekt).ToList();

				foreach (var r in records)
				{
					table.AddCell(new Cell().Add(new Paragraph(r.Projekt).SetFont(regularFont)));
					table.AddCell(new Cell().Add(new Paragraph(r.Stunden.ToString("F2", CultureInfo.InvariantCulture)).SetFont(regularFont)));
					table.AddCell(new Cell().Add(new Paragraph(r.Tagen.ToString("F2", CultureInfo.InvariantCulture)).SetFont(regularFont)));
				}

				// Write Total
				var totalRecord = sheet.Records.Where(p => p.Projekt.Trim().ToLowerInvariant() == "total").OrderBy(r => r.Projekt).ToList();
				foreach (var r in totalRecord)
				{
					table.AddCell(new Cell().Add(new Paragraph(r.Projekt).SetFont(boldFont)));
					table.AddCell(new Cell().Add(new Paragraph(r.Stunden.ToString("F2", CultureInfo.InvariantCulture)).SetFont(boldFont)));
					table.AddCell(new Cell().Add(new Paragraph(r.Tagen.ToString("F2", CultureInfo.InvariantCulture)).SetFont(boldFont)));
				}

				document.Add(table);
			}
		});
	}

	static async Task WriteSummaryPdfAsync(List<MseProject> projects, string outputPath, string teamName)
	{
		await Task.Run(() =>
		{
			using var writer = new PdfWriter(outputPath);
			using var pdf = new PdfDocument(writer);
			using var document = new Document(pdf, PageSize.A4);

			// Create fonts
			PdfFont boldFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
			PdfFont regularFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);

			//var title = fileNameWithoutExtension.ToUpper() + " - Working Reports";
			var title = teamName.ToUpper() + " - Project Summary Report";

			// Sheet header
			document.Add(new Paragraph(title)
				.SetFont(boldFont)
				.SetFontSize(16));

			// Table
			Table table = new Table(4).UseAllAvailableWidth();
			table.AddHeaderCell(new Cell().Add(new Paragraph("Projekt").SetFont(boldFont)));
			table.AddHeaderCell(new Cell().Add(new Paragraph("OrderAmount").SetFont(boldFont)));
			table.AddHeaderCell(new Cell().Add(new Paragraph("Available Work Days").SetFont(boldFont)));
			table.AddHeaderCell(new Cell().Add(new Paragraph("Completed Work Days").SetFont(boldFont)));

			foreach (var project in projects)
			{
				table.AddCell(new Cell().Add(new Paragraph(project.ProjectName).SetFont(regularFont)));
				table.AddCell(new Cell().Add(new Paragraph(project.OrderAmount.ToString("F2", CultureInfo.InvariantCulture)).SetFont(regularFont)));
				table.AddCell(new Cell().Add(new Paragraph(project.AvailableWorkingDays.ToString("F2", CultureInfo.InvariantCulture)).SetFont(regularFont)));
				table.AddCell(new Cell().Add(new Paragraph(project.CompletedWorkingDays.ToString("F2", CultureInfo.InvariantCulture)).SetFont(regularFont)));
			}

			document.Add(table);
			
		});
	}

}
