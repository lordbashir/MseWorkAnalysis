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

			var analysisFiles = Directory.GetFiles(args[0]);

			List<SheetResult> combineSheetResult = new List<SheetResult>();

			foreach (var analysisFile in analysisFiles) 
			{
				FileInfo inputExcelFileInfo = new(analysisFile.Trim());
				

				Log.Information($"Starting file analysis - {inputExcelFileInfo.Name}");

				var result = int.TryParse(args[1].Trim(), out var analysisType);

				if (analysisType == (int)AnalysisType.Attendance)
				{
					Log.Information("Analysis type is - Attendance");
					var data = await ReadAttendanceExcelStreamingParallelAsync(inputExcelFileInfo);
					foreach (var sheetResult in data)
					{
						combineSheetResult.Add(sheetResult);
					}
				}
				
			}

			var teamName = args[0].Split('\\').Last();

			string resultPdfPath = System.IO.Path.Combine(args[0], teamName + "TimesheetAnalysis.pdf");

			await WritePdfAsync(combineSheetResult, resultPdfPath, teamName);
			Log.Information("PDF successfully created");

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

				foreach (var r in sheet.Records)
				{
					table.AddCell(new Cell().Add(new Paragraph(r.Projekt).SetFont(regularFont)));
					table.AddCell(new Cell().Add(new Paragraph(r.Stunden.ToString("F2", CultureInfo.InvariantCulture)).SetFont(regularFont)));
					table.AddCell(new Cell().Add(new Paragraph(r.Tagen.ToString("F2", CultureInfo.InvariantCulture)).SetFont(regularFont)));
				}

				document.Add(table);
			}
		});
	}
	
}
