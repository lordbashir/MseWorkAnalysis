using ExcelDataReader;
using iText.IO.Font.Constants;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using MseExcelAnalysis.Models;
using Serilog;
using System;
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

			foreach (var analysisFile in analysisFiles) 
			{
				FileInfo inputExcelFileInfo = new(analysisFile.Trim());
				var fileNameWithoutExtension = inputExcelFileInfo.Name.Split('.').First();

				string resultPdfPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileNameWithoutExtension + "TimesheetAnalysis.pdf");

				Log.Information($"Starting file analysis - {inputExcelFileInfo.Name}");

				if (inputExcelFileInfo.DirectoryName != null)
				{
					resultPdfPath = System.IO.Path.Combine(inputExcelFileInfo.DirectoryName, fileNameWithoutExtension + "TimesheetAnalysis.pdf");
				}
				var result = int.TryParse(args[1].Trim(), out var analysisType);

				if (analysisType == (int)AnalysisType.Attendance)
				{
					Log.Information("Analysis type is - Attendance");
					var data = await ReadAttendanceExcelStreamingParallelAsync(inputExcelFileInfo.FullName);

					await WritePdfAsync(data, resultPdfPath, fileNameWithoutExtension);
				}
				Log.Information("PDF successfully created");
			}

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
	static async Task<List<SheetResult>> ReadAttendanceExcelStreamingParallelAsync(string path)
	{
		return await Task.Run(() =>
		{
			var results = new List<SheetResult>();
			var syncLock = new object();

			System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

			using var stream = File.Open(path, FileMode.Open, FileAccess.Read);
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

			//if (!headerMap.ContainsKey("Projekt") || !headerMap.ContainsKey("Stunden"))
			//{
			//	return;
			//}


			for (int r = headerRow + 1; r < sheet.Rows.Count; r++)
			{
				var row = sheet.Rows[r];
				if (row[0] == null || row[headerMap["Projekt"]]?.ToString() == null) 
				{ 
					continue; 
				}

				//var isDateFirstColumn = DateTime.TryParse(
				//			row[0].ToString(),
				//			CultureInfo.InvariantCulture,
				//			DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AssumeUniversal,
				//			out var dateTime
				//		);


				////var DateColumn = DateTime.ParseExact(row[0].ToString(), "dd.MM.yyyy", CultureInfo.InvariantCulture);
				//if( !isDateFirstColumn)
				//{
				//	continue;
				//}

				string projekt = row[headerMap["Projekt"]]?.ToString();
				string stundenStr = row[headerMap["Stunden"]]?.ToString();

				List<string> nonProductiveProjects = new List<string> { "urlaub", "krank", "feiertag" };

				if (!string.IsNullOrWhiteSpace(projekt) || (projekt != null && projekt.ToLower() == "x"))
				{
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
			}

			foreach (var item in projektStundenMap)
			{
				records.Add(new AttendanceRecord(item.Key, item.Value, item.Value/8));
			}

			lock (syncLock)
			{
				results.Add(new SheetResult(sheet.Name, records));
				Log.Information("Processed {Sheet} ({Count} rows)", sheet.Name, records.Count);
			}

			//// Second pass: process sheets in parallel
			//Parallel.ForEach(sheets, sheet =>
			//{
				
			//});

			return results;
		});
	}


	// =========================
	// PDF Writing + Chart (Async)
	// =========================
	static async Task WritePdfAsync(List<SheetResult> sheets,string outputPath, string fileNameWithoutExtension)
	{
		await Task.Run(() =>
		{
			using var writer = new PdfWriter(outputPath);
			using var pdf = new PdfDocument(writer);
			using var document = new Document(pdf, PageSize.A4);

			// Create fonts
			PdfFont boldFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
			PdfFont regularFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA);

			// Title
			//document.Add(new Paragraph("Project Working Hours Report")
			//	.SetFont(boldFont)
			//	.SetFontSize(18)
			//	.SetTextAlignment(TextAlignment.CENTER));
			var title = fileNameWithoutExtension.ToUpper() + " - Working Reports";

			foreach (var sheet in sheets)
			{
				//document.Add(new AreaBreak());

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

				// Charts
				//PdfPage chartPage = pdf.AddNewPage();
				//PdfCanvas canvas = new PdfCanvas(chartPage);

				//foreach (var r in sheet.Records)
				//{
				//	// Bar chart
				//	DrawBarChart(
				//		canvas,
				//		new Rectangle(50, 400, 200, 200),
				//		Convert.ToInt32( r.Stunden),
				//		r.Projekt);
				//}

				//// Bar chart
				//DrawBarChart(
				//	canvas,
				//	new Rectangle(50, 400, 200, 200),
				//	sheet.Records.Count,
				//	"Projects");

				//// Pie chart
				//DrawPieChart(
				//	canvas,
				//	400, 500,
				//	80,
				//	sheet.Records.Count,
				//	sheets.Sum(s => s.Records.Count));
			}
		});
	}


	//// =========================
	//// Simple Bar Chart Drawing
	//// =========================
	//static void DrawBarChart(PdfCanvas canvas, Rectangle area, int value, string label)
	//{
	//	float maxBarHeight = area.GetHeight() - 40;
	//	float barHeight = Math.Min(value * 10, maxBarHeight);

	//	// Axis
	//	canvas.MoveTo(area.GetLeft(), area.GetBottom())
	//		  .LineTo(area.GetLeft(), area.GetTop())
	//		  .LineTo(area.GetRight(), area.GetBottom())
	//		  .Stroke();

	//	// Bar
	//	canvas.SetFillColor(ColorConstants.BLUE);
	//	canvas.Rectangle(
	//		area.GetLeft() + 40,
	//		area.GetBottom(),
	//		60,
	//		barHeight);
	//	canvas.Fill();

	//	// Label
	//	canvas.BeginText()
	//		.MoveText(area.GetLeft() + 40, area.GetBottom() + barHeight + 10)
	//		.SetFontAndSize(iText.Kernel.Font.PdfFontFactory.CreateFont(), 10)
	//		.ShowText($"{label}: {value}")
	//		.EndText();
	//}

	//// =========================
	//// Simple Pie Chart Drawing
	//// =========================
	//static void DrawPieChart(PdfCanvas canvas, float centerX, float centerY, float radius, int value, int total)
	//{
	//	float sweepAngle = 360f * value / total;

	//	canvas.SetFillColor(ColorConstants.GREEN);
	//	canvas.MoveTo(centerX, centerY);
	//	canvas.Arc(
	//		centerX - radius,
	//		centerY - radius,
	//		centerX + radius,
	//		centerY + radius,
	//		0,
	//		sweepAngle);
	//	canvas.ClosePath();
	//	canvas.Fill();

	//	canvas.SetFillColor(ColorConstants.LIGHT_GRAY);
	//	canvas.MoveTo(centerX, centerY);
	//	canvas.Arc(
	//		centerX - radius,
	//		centerY - radius,
	//		centerX + radius,
	//		centerY + radius,
	//		sweepAngle,
	//		360 - sweepAngle);
	//	canvas.ClosePath();
	//	canvas.Fill();
	//}
}
