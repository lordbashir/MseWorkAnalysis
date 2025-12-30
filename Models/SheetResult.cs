namespace MseExcelAnalysis.Models
{
	using System.Collections.Generic;

	public record SheetResult(string SheetName, List<AttendanceRecord> Records);
}
