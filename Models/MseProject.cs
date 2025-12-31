namespace MseExcelAnalysis.Models
{
	public class MseProject
	{
		private readonly decimal _tagessatz1 = 1100.0m;
		private readonly decimal _tagessatz2 = 1122.0m;
		private readonly decimal _tagessatz3 = 600.0m;
		private readonly decimal _tagessatz4 = 2000.0m;

		private readonly List<string> _tagesSatz1Projects = new List<string>{"MBAG", "RAD-DB WARTUNG", "Rad-DB", "Rad-DB (AMG)", "QMSR" };
		private readonly List<string> _tagesSatz3Projects = new List<string> { "SOLARSCHMIEDE" };
		private readonly List<string> _tagesSatz4Projects = new List<string> { "Vesuv Wartung" };

		public required string ProjectName { get; set; }
		public decimal OrderAmount { get; set; }
		public decimal AvailableWorkingDays
		{
			get
			{
				if(_tagesSatz1Projects.Contains(ProjectName))
				{
					return OrderAmount / _tagessatz1;
				}
				else if (_tagesSatz3Projects.Contains(ProjectName))
				{
					return OrderAmount / _tagessatz3;
				}
				else if (_tagesSatz4Projects.Contains(ProjectName))
				{
					return OrderAmount / _tagessatz4;
				}
				else
				{
					return OrderAmount / _tagessatz2;
				}
			}
		}
		public decimal CompletedWorkingDays { get; set; }
	}
}
