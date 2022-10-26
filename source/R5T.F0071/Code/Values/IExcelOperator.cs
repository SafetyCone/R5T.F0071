using System;
using System.Threading.Tasks;

using R5T.F0069;
using R5T.T0131;


namespace R5T.F0071
{
	[ValuesMarker]
	public partial interface IExcelOperator : IValuesMarker
	{
		public void InModifyContext(
			string workbookFilePath,
			Action<Workbook> workbookModifyAction)
		{
			using var application = new Application();

			var workbook = application.OpenWorkbook(workbookFilePath);

			workbookModifyAction(workbook);

			workbook.Save();

			workbook.Close();
		}

		public async Task InModifyContext(
			string workbookFilePath,
			Func<Workbook, Task> workbookModifyAction)
		{
			using var application = new Application();

			var workbook = application.OpenWorkbook(workbookFilePath);

			await workbookModifyAction(workbook);

			workbook.Save();

			workbook.Close();
		}

		public TOutput InQueryContext<TOutput>(
			string workbookFilePath,
			Func<Workbook, TOutput> workbookQueryFunction)
        {
			using var application = new Application();

			var workbook = application.OpenWorkbook(workbookFilePath);

			var output = workbookQueryFunction(workbook);

			workbook.Close();

			return output;
        }

		public async Task<TOutput> InQueryContext<TOutput>(
			string workbookFilePath,
			Func<Workbook, Task<TOutput>> workbookQueryFunction)
		{
			using var application = new Application();

			var workbook = application.OpenWorkbook(workbookFilePath);

			var output = await workbookQueryFunction(workbook);

			workbook.Close();

			return output;
		}

		public Application OpenWorkbook(string workbookFilePath)
        {
			var application = new Application();

			// Ignore the returned workbook.
			_ = application.OpenWorkbook(workbookFilePath);

			return application;
        }
	}
}