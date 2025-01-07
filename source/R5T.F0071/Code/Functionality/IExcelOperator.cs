using System;
using System.Threading.Tasks;

using R5T.F0069;
using R5T.T0132;


namespace R5T.F0071
{
	[FunctionalityMarker]
	public partial interface IExcelOperator : IFunctionalityMarker
	{
        public void In_ApplicationContext_Synchronous(
            Action<Application> applicationAction)
        {
            using var application = new Application();

			applicationAction(application);
        }

        public async Task In_ApplicationContext_Asynchronous(
			Func<Application, Task> applicationAction)
		{
            using var application = new Application();

			await applicationAction(application);
        }

		/// <summary>
		/// Chooses <see cref="In_ApplicationContext_Asynchronous(Func{Application, Task})"/> as the default.
		/// </summary>
		public Task In_ApplicationContext(
			Func<Application, Task> applicationAction)
			=> this.In_ApplicationContext_Asynchronous(
				applicationAction);

		public async Task In_WorkbookContext_New(
			Func<Workbook, Application, Task> workbookAction)
		{
			await this.In_ApplicationContext(
				async application =>
				{
					var workbook = application.NewWorkbook();

					await workbookAction(
						workbook,
						application);
				});
		}

        public async Task In_WorkbookContext_New(
			string excelWorkbookFilePath,
            Func<Workbook, Application, Task> workbookAction)
		{
			async Task Internal(Workbook workbook, Application application)
			{
				await workbookAction(workbook, application);

				workbook.SaveAs(excelWorkbookFilePath);
			}

			await this.In_WorkbookContext_New(
				Internal);
		}

        public async Task In_WorkbookContext_New(
            Func<Workbook, Task> workbookAction)
        {
            await this.In_ApplicationContext(
                async application =>
                {
                    var workbook = application.NewWorkbook();

                    await workbookAction(
                        workbook);
                });
        }

        public async Task In_WorkbookContext_New(
            string excelWorkbookFilePath,
            Func<Workbook, Task> workbookAction)
        {
            async Task Internal(Workbook workbook)
            {
                await workbookAction(workbook);

                workbook.SaveAs(excelWorkbookFilePath);
            }

            await this.In_WorkbookContext_New(
                Internal);
        }


        public void In_WorkbookContext_New(
            Action<Workbook, Application> workbookAction)
        {
            this.In_ApplicationContext_Synchronous(
                application =>
                {
                    var workbook = application.NewWorkbook();

                    workbookAction(
                        workbook,
                        application);
                });
        }

        public void In_WorkbookContext_New(
            string excelWorkbookFilePath,
            Action<Workbook, Application> workbookAction)
        {
            void Internal(Workbook workbook, Application application)
            {
                workbookAction(workbook, application);

                workbook.SaveAs(excelWorkbookFilePath);
            }

            this.In_WorkbookContext_New(
                Internal);
        }

        public void In_WorkbookContext_New(
            Action<Workbook> workbookAction)
        {
            this.In_ApplicationContext_Synchronous(
                application =>
                {
                    using var workbook = application.NewWorkbook();

                    workbookAction(
                        workbook);
                });
        }

        public void In_WorkbookContext_New(
            string excelWorkbookFilePath,
            Action<Workbook> workbookAction)
        {
            void Internal(Workbook workbook)
            {
                workbookAction(workbook);

                workbook.SaveAs(excelWorkbookFilePath);
            }

            this.In_WorkbookContext_New(
                Internal);
        }

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

			using var workbook = application.OpenWorkbook(workbookFilePath);

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

        public void Open_Workbook_ForInspection(
            string excelExecutableFilePath,
            string excelWorkbookFilePath)
        {
            var argumentsString_Enquoted = Instances.StringOperator.Ensure_Enquoted(
                excelWorkbookFilePath);

            Instances.ProcessOperator.Start(
                excelExecutableFilePath,
                argumentsString_Enquoted);
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