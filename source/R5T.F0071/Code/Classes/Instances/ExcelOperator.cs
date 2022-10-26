using System;


namespace R5T.F0071
{
	public class ExcelOperator : IExcelOperator
	{
		#region Infrastructure

	    public static IExcelOperator Instance { get; } = new ExcelOperator();

	    private ExcelOperator()
	    {
        }

	    #endregion
	}
}