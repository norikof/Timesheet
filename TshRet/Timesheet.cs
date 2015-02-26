using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace TshRet
{
	public class CTimesheet
	{
		public string sFileTitle;
		public string sMessage;
		public string sError;

		public CTimesheet()
		{
			sFileTitle	= string.Empty;
			sMessage	= string.Empty;
			sError		= string.Empty;
		}

		~CTimesheet()
		{
			sFileTitle	= null;
			sMessage	= null;
			sError		= null;
		}

		public bool CheckTimesheet(string sTempXlsx)
		{
			if (! System.IO.File.Exists(sTempXlsx)) return false;

			Excel.Application	app		= new Excel.Application();
			Excel.Workbook		wbk		= app.Workbooks.Open(sTempXlsx);
			Excel.Worksheet		wsh		= wbk.Worksheets[wbk.Worksheets.Count];

			bool bState = CheckContents(wsh);

			wbk.Close();
			app = null;
			return true;
		}

		public bool CheckContents(Excel.Worksheet wsh)
		{
			if (!CheckTimesheetFormat(wsh)) {
				sError = "Invalid timesheet format.";
				return false;
			}
			int iTotalRow	= GetTotalHourRow(wsh);
			if (iTotalRow < 0) {
				sError = "Invalid timesheet format.";
				return false;
			}

			string sName	= wsh.Cells[2, 2].Value;
			string sPeriod	= GetPeriod(wsh);
			int iAttendance	= CountAttendance(iTotalRow, wsh);

			sMessage += wsh.Cells[2, 1].Value + " " + sName + "<br />";
			sMessage += wsh.Cells[2, 5].Value + " " + wsh.Cells[2, 6].Value + "<br />";
			sMessage += wsh.Cells[4, 1].Value + " " + sPeriod + "<br />";
			sMessage += "Total days: " + iAttendance.ToString() + "<br />";
			sMessage += "Total hours: " + wsh.Cells[iTotalRow, 6].Value + "<br />";

			sFileTitle = sName + " " + sPeriod;

			return true;
		}

		private bool CheckTimesheetFormat(Excel.Worksheet wsh)
		{
			string sTitle = wsh.Cells[1, 1].Value;
			if (wsh.Cells[1, 1].Value != "TIME SHEET")	return false;
			if (wsh.Cells[2, 1].Value != "Name:")		return false;
			if (wsh.Cells[2, 2].Value == null)			return false;
			if (wsh.Cells[2, 5].Value != "Client:")		return false;
			if (wsh.Cells[2, 6].Value == null)			return false;
			if (wsh.Cells[4, 1].Value != "Period:")		return false;
			if (wsh.Cells[4, 2].Value == null)			return false;
			return true;
		}

		private string GetPeriod(Excel.Worksheet wsh)
		{
			if (!(wsh.Cells[4, 2].Value is DateTime)) return string.Empty;
			DateTime dt = (DateTime)wsh.Cells[4, 2].Value;
			string sYear = dt.Year.ToString("0000");
			string sMonth = dt.Month.ToString("00");
			return sYear + "-" + sMonth;
		}

		private int GetTotalHourRow(Excel.Worksheet wsh)
		{
			int iRow;
			for (iRow = 7; iRow < 40; iRow++) {
				if (wsh.Cells[iRow, 4].Value is string)
					if (wsh.Cells[iRow, 4].Value == "Total:")
						return iRow;
			}
			return -1;
		}

		private int CountAttendance(int iTotalRow, Excel.Worksheet wsh)
		{
			int		iCount = 0;
			double	dWorking;
			double	dHours = 0;
			for (int i = 7; i < iTotalRow; i++) {
				if (wsh.Cells[i, 6].Value is double) {
					bool bNumeric = double.TryParse(wsh.Cells[i, 6].Text, out dWorking);
					if (!bNumeric)						continue;
					dHours += dWorking;
					iCount++;
				}
			}
			return iCount;
		}
	}
}