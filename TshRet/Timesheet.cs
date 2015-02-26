using System;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace TshRet
{
	public class CTimesheet
	{
		public string sFileTitle;
		public string sMessage;
		public string sError;
        public DateTime dPeriod;

		public CTimesheet()
		{
			sFileTitle	= string.Empty;
			sMessage	= string.Empty;
			sError		= string.Empty;
            dPeriod = DateTime.Today;
		}

		~CTimesheet()
		{
			sFileTitle	= null;
			sMessage	= null;
			sError		= null;
		}

        public bool CheckTimesheet(string sImportXlsx, string sTimesheetXls)
		{
            Excel.Workbook wbkTimesheet;    //timesheet excel file
            Excel.Worksheet wshTimesheet;   //its month's sheet of timesheet
            string sSheetName = null;            

            if (!System.IO.File.Exists(sTimesheetXls)) return false; //Exit if timesheet file is not exist

            CCheckDate checkdate = new CCheckDate();
            dPeriod = checkdate.CheckPeriod();   //Get period of this time
            if (dPeriod.Day == 1)
                sFileTitle = dPeriod.Year + "-" + dPeriod.ToString("MM") + "-Anterior";
            else
                sFileTitle = dPeriod.Year + "-" + dPeriod.ToString("MM") + "-Posterior";

			Excel.Application	app		= new Excel.Application();
            object misValue = System.Reflection.Missing.Value;

            //Open timesheet excel file
            FileInfo fiS = new FileInfo(sTimesheetXls); //Get timesheet file info

            try
            {
                wbkTimesheet = app.Workbooks.Open(fiS.FullName,
                                    0,
                                    Type.Missing, Type.Missing, "", "",//Enter empty password
                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                    Type.Missing, Type.Missing, Type.Missing, Type.Missing); //Open timesheet excel file

                sSheetName = dPeriod.ToString("MMM") + ". " + dPeriod.ToString("yyyy");  //Get sheet name of timesheet
                wshTimesheet = wbkTimesheet.Worksheets[sSheetName]; //Open its month's sheet
            }
            catch (COMException)
            {
                return false;
            }

            //Open template excel file
            FileInfo fiD = new FileInfo(sImportXlsx);
            string sImportFullPath = fiD.DirectoryName;
            string sSaveImportFullPath = sImportFullPath + "\\" + sFileTitle + ".xlsx";

            //Open saved import excel file
            FileInfo fiV = new FileInfo(sSaveImportFullPath);
            Excel.Workbook wbkImport;
            if (File.Exists(fiV.FullName))
            {
                wbkImport = app.Workbooks.Open(fiV.FullName);   //Open saved import excel file if it exists
            }
            else if (File.Exists(sImportXlsx))
            {
                wbkImport = app.Workbooks.Open(fiD.FullName);   //Open template excel file if saved import excel file deosn't exist
            }
            else
            {
                wbkImport = app.Workbooks.Add();        //Open new excel file if nothing exists
            }
            Excel.Worksheet wshImport = wbkImport.Worksheets[1];    //Open first sheet

			bool bState = CheckContents(wshTimesheet);
            if (bState == true)
            {
                bState = CreateTimeStarImportXlsx(wshImport, wshTimesheet);
            }

            wbkTimesheet.Close(false, misValue, misValue);  //Close timesheet excel file

            wbkImport.Worksheets[1].Activate();
            if (File.Exists(fiV.FullName))
                wbkImport.Save();   //Overwrite saved import excel file if it already exists
            else
                wbkImport.SaveAs(fiV.FullName); //Save as import excel file if not exists
            wbkImport.Close();

            app.Quit();
            return bState;
		}

        private bool CreateTimeStarImportXlsx(Excel.Worksheet wshImport, Excel.Worksheet wshTimesheet)
        {
            return true;
        }

		public bool CheckContents(Excel.Worksheet wshTimesheet)
		{
			if (!CheckTimesheetFormat(wshTimesheet)) {
				sError = "Invalid timesheet format.";
				return false;
			}
			int iTotalRow	= GetTotalHourRow(wshTimesheet);
			if (iTotalRow < 0) {
				sError = "Invalid timesheet format.";
				return false;
			}

			string sName	= wshTimesheet.Cells[2, 2].Value;
			string sPeriod	= GetPeriod(wshTimesheet);
			int iAttendance	= CountAttendance(iTotalRow, wshTimesheet);

			sMessage += wshTimesheet.Cells[2, 1].Value + " " + sName + "<br />";
			sMessage += wshTimesheet.Cells[2, 5].Value + " " + wshTimesheet.Cells[2, 6].Value + "<br />";
			sMessage += wshTimesheet.Cells[4, 1].Value + " " + sPeriod + "<br />";
			sMessage += "Total days: " + iAttendance.ToString() + "<br />";
			sMessage += "Total hours: " + wshTimesheet.Cells[iTotalRow, 6].Value + "<br />";

			sFileTitle = sName + " " + sPeriod;

			return true;
		}

		private bool CheckTimesheetFormat(Excel.Worksheet wshTimesheet)
		{
			string sTitle = wshTimesheet.Cells[1, 1].Value;
			if (wshTimesheet.Cells[1, 1].Value != "TIME SHEET")	return false;
			if (wshTimesheet.Cells[2, 1].Value != "Name:")		return false;
			if (wshTimesheet.Cells[2, 2].Value == null)			return false;
			if (wshTimesheet.Cells[2, 5].Value != "Client:")		return false;
			if (wshTimesheet.Cells[2, 6].Value == null)			return false;
			if (wshTimesheet.Cells[4, 1].Value != "Period:")		return false;
			if (wshTimesheet.Cells[4, 2].Value == null)			return false;
			return true;
		}

		private string GetPeriod(Excel.Worksheet wshTimesheet)
		{
			if (!(wshTimesheet.Cells[4, 2].Value is DateTime)) return string.Empty;
			DateTime dt = (DateTime)wshTimesheet.Cells[4, 2].Value;
			string sYear = dt.Year.ToString("0000");
			string sMonth = dt.Month.ToString("00");
			return sYear + "-" + sMonth;
		}

		private int GetTotalHourRow(Excel.Worksheet wshTimesheet)
		{
			int iRow;
			for (iRow = 7; iRow < 40; iRow++) {
				if (wshTimesheet.Cells[iRow, 4].Value is string)
					if (wshTimesheet.Cells[iRow, 4].Value == "Total:")
						return iRow;
			}
			return -1;
		}

		private int CountAttendance(int iTotalRow, Excel.Worksheet wshTimesheet)
		{
			int		iCount = 0;
			double	dWorking;
			double	dHours = 0;
			for (int i = 7; i < iTotalRow; i++) {
				if (wshTimesheet.Cells[i, 6].Value is double) {
					bool bNumeric = double.TryParse(wshTimesheet.Cells[i, 6].Text, out dWorking);
					if (!bNumeric)						continue;
					dHours += dWorking;
					iCount++;
				}
			}
			return iCount;
		}
	}
}