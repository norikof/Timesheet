using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace TshRet
{
	class CTsImportFile
    {
        public string sFileTitle;
        public DateTime dPeriod;

		public CTsImportFile()
        {
            sFileTitle = string.Empty;
            dPeriod = DateTime.Today;
		}

		~CTsImportFile()
        {
            sFileTitle = null;
		}

        public bool CrateImportFileFromFieldglassData(string sImportXlsx, string sFieldglassXls, string sTEListXls)
		{
            if (!System.IO.File.Exists(sFieldglassXls)) return false; //Exit if Fieldglass file is not exist

            CCheckDate checkdate = new CCheckDate();
            dPeriod = checkdate.CheckPeriod();   //Get period of this time
            if (dPeriod.Day == 1)
                sFileTitle = dPeriod.Year + "-" + dPeriod.ToString("MM") + "-Anterior";
            else
                sFileTitle = dPeriod.Year + "-" + dPeriod.ToString("MM") + "-Posterior";

            Excel.Application app = new Excel.Application();
            object misValue = System.Reflection.Missing.Value;

			Excel.Workbook		wbkImport;
            
            //Open Fieldglass excel file
            FileInfo fiD = new FileInfo(sFieldglassXls); //Get Fieldglass file info
			string sImportFullPath = fiD.DirectoryName;

            //Open this month's import excel file
            string sSaveImportFullPath = sImportFullPath + "\\" + sFileTitle + ".xlsx";

            if (File.Exists(sSaveImportFullPath))
            {
                wbkImport = app.Workbooks.Open(sSaveImportFullPath);   //Open saved import excel file if it exists
            }
			else if (File.Exists(sImportXlsx)) {
                wbkImport = app.Workbooks.Open(sImportXlsx);   //Open template excel file if saved import excel file deosn't exist
			} else {
                wbkImport = app.Workbooks.Add();        //Open new excel file if nothing exists
			}
            Excel.Worksheet wshImport = wbkImport.Worksheets[1];    //Open first sheet

			FileInfo fiS = new FileInfo(sFieldglassXls);
			Excel.Workbook	wbkFieldglass	= app.Workbooks.Open(fiS.FullName);
            Excel.Worksheet wshFGTimesheet = wbkFieldglass.Worksheets[1];    //Open Timesheet sheet
            Excel.Worksheet wshFGAbsence = wbkFieldglass.Worksheets[2];    //Open Absence sheet

            //Open TE List excel file
            FileInfo fiT = new FileInfo(sTEListXls);
            Excel.Workbook wbkTEList = app.Workbooks.Open(fiT.FullName);
            Excel.Worksheet wshTEList = wbkTEList.Worksheets["TE"];    //Open first sheet

            bool bState = CheckFieldglassDataContents(wshFGTimesheet, wshFGAbsence);
			if (bState == true) {
				bState = CreateTimeStarImportXlsx(wshImport, wshFGTimesheet);
			}

			wbkFieldglass.Close();
			wbkImport.Worksheets[1].Activate();
			wbkImport.SaveAs(sImportFullPath);
			wbkImport.Close();
			app.Quit();
			return bState;
		}

		private int GetCurrentTimeSheet(Excel.Workbook wbk)
		{
			DateTime dtNow = DateTime.Now;

			foreach (Excel.Worksheet wsh in wbk.Worksheets) {
				DateTime dtTsh = new DateTime();
				bool bDate = DateTime.TryParse(wsh.Range["B4"].Text, out dtTsh);
				if (!bDate) continue;
				if (dtTsh.Year != dtNow.Year)	continue;
				if (dtTsh.Month != dtNow.Month)	continue;
				return wsh.Index;
			}
			return -1;
		}

        private bool CheckFieldglassDataContents(Excel.Worksheet wshFGTimesheet, Excel.Worksheet wshFGAbsence)
		{
			// ここで 内容を Fieldglass からきた file のcheck．
            if (!CheckFGTimesheetFormat(wshFGTimesheet))    //Check timesheet format
            {
                return false;
            }
            if (!CheckFGAbsenceReportFormat(wshFGAbsence))  //Check Absence Report format
            {
                return false;
            }
			return true;
		}

        private bool CheckFGTimesheetFormat(Excel.Worksheet wshFGTimesheet)
        {
            if (wshFGTimesheet.Cells[2, 1].Value != "Worker") return false;
            if (wshFGTimesheet.Cells[2, 2].Value != "Time Entry Date") return false;
            if (wshFGTimesheet.Cells[2, 3].Value != "Net Billable Hours") return false;
            if (wshFGTimesheet.Cells[2, 4].Value != "Net Non-billable Hours") return false;
            if (wshFGTimesheet.Cells[2, 5].Value != "Main Document ID") return false;
            if (wshFGTimesheet.Cells[2, 6].Value != "Time Sheet ID") return false;
            if (wshFGTimesheet.Cells[2, 7].Value != "Worker Comments") return false;
            if (wshFGTimesheet.Cells[2, 8].Value != "Time Sheet Comments (Separately)") return false;

            return true;
        }

        private bool CheckFGAbsenceReportFormat(Excel.Worksheet wshFGAbsence)
        {
            if (wshFGAbsence.Cells[2, 1].Value != "Worker ID") return false;
            if (wshFGAbsence.Cells[2, 2].Value != "Main Document ID") return false;
            if (wshFGAbsence.Cells[2, 3].Value != "Worker") return false;
            if (wshFGAbsence.Cells[2, 4].Value != "First Name") return false;
            if (wshFGAbsence.Cells[2, 5].Value != "Last Name") return false;
            if (wshFGAbsence.Cells[2, 6].Value != "Absence Start Date") return false;
            if (wshFGAbsence.Cells[2, 7].Value != "Absence End Date") return false;
            if (wshFGAbsence.Cells[2, 8].Value != "Absence Comment") return false;
            if (wshFGAbsence.Cells[2, 9].Value != "Absence Reason") return false;
            if (wshFGAbsence.Cells[2, 10].Value != "Partial Absence Hours") return false;
            if (wshFGAbsence.Cells[2, 11].Value != "Absence Submit Date") return false;

            return true;
        }

		private bool CreateTimeStarImportXlsx(Excel.Worksheet wshImport, Excel.Worksheet wshFieldglass)
		{
			// ここで import させる data を作成．
            int iImportRow = wshImport.UsedRange.Rows.Count + 1;
            int iFieldglassRow = 3;
            TimeSpan tEndTime = new TimeSpan(0, 0, 0, 0, 0);

            while (wshFieldglass.Cells[iFieldglassRow, 1].Value != null)        //Repeat until the last line of Field Glass
            {
                int iPTO = 0;
                double dTotalTime = (double)wshFieldglass.Cells[iFieldglassRow, 3].Value;   //Get Net Billable Hours of Field Glass
                double dNonbillableTime = (double)wshFieldglass.Cells[iFieldglassRow, 4].Value;   //Get Net Non-billable Hours of Field Glass
                double dTimeCul = 0;
                Excel.Range oRange;


            }
			return true;
		}
	}
}
