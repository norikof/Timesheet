﻿using System;
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
        //public string sName;

		public CTimesheet()
		{
			sFileTitle	= string.Empty;
			sMessage	= string.Empty;
			sError		= string.Empty;
            dPeriod = DateTime.Today;
            //sName = string.Empty;
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
            if (fiS.Name.StartsWith("~$")) return false; //Exit if timesheet file is temp file

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

            //bool bState = CheckContents(wshTimesheet);
            bool bState = CheckTimesheetFormat(wshTimesheet);
            if (bState == true)
            {
                bState = CreateTimeStarImportXlsx(wshImport, wshTimesheet);
            }

            wbkTimesheet.Close(false, misValue, misValue);  //Close timesheet excel file

            if (bState == true)
            {
                wbkImport.Worksheets[1].Activate();
                if (File.Exists(fiV.FullName))
                    wbkImport.Save();   //Overwrite saved import excel file if it already exists
                else
                    wbkImport.SaveAs(fiV.FullName); //Save as import excel file if not exists
            }
            wbkImport.Close(false, misValue, misValue);

            app.Quit();
            return bState;
		}

        private bool CreateTimeStarImportXlsx(Excel.Worksheet wshImport, Excel.Worksheet wshTimesheet)
        {
            //Create data to import to TimeStar

            int iImportRow = wshImport.UsedRange.Rows.Count + 1;
            int iTimesheetRow = 7;
            TimeSpan tEndTime = new TimeSpan(0, 0, 0, 0, 0);

            while (wshTimesheet.Cells[iTimesheetRow, 2].Value < dPeriod)    //Skip Anterior if the period is Posterior  
            {
                iTimesheetRow++;
            }

            while (wshTimesheet.Cells[iTimesheetRow, 2].Value != null)        //Repeat until the last line of Timesheet
            {
                int iPTO = 0;
                double dPTOTime = 0;

                if (wshTimesheet.Cells[iTimesheetRow, 11].Value != null) //Check Sick
                {
                    iPTO = 1;
                    dPTOTime = (double)wshTimesheet.Cells[iTimesheetRow, 11].Value;
                }
                if (wshTimesheet.Cells[iTimesheetRow, 12].Value != null)    //Check PTO
                {
                    iPTO = 2;
                    dPTOTime = (double)wshTimesheet.Cells[iTimesheetRow, 12].Value;
                }

                if ((wshTimesheet.Cells[iTimesheetRow, 4].Value == null) && (wshTimesheet.Cells[iTimesheetRow, 7].Value == null) && (dPTOTime == 0))    //Skip blank row
                {
                    iTimesheetRow++;
                    continue;
                }

                if (iPTO == 0)
                {
                    wshImport.Cells[iImportRow, 2] = wshTimesheet.Cells[3, 2].Value;  //Enter Worker Name
                    wshImport.Cells[iImportRow, 3] = wshTimesheet.Cells[iTimesheetRow, 2].Value;    //Enter Time Entry Date
                    wshImport.Cells[iImportRow, 4] = wshTimesheet.Cells[iTimesheetRow, 3].Value;    //Enter Punch In
                    wshImport.Cells[iImportRow, 5] = "IND";
                    iImportRow++;

                    wshImport.Cells[iImportRow, 2] = wshTimesheet.Cells[3, 2].Value;  //Enter Worker Name
                    wshImport.Cells[iImportRow, 3] = wshTimesheet.Cells[iTimesheetRow, 2].Value;    //Enter Time Entry Date
                    if (wshTimesheet.Cells[iTimesheetRow, 7].Value != null)
                        wshImport.Cells[iImportRow, 4] = wshTimesheet.Cells[iTimesheetRow, 7].Value;    //Enter Punch Out
                    else
                        wshImport.Cells[iImportRow, 4] = wshTimesheet.Cells[iTimesheetRow, 4].Value;    //Enter Lunch Out
                    wshImport.Cells[iImportRow, 5] = "OUT";
                    iImportRow++;
                }
                else
                {
                    wshImport.Cells[iImportRow, 2] = wshTimesheet.Cells[3, 2].Value;  //Enter Worker Name
                    wshImport.Cells[iImportRow, 3] = wshTimesheet.Cells[iTimesheetRow, 2].Value;  //Enter Time Entry Date
                    wshImport.Cells[iImportRow, 6] = dPTOTime;                      //Enter PTO hours
                    if (iPTO == 1) wshImport.Cells[iImportRow, 7] = "Sick";
                    if (iPTO == 2) wshImport.Cells[iImportRow, 7] = "PTO";
                    iImportRow++;
                }
                iTimesheetRow++;
            }
            return true;
        }

		private bool CheckTimesheetFormat(Excel.Worksheet wshTimesheet)
		{
            string sTitle = wshTimesheet.Cells[1, 1].Value;
            if (wshTimesheet.Cells[1, 1].Value != "TIME SHEET") return false;
            if (wshTimesheet.Cells[3, 1].Value != "Name:") return false;
            if (wshTimesheet.Cells[3, 2].Value == null) return false;
            if (wshTimesheet.Cells[3, 6].Value != "Client:") return false;
            if (wshTimesheet.Cells[3, 7].Value == null) return false;
            if (wshTimesheet.Cells[4, 1].Value != "Period:") return false;
            if (wshTimesheet.Cells[4, 2].Value == null) return false;
			return true;
		}
	}
}