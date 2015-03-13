using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace TshRet
{
	class CTsImportFile
    {
        public string sFileTitle;
        public DateTime dPeriod;
        public int[,] sSickInfo = null;

		public CTsImportFile()
        {
            sFileTitle = string.Empty;
            dPeriod = DateTime.Today;
            sSickInfo = new int[,] { };
		}

		~CTsImportFile()
        {
            sFileTitle = null;
		}

        public bool CreateImportFileFromFieldglassData(string sImportXlsx, string sFieldglassXls, string sTEListXls)
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
			if (bState == true) 
            {
                bState = CreateTimeStarImportXlsx(wshImport, wshFGTimesheet, wshFGAbsence, wshTEList);
			}

            wbkFieldglass.Close(false, misValue, misValue);  //Close Fieldglass excel file

            if (bState == true)
            {
                wbkImport.Worksheets[1].Activate();
                if (File.Exists(sSaveImportFullPath))
                    wbkImport.Save();   //Overwrite saved import excel file if it already exists
                else
                    wbkImport.SaveAs(sSaveImportFullPath); //Save as import excel file if not exists
            }
            wbkImport.Close(false, misValue, misValue);

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

		private bool CreateTimeStarImportXlsx(Excel.Worksheet wshImport, Excel.Worksheet wshFGTimesheet, Excel.Worksheet wshFGAbsence, Excel.Worksheet wshTEList)
		{
			// ここで import させる data を作成．
            int iImportRow = wshImport.UsedRange.Rows.Count + 1;
            int iFieldglassRow = 3;
            TimeSpan tEndTime = new TimeSpan(0, 0, 0, 0, 0);

            //bool bSick = CheckSickEmployeeID(wshTEList, wshFGAbsence);

            while (wshFGTimesheet.Cells[iFieldglassRow, 1].Value != null)        //Repeat until the last line of Field Glass
            {
                int iPTO = 0;
                double dTotalTime = (double)wshFGTimesheet.Cells[iFieldglassRow, 3].Value;   //Get Net Billable Hours of Field Glass
                double dNonbillableTime = (double)wshFGTimesheet.Cells[iFieldglassRow, 4].Value;   //Get Net Non-billable Hours of Field Glass
                double dTimeCul = 0;
                Excel.Range oRange;

                string sName = wshFGTimesheet.Cells[iFieldglassRow, 1].Value;
                int iId = CheckEmployeeID(wshTEList, sName);
                if (iId == 0)
                {
                    iFieldglassRow++;
                    continue;
                }
                
                if (wshFGTimesheet.Cells[iFieldglassRow, 2].Value == wshFGTimesheet.Cells[iFieldglassRow + 1, 2].Value)
                {
                    if (wshFGTimesheet.Cells[iFieldglassRow, 1].Value == wshFGTimesheet.Cells[iFieldglassRow + 1, 1].Value)
                    {
                        iFieldglassRow++;
                        continue;
                    }
                }

                if (dTotalTime != 0)
                {
                    dTimeCul = 8 + dTotalTime;                              //Culculate out time
                    tEndTime = TimeSpan.FromHours(dTimeCul);
                }
                else
                {
                    iFieldglassRow++;
                    continue;
                }

                if (iPTO == 0)
                {
                    //wshImport.Cells[iImportRow, 1] = wshFGTimesheet.Cells[iFieldglassRow, 5].Value;  //Enter Main Document ID
                    wshImport.Cells[iImportRow, 1] = iId;  //Enter Employee ID
                    wshImport.Cells[iImportRow, 2] = wshFGTimesheet.Cells[iFieldglassRow, 1].Value;  //Enter Worker Name
                    wshImport.Cells[iImportRow, 3] = wshFGTimesheet.Cells[iFieldglassRow, 2].Value;  //Enter Time Entry Date

                    wshImport.Cells[iImportRow, 4] = "8:00:00 AM";                              //Enter Punch In
                    wshImport.Cells[iImportRow, 5] = "IND";
                    if (wshFGTimesheet.Cells[iFieldglassRow, 7].Value != null)                   //Check Worker Comments
                    {
                        wshImport.Cells[iImportRow, 9] = wshFGTimesheet.Cells[iFieldglassRow, 7].Value;  //Enter Worker Comments
                        oRange = wshImport.Cells[iImportRow, 9];
                        oRange.Font.Bold = true;
                    }
                    if (wshFGTimesheet.Cells[iFieldglassRow, 8].Value != null)                   //Check Time Sheet Comments
                    {
                        if ((wshFGTimesheet.Cells[iFieldglassRow, 7].Value != null) && (wshFGTimesheet.Cells[iFieldglassRow, 8].Value != null))
                        {
                            if (wshImport.Cells[iImportRow, 9].Value != wshFGTimesheet.Cells[iFieldglassRow, 8].Value)
                            {
                                wshImport.Cells[iImportRow, 9] += "<br />";                     //Enter Time Sheet Comments if it's different from Worker Comments
                                wshImport.Cells[iImportRow, 9] += wshFGTimesheet.Cells[iFieldglassRow, 8].Value;
                                oRange = wshImport.Cells[iImportRow, 9];
                                oRange.Font.Bold = true;
                            }
                        }
                        else wshImport.Cells[iImportRow, 9] = wshFGTimesheet.Cells[iFieldglassRow, 8].Value;  //Enter Time Sheet Comments                        
                        oRange = wshImport.Cells[iImportRow, 9];
                        oRange.Font.Bold = true;
                    }                    
                    iImportRow++;

                    //wshImport.Cells[iImportRow, 1] = wshFGTimesheet.Cells[iFieldglassRow, 5].Value;  //Enter Main Document ID
                    wshImport.Cells[iImportRow, 1] = iId;  //Enter Employee ID
                    wshImport.Cells[iImportRow, 2] = wshFGTimesheet.Cells[iFieldglassRow, 1].Value;  //Enter Worker Name
                    wshImport.Cells[iImportRow, 3] = wshFGTimesheet.Cells[iFieldglassRow, 2].Value;  //Enter Time Entry Date
                    wshImport.Cells[iImportRow, 4] = tEndTime.ToString("hh':'mm");                  //Enter Punch Out
                    wshImport.Cells[iImportRow, 5] = "OUT";
                    iImportRow++;
                }
                else
                {
                    wshImport.Cells[iImportRow, 1] = wshFGTimesheet.Cells[iFieldglassRow, 5].Value;  //Enter Main Document ID
                    wshImport.Cells[iImportRow, 2] = wshFGTimesheet.Cells[iFieldglassRow, 1].Value;  //Enter Worker Name
                    wshImport.Cells[iImportRow, 3] = wshFGTimesheet.Cells[iFieldglassRow, 2].Value;  //Enter Time Entry Date
                    if (dNonbillableTime != 0) wshImport.Cells[iImportRow, 6] = dNonbillableTime;    //Enter Net Non-billable Hours if it's not zero
                    else wshImport.Cells[iImportRow, 6] = 8;                          //Enter 8 hours if Net Non-billable Hours is zero
                    wshImport.Cells[iImportRow, 7] = "Sick";
                    iImportRow++;
                }
                iFieldglassRow++;

            }
			return true;
		}

        private int CheckEmployeeID(Excel.Worksheet wshTEList, string sSearch)
        {
            //string[] aName = sName.Split(' ');
            //string sSearch = aName[1] + ", " + aName[0];

            Excel.Range currentFind = null;
            Excel.Range firstFind = null;
            int iId = 0;

            Excel.Range oTEList = wshTEList.get_Range("A1", "C30");
            // You should specify all these parameters every time you call this method,
            // since they can be overridden in the user interface. 
            currentFind = oTEList.Find(sSearch, Type.Missing,
                Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
                Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false,
                Type.Missing, Type.Missing);

            while (currentFind != null)
            {
                // Keep track of the first range you find. 
                if (firstFind == null)
                {
                    firstFind = currentFind;
                }

                // If you didn't move to a new range, you are done.
                else if (currentFind.get_Address(Excel.XlReferenceStyle.xlA1)
                      == firstFind.get_Address(Excel.XlReferenceStyle.xlA1))
                {
                    int iRow = currentFind.Row;
                    iId = (int)wshTEList.Cells[iRow, 1].Value;
                    break;
                }
                currentFind = oTEList.FindNext(currentFind);
            }
            return iId;
        }

        private bool CheckSickEmployeeID(Excel.Worksheet wshTEList, Excel.Worksheet wshFGAbsence)
        {
            int iAbsenceRow = 3;
            //int[] iArray = new int[20];
            int iArraynum = 0;

            if (wshFGAbsence.Cells[iAbsenceRow, 1].Value == null) return false;

            while (wshFGAbsence.Cells[iAbsenceRow, 1].Value != null)        //Repeat until the last line of Absence Report
            {
                int iResult = CheckEmployeeID(wshTEList, wshFGAbsence.Cells[iAbsenceRow, 3].Value);
                sSickInfo[iArraynum, 0] = iResult;
                sSickInfo[iArraynum, 1] = wshFGAbsence.Cells[iAbsenceRow, 6].Value;
                sSickInfo[iArraynum, 2] = wshFGAbsence.Cells[iAbsenceRow, 7].Value;
                sSickInfo[iArraynum, 3] = wshFGAbsence.Cells[iAbsenceRow, 10].Value;
                iArraynum++;
                iAbsenceRow++;
            }
            return true;
            //return iArray;
        }
	}
}
