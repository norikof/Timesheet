using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace TshRet
{
	class CTsImportFile
	{
		public CTsImportFile()
		{
		}

		~CTsImportFile()
		{
		}

		public bool CrateImportFileFromFieldglassData(string sImportXlsx, string sFieldglassXls)
		{
			if (!System.IO.File.Exists(sFieldglassXls)) return false;

			Excel.Application	app		= new Excel.Application();

			Excel.Workbook		wbkImport;

			FileInfo fiD = new FileInfo(sFieldglassXls);
			string sImportFullPath = fiD.DirectoryName;
			if (File.Exists(sImportXlsx)) {
				wbkImport	= app.Workbooks.Open(fiD.FullName);
			} else {
				wbkImport	= app.Workbooks.Add();
			}

			Excel.Worksheet	wshImport		= wbkImport.Worksheets[1];
			FileInfo fiS = new FileInfo(sFieldglassXls);
			Excel.Workbook	wbkFieldglass	= app.Workbooks.Open(fiS.FullName);
			Excel.Worksheet	wshFieldglass	= wbkFieldglass.Worksheets[1];

			bool bState = CheckFieldglassDataContents(wshFieldglass);
			if (bState == true) {
				bState = CreateTimeStarImportXlsx(wshImport, wshFieldglass);
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

		private bool CheckFieldglassDataContents(Excel.Worksheet wshFieldglass)
		{
			// ここで 内容を Fieldglass からきた file のcheck．
			return true;
		}

		private bool CreateTimeStarImportXlsx(Excel.Worksheet wshImport, Excel.Worksheet wshFieldglass)
		{
			// ここで import させる data を作成．
			return true;
		}
	}
}
