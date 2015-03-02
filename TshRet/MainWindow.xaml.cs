using System.Windows;
using System.IO;
using System;

namespace TshRet
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		public MainWindow()
		{
			InitializeComponent();
		}

		private void btnClose_Click(object sender, RoutedEventArgs e)
		{
			this.Close();
		}

		private void btnDownloadTimesheets_Click(object sender, RoutedEventArgs e)
		{

		}

		private void btnCreateUploadFile_Click(object sender, RoutedEventArgs e)
		{
			CreateTimStarImportFile();
		}

		private void CreateTimStarImportFile()
		{
			CTsImportFile tsimport = new CTsImportFile();
            CTimesheet timesheet = new CTimesheet();
            CCheckDate checkdate = new CCheckDate();

            DateTime dPeriod = checkdate.CheckPeriod();
            int sPeriodEndDay;
            if (dPeriod.Day == 1) sPeriodEndDay = 15;
            else sPeriodEndDay = DateTime.DaysInMonth(dPeriod.Year, dPeriod.Month);

			string sImportXlsx		= @"..\..\..\..\TimeImport.xlsx";
            string sFieldGlassXls = @"..\..\..\..\CAC_Time_Sheet_Combine_Report_02Posterior.xls";
            string sTimesheetFolder = @"\\sv11\Data\06_TimeSheet\TE Timesheet 2015\" + dPeriod.ToString("MM") + sPeriodEndDay;

            string[] saTimesheetXls = Directory.GetFiles(sTimesheetFolder, "*.xlsx");
            foreach (string sTimesheetXls in saTimesheetXls)
                timesheet.CheckTimesheet(sImportXlsx, sTimesheetXls);
            tsimport.CrateImportFileFromFieldglassData(sImportXlsx, sFieldGlassXls);
		}
	}
}
