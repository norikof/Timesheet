using System.Windows;
using System.IO;

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

			string sImportXlsx		= @"..\..\..\..\TimeImport.xlsx";
			string sFieldGlassXls	= @"..\..\..\..\CAC_Time_Sheet_Report.xls";
            string sTimesheetFolder = @"\\sv11\Data\06_TimeSheet\TE Timesheet 2015\" + dPeriod.ToString("MM") + sPeriodEndDay;

			tsimport.CrateImportFileFromFieldglassData(sImportXlsx, sFieldGlassXls);
		}
	}
}
