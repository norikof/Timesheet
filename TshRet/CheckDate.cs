using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace TshRet
{
    public class CCheckDate
    {
        public string sError;    
        
        public CCheckDate()
        {
            sError = string.Empty;
        }

        ~CCheckDate()
        {
            sError = null;
        }

        public DateTime CheckPeriod()
        {
            DateTime dPeriod;
            //DateTime dToday = DateTime.Today;
            DateTime dToday = new DateTime(2015, 2, 10);

            if (dToday.Day < 16){
                if (dToday.Month == 1)      dPeriod = new DateTime(dToday.Year - 1, 12, 16);
                else                        dPeriod = new DateTime(dToday.Year, dToday.Month-1, 16);
            }else
                dPeriod = new DateTime(dToday.Year, dToday.Month, 1);

            return dPeriod; 
        }
    }


}
