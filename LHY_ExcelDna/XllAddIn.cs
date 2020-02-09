using Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;

namespace LHY_ExcelDna
{
    class XllAddIn : IExcelAddIn
    {

        public void AutoOpen()
        {
            //XlCall.Excel(XlCall.xlcAlert, "AutoOpen");
        }

        public void AutoClose()
        {
        }

        //public void WorkbookActivate(Workbook wb)
        //{
        //}

        //private void Workbook_SheetActivate(object Sh)
        //{
        //}
    }
}
