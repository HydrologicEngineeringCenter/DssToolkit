using Hec.Dss;
using Hec.Excel;
using SpreadsheetGear;
using System;

namespace DssExcel
{
  internal class PairedDataReviewVM : ValidationVM
  {
    private IWorksheet worksheet1 = null;
    string dssFileName1 = "";
    public PairedDataReviewVM(IWorksheet worksheet, string dssFileName)
    {
      this.worksheet1 = worksheet;
      dssFileName1 = dssFileName;
    }
    public override bool Validate(out string errorMessage)
    {
      errorMessage = "";
      using (DssWriter writer = new DssWriter(dssFileName1))
      {
        try
        {
          worksheet1.WorkbookSet.GetLock();

          PairedData pd = ExcelPairedData.Read(worksheet1);
          // write to DSS

          writer.Write(pd);

        }
        catch (Exception e)
        {
          errorMessage = e.Message;
          return false;
        }
        finally
        {
          worksheet1.WorkbookSet.ReleaseLock();
        }
      }
      return true;
    }
  }
}