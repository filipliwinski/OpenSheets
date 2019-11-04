using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.IO;
using System.Linq;

namespace OpenSheets
{
    public class Spreadsheet : IDisposable
    {
        private Stream spreadsheetStream;
        private SpreadsheetDocument spreadsheet;


        public Spreadsheet()
        {
            spreadsheetStream = new MemoryStream();

            // Create a spreadsheet document
            spreadsheet = SpreadsheetDocument.Create(spreadsheetStream, SpreadsheetDocumentType.Workbook);

            // Add a WorkbookPart to the document.
            var workbookpart = spreadsheet.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart.
            var worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            workbookpart.Workbook.Save();
        }

        public string AddSheet(string name)
        {
            // Add a WorksheetPart to the WorkbookPart.
            var worksheetPart = spreadsheet.WorkbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            worksheetPart.Worksheet.Save();

            // Add Sheets to the Workbook.
            var sheets = spreadsheet.WorkbookPart.Workbook.AppendChild(new Sheets());

            uint sheetId;

            if (spreadsheet.WorkbookPart.Workbook.Sheets == null)
            {
                spreadsheet.WorkbookPart.Workbook.AppendChild(new Sheets());
                sheetId = 1;
            }
            else
            {
                sheetId = Convert.ToUInt32(spreadsheet.WorkbookPart.Workbook.Sheets.Count() + 1);
            }

            // Append a new worksheet and associate it with the workbook.
            var sheet = new Sheet()
            {
                Id = spreadsheet.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = sheetId,
                Name = name
            };
            sheets.Append(sheet);

            spreadsheet.WorkbookPart.Workbook.Save();

            return spreadsheet.WorkbookPart.GetIdOfPart(worksheetPart);
        }

        public void Close()
        {
            // Close the document.
            spreadsheet.Close();
        }

        #region IDisposable Support
        private bool disposedValue = false; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects).
                    if (spreadsheet != null) spreadsheet.Dispose();
                    if (spreadsheetStream != null) spreadsheetStream.Dispose();
                }

                // TODO: free unmanaged resources (unmanaged objects) and override a finalizer below.
                // TODO: set large fields to null.

                disposedValue = true;
            }
        }

        // TODO: override a finalizer only if Dispose(bool disposing) above has code to free unmanaged resources.
        // ~OpenSheets()
        // {
        //   // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
        //   Dispose(false);
        // }

        // This code added to correctly implement the disposable pattern.
        public void Dispose()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(true);
            // TODO: uncomment the following line if the finalizer is overridden above.
            // GC.SuppressFinalize(this);
        }
        #endregion
    }
}
