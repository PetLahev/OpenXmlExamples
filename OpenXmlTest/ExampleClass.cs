using System;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using openXmlSpreadsheet = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using System.IO;

namespace OpenXmlTest
{
    /// <summary> 
    /// Performs a check to find out if given chart id exists in given file
    /// </summary>
    public class ExampleClass : IDisposable
    {
        public delegate void ReportProgressDel(string message);
        public event ReportProgressDel ReportProgress;

        private Stream _docStream;
        private SpreadsheetDocument _openXmlDoc;
        private WorkbookPart _wkb;
        private Worksheet _wks;
        private readonly string _filePath;
        private readonly int _sheetId;
        private readonly string _chartId;

        /// <summary>
        /// Opens the given file as a stream and checks if given sheet and chart ID exists
        /// </summary>
        /// <param name="filePath">a file to be checked (can be opened by Excel already)</param>
        /// <param name="sheetId">the ID of sheet where to look for a chart</param>
        /// <param name="chartId">the chart ID to look for on given sheet</param>
        public ExampleClass(string filePath, int sheetId, string chartId)
        {
            if (string.IsNullOrWhiteSpace(filePath) || sheetId <= 0)
            {
                throw new ArgumentException("Either the file path or the sheet ID was not provided");                
            }

            _filePath = filePath;
            _sheetId = sheetId;
            _chartId = chartId;
            OpenFile();
        }

        /// <summary> Opens the given file via the DocumentFormat package </summary>
        /// <returns></returns>
        private bool OpenFile()
        {
            
            try
            {
                _docStream = new FileStream(_filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);                    
                _openXmlDoc = SpreadsheetDocument.Open(_docStream, false);
                _wkb = _openXmlDoc.WorkbookPart;
                ReportProgress?.Invoke($"File opened");
                return true;
            }
            catch (Exception ex)
            {
                ReportProgress?.Invoke($"File couldn't be opened!\n{ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Applies for Office 2010 and higher 
        /// True if the chart check should be considered just for a numeric ID, false for GUID check
        /// </summary>
        public bool SimpleChartID { get; set; }

        /// <summary> Returns true if sheet with given ID exists </summary>
        public bool SheetExists
        {
            get
            {
                return _wkb.Workbook.Descendants<Sheet>().Where(s => s.SheetId == _sheetId).FirstOrDefault() != null;
            }
        }

        /// <summary> Returns true if chart with given ID exists </summary>
        public bool ChartExists
        {
            get
            {
                if (string.IsNullOrWhiteSpace(_chartId)) return false;
                if (_wks== null)
                {                    
                    CheckSheetExists();
                }

                return CheckChartExists();
            }
        }

        /// <summary> Finds the Sheet based on given sheet ID </summary>        
        /// <returns>True if exists, otherwise false</returns>
        private bool CheckSheetExists()
        {
            Sheet theSheet = _wkb.Workbook.Descendants<Sheet>().Where(s => s.SheetId ==_sheetId).FirstOrDefault();
            if (theSheet == null)
            {
                ReportProgress?.Invoke("Couldn't find the worksheet");
                return false;
            }                
                        
            // Retrieve a reference to the worksheet part.
            WorksheetPart wsPart = (WorksheetPart)(_wkb.GetPartById(theSheet.Id));
            _wks = wsPart.Worksheet;            
            ReportProgress?.Invoke($"Sheet found: {theSheet.Name} ({theSheet.Id})");
            return true;
        }

        /// <summary>
        /// Check if given chart ID exists
        /// For Office 2007 it simply checks the given ID which is a number
        /// For Office 2010 it check the creation GUID        
        /// </summary>        
        /// <returns>True if a chart exists on given sheet otherwise false </returns>
        private bool CheckChartExists()
        {            
            var theDrawing = _wks.Descendants<Drawing>().FirstOrDefault();
            if (theDrawing == null)
            {
                ReportProgress?.Invoke("No chart on given worksheet");
                return false;
            }

            var drawingPart = _wks.WorksheetPart.GetPartById(theDrawing.Id);
            openXmlSpreadsheet.WorksheetDrawing wksDrawing = (openXmlSpreadsheet.WorksheetDrawing)drawingPart.RootElement;

            ReportProgress?.Invoke("Drawing part found, checking chart ID");
            // the chartObjects are stored in object called TwoCellAnchor
            // the Drawing part will have as many TwoCellAnchor objects as many chart objects are stored on the sheet
            foreach (var item in wksDrawing.ChildElements)
            {
                openXmlSpreadsheet.TwoCellAnchor twoCellAnchor = (openXmlSpreadsheet.TwoCellAnchor)item;
                // need to get GraphicFrame object where the Chart id is stored
                var graphicFrame = twoCellAnchor.GetFirstChild<openXmlSpreadsheet.GraphicFrame>();

                /*  
                 *  The chart Id is stored in the frame property but the ID may not be unique if someone
                 *  1, remove the chart e.g. with Id=1
                 *  2, save and close the file
                 *  3, re-open the file and insert a new chart
                 *  4, the new chart will get the Id=1
                 *  In that case we cannot be sure  the ID is unique
                */
                dynamic frameProps = graphicFrame.Descendants<openXmlSpreadsheet.NonVisualDrawingProperties>().First();

                /*
                 * Since Office 2010 a new property was added called 'creationId' which has a globally unique ID (GUID)
                 * This is more safe but available only in 2010 and higher
                */
                try
                {
                    dynamic frameProps2010 = graphicFrame.Descendants<DocumentFormat.OpenXml.Drawing.NonVisualDrawingPropertiesExtension>().First();

                    if (SimpleChartID)
                    {
                        if (_chartId.Equals((string)frameProps.Id)) return true;
                    }
                    else
                    {
                        var creationElement = frameProps2010.FirstChild.ExtendedAttributes;
                        string guid = creationElement[0].Value;
                        if (_chartId.Equals(guid)) return true;
                    }
                }
                catch (Exception)
                {                    
                    if (_chartId.Equals((string)frameProps.Id)) return true;
                }                
            }
            
            return false;
        }

        public void Dispose()
        {
            _openXmlDoc?.Close();
            _docStream?.Close();            
        }
    }
}
