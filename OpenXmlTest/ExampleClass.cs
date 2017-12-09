using System;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using openXmlSpreadsheet = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OpenXmlTest
{
    public class ExampleClass
    {
        public delegate void ReportProgressDel(string message);
        public event ReportProgressDel ReportProgress;

        private WorkbookPart _wkb;
        private Worksheet _wks;
        private readonly string _filePath;
        private readonly int _sheetId;
        private readonly string _chartId;

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
                var file = SpreadsheetDocument.Open(_filePath, false);
                _wkb = file.WorkbookPart;
                ReportProgress?.Invoke($"File opened");
                return true;
            }
            catch (Exception ex)
            {
                ReportProgress?.Invoke($"File couldn't be opened!\n{ex.Message}");
                return false;
            }            
        }

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
        /// 
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
                    var creationElement = frameProps2010.FirstChild.ExtendedAttributes;
                    string guid = creationElement[0].Value;                    
                    if (_chartId.Equals(guid)) return true;
                }
                catch (Exception)
                {                    
                    if (_chartId.Equals((string)frameProps.Id)) return true;
                }                
            }
            
            return false;
        }
        
    }
}
