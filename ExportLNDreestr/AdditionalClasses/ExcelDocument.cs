using Excel = Microsoft.Office.Interop.Excel;

namespace ExportLNDreestr.AdditionalClasses
{
    public class ExcelDocument
    {
        #region Properties
        private Excel.Application _application = null;
        private Excel.Workbook _workBook = null;
        private Excel.Worksheet _workSheet = null;
        private object _missingObj = System.Reflection.Missing.Value;
        #endregion
        #region Constructor
        public ExcelDocument()
        {
            _application = new Excel.Application();
            _workBook = _application.Workbooks.Add(_missingObj);
            _workSheet = _workBook.Worksheets[1] as Excel.Worksheet;
        }

        public ExcelDocument(string pathToTemplate)
        {
            _application = new Excel.Application();
            _workBook = _application.Workbooks.Open(pathToTemplate);
            _workSheet = _workBook.Worksheets[1] as Excel.Worksheet;
        }
        #endregion

        #region Methods
        // ВИДИМОСТЬ ДОКУМЕНТА
        /// <summary>
        /// Устанавливает видимость документа
        /// </summary>
        public bool Visible
        {
            get
            {
                return _application.Visible;
            }
            set
            {
                _application.Visible = value;
            }
        }
       
        /// <summary>
        /// Установка значений в ячейку
        /// </summary>
        /// <param name="cellValue">Значение ячейки</param>
        /// <param name="rowIndex">Индекс строки</param>
        /// <param name="columnIndex">Индекс столбца</param>
        /// <param name="color">Цвет шрифта в ячейке</param>
        /// <param name="isItalic">Курсив</param>
        public void SetCellValue(string cellValue, int rowIndex, int columnIndex, System.Drawing.Color color = default(System.Drawing.Color), bool isItalic = false)
        {
            Excel.Range rng2 = _workSheet.Cells[rowIndex, columnIndex] as Excel.Range;
            rng2.Value2 = cellValue;
            rng2.Font.Color = color;
            rng2.Font.Italic = isItalic;
            rng2.VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            //rng2.AutoFit();
        }
        /// <summary>
        /// Закрытие приложения
        /// </summary>
        public void Close()
        {
            _workBook.Close(false, _missingObj, _missingObj);

            _application.Quit();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(_application);

            _application = null;
            _workBook = null;
            _workSheet = null;

            System.GC.Collect();
        }
        #endregion
    }
}
