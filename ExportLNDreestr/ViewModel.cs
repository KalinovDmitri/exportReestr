using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Drawing;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using DocsVision.Platform.ObjectManager;
using DocsVision.Platform.ObjectModel;
using DocsVisionContext;

using Excel = Microsoft.Office.Interop.Excel;
using DocsVision.BackOffice.ObjectModel;
using GalaSoft.MvvmLight.CommandWpf;
using GalaSoft.MvvmLight;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Threading;



namespace ExportLNDreestr
{
    public class ViewModel : INotifyPropertyChanged //ViewModelBase
    {
        private DocsVisionContext.DocsVisionContext dvContext = null;
        private ObjectContext Context = null;
        private UserSession Session = null;
        private string _log;
        private int _progress;
        private int _maxProgress = 10;
        public event PropertyChangedEventHandler PropertyChanged;

        public string LogBox
        {
            get { return _log; }
            set
            {
                if (_log == value) return;
                _log = value;
                OnPropertyChanged("LogBox");
            }
        }
        
        public int MaxProgress
        {
            get { return _maxProgress; }
            set
            {
                if (_maxProgress == value) return;
                _maxProgress = value;
                OnPropertyChanged("MaxProgress");
            }
        }
        public int Progress
        {
            get { return _progress; }
            set
            {
                if (_progress == value) return;
                _progress = value;
                OnPropertyChanged("Progress");
            }
        }
        public ViewModel()
        {
           
        }


        private void InicializeContext()
        {
            dvContext = DocsVisionContextFactory.CreateDefault();
            Context = dvContext.CurrentContext;
            Session = dvContext.CurrentSession;
        }

        private ICommand _exportRO;
       


        public ICommand ExportRO
        {
            get
            {
                return _exportRO ?? (_exportRO = new RelayCommand(Start_thread));
            }
        }

        public void Start_thread()
        {
            Thread tr = new Thread(Start_export);
            tr.Start();
            
        }
        private void Start_export()
        {
            string filePath = @"D:\Реетры ЛНД\РЛО.xlsx";
            StreamReader reader = new StreamReader("AllLND.txt");
            string queryXML = reader.ReadToEnd();
            reader.Close();
            LogBox += "Запуск\n";
            InicializeContext();

            List<Guid> IDs = new List<Guid>();
            CardDataCollection coll = Session.CardManager.FindCards(queryXML);
            LogBox += string.Format("{1}|| Найдено {0} карточек ЛНД\n", coll.Count.ToString(), DateTime.Now.ToString("u"));
            CardManager CM = Session.CardManager;
            CardData cd = CM.GetCardData(IDs[0]);
            

            //cd.Sections[CM.CardTypes.]
     
            foreach (CardData el in coll)
            {
                if (!Equals(el.Id, Guid.Empty))
                    IDs.Add(el.Id);
            }

            int counter = 100;// IDs.Count;
            LogBox += string.Format("{1}|| Преобразовано в Guid {0} объектов\n", counter.ToString(), DateTime.Now.ToString("u"));

            ExcelDocument exelDoc = new ExcelDocument(filePath);
            exelDoc.Visible = true;
            int indexRow = 5;
            MaxProgress = counter;
            for (int i = 0; i < counter; i++)
            {
                LogBox += string.Format("{3}|| Начата обработка карточки с ID: {0} {1} из {2}\n", IDs[i].ToString().ToUpper(), i + 1, counter, DateTime.Now.ToString("u"));

                indexRow = ExcelExportCurrentCard(exelDoc, IDs[i], indexRow) + 1;
                LogBox += string.Format("{3}|| Завершена обработка карточки с ID: {0} {1} из {2}\n", IDs[i].ToString().ToUpper(), i + 1, counter, DateTime.Now.ToString("u"));
                Progress = i + 1;
            }


            Context.Dispose();
            Session.Close();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            MessageBox.Show("Выгрузка завершена");
        }
        private int ExcelExportCurrentCard(ExcelDocument exelDoc, Guid CardId, int index)
        {

            Document cardLnd = Context.GetObject<Document>(CardId);
            BaseCardSectionRow additionalPropertiesLNDSection = cardLnd.GetSectionRow("AdditionalPropertiesLND");
            IList<BaseCardSectionRow> hystoryPGSection = (IList<BaseCardSectionRow>)cardLnd.GetSection("HystoryPG");
            IList<BaseCardSectionRow> hystoryLNDSection = (IList<BaseCardSectionRow>)cardLnd.GetSection("HystoryLND");
            IList<BaseCardSectionRow> applicationLNDSection = (IList<BaseCardSectionRow>)cardLnd.GetSection("ApplicationLND");
            System.Drawing.Color color = default(System.Drawing.Color);

            string name = cardLnd.MainInfo.Name;
            string numberPG = string.Empty;
            string datePD = string.Empty;

            StaffUnit owner = Context.GetObject<StaffUnit>(new Guid(additionalPropertiesLNDSection["Owner"].ToString()));//проверить что будет работать при отсутвие владельца
            if (hystoryPGSection.Count > 0)
            {
                BaseCardSectionRow row = hystoryPGSection.FirstOrDefault(r => int.Parse(r["ChangeNumber"].ToString()) == 0);
                if (row != null)
                {
                    numberPG = row["NumberPG"] != null ? row["NumberPG"].ToString() : string.Empty;
                    numberPG = row["NumberPG"]?.ToString() ?? string.Empty;
                    datePD = row["DatePG"] != null ? ((DateTime?)row["DatePG"]).Value.ToString("yyyy") : string.Empty;
                }
            }

            string fileName = additionalPropertiesLNDSection["FileName"] != null ? additionalPropertiesLNDSection["FileName"].ToString() : string.Empty;
            string сotent = cardLnd.MainInfo["Content"] != null ? cardLnd.MainInfo["Content"].ToString() : string.Empty;
            string number = additionalPropertiesLNDSection["Number"] != null ? additionalPropertiesLNDSection["Number"].ToString() : string.Empty;
            string approvalNumber = additionalPropertiesLNDSection["ApprovalNumber"] != null ? additionalPropertiesLNDSection["ApprovalNumber"].ToString() : string.Empty;
            string version = additionalPropertiesLNDSection["Version"] != null ? additionalPropertiesLNDSection["Version"].ToString() : string.Empty;

            string typeLNDid = additionalPropertiesLNDSection["TypeLND"] != null ? additionalPropertiesLNDSection["TypeLND"].ToString() : string.Empty;
            string typeLNDstr = string.Empty;
            if (typeLNDid != string.Empty)
            {
                BaseUniversalItem typeLND = Context.GetObject<BaseUniversalItem>(new Guid(typeLNDid));
                typeLNDstr = typeLND.Name;
            }
            DateTime? commissioningDate = (DateTime?)additionalPropertiesLNDSection["CommissioningDate"];
            string stageName = string.Empty;
            string stateName = string.Empty;
            switch (cardLnd.SystemInfo.State.DefaultName)
            {
                case "ProjectLND":
                    stageName = "Проект";
                    stateName = "Не действует";
                    color = System.Drawing.Color.Blue;
                    break;
                case "Approved":
                    stageName = "Утвержден";
                    stateName = "Действует";
                    break;
                case "Cancelled":
                    stageName = "Утратил силу";
                    stateName = "Не действует";
                    color = System.Drawing.Color.Red;
                    break;
                case "NotValid":
                    stageName = "Утратил силу";
                    stateName = "Не действует";
                    color = System.Drawing.Color.Red;
                    break;
                default:
                    stageName = "Значение не попало в ожидаемы диапазон";
                    stateName = "Значение не попало в ожидаемы диапазон";
                    break;
            }

            string classificationId = additionalPropertiesLNDSection["Classification"] != null ? additionalPropertiesLNDSection["Classification"].ToString() : string.Empty;
            string classificationStr = string.Empty;
            if (classificationId != string.Empty)
            {
                BaseUniversalItem classification = Context.GetObject<BaseUniversalItem>(new Guid(classificationId));
                classificationStr = classification != null ? classification.Name : string.Empty;
            }

            //ввод
            string nameRDvv = string.Empty;
            string namberRDvv = string.Empty;
            DateTime? regDateRDvv = null;
            string fileNameRDvv = string.Empty;
            string typeRDvv = string.Empty;
            //отмена
            string nameRDotm = string.Empty;
            string namberRDotm = string.Empty;
            DateTime? regDateRDotm = null;
            string fileNameRDotm = string.Empty;
            string typeRDotm = string.Empty;
            List<BaseCardSectionRow> rowsAct = new List<BaseCardSectionRow>();
            IEnumerable<BaseCardSectionRow> rowsAct1;

            if (hystoryLNDSection.Count > 0)
            {
                BaseCardSectionRow rowVV = hystoryLNDSection.FirstOrDefault(r => (Context.GetObject<BaseUniversalItem>(new Guid(r["Type"].ToString())).Name == "РД Общества о вводе"));
                if (rowVV != null)
                {
                    Document rdOvv = Context.GetObject<Document>(new Guid(rowVV["RDId"].ToString()));
                    if (rdOvv != null)
                        GetDataFromRD(rdOvv, out nameRDvv, out namberRDvv, out fileNameRDvv, out regDateRDvv, out typeRDvv);

                }
                BaseCardSectionRow rowOtm = hystoryLNDSection.FirstOrDefault(r => (Context.GetObject<BaseUniversalItem>(new Guid(r["Type"].ToString())).Name == "РД Общества об отмене"));
                if (rowOtm != null)
                {
                    Document rdOtm = Context.GetObject<Document>(new Guid(rowOtm["RDId"].ToString()));
                    if (rdOtm != null)
                        GetDataFromRD(rdOtm, out nameRDotm, out namberRDotm, out fileNameRDotm, out regDateRDotm, out typeRDotm);

                }

                rowsAct1 = hystoryLNDSection.Where(r =>
                (Context.GetObject<BaseUniversalItem>(new Guid(r["Type"].ToString())).Name == "РД Общества об актуализации")).OrderBy(r => int.Parse(r["Number"].ToString()));
                foreach (BaseCardSectionRow row in rowsAct1)
                {
                    rowsAct.Add(row);
                }
            }




            exelDoc.SetCellValue("ПАО \"Самаранефтехимроект\"", index, 1, color);// Юридическое лицо
            exelDoc.SetCellValue(owner != null ? owner.FullName : string.Empty, index, 2, color);
            exelDoc.SetCellValue(numberPG, index, 3, color); // Номер в ПГ
            exelDoc.SetCellValue(datePD, index, 4, color); // Год планирования
            exelDoc.SetCellValue(name, index, 5, color); // Наименование ЛНД/ Приложения
            exelDoc.SetCellValue(fileName, index, 6, color); // Имя файла ЛНД/приложения/ редакции ЛНД с изм.
            exelDoc.SetCellValue(сotent, index, 7, color); // Описание
            exelDoc.SetCellValue(number, index, 8, color); // Регистрационный номер ЛНД
            exelDoc.SetCellValue(approvalNumber, index, 9, color); // Номер утверждения ЛНД
            exelDoc.SetCellValue(version, index, 10, color); // Версия ЛНД
            exelDoc.SetCellValue(typeLNDstr, index, 11, color); // Вид ЛНД
            exelDoc.SetCellValue("Нормативный", index, 12, color); // Тип документа
            exelDoc.SetCellValue(commissioningDate != null ? commissioningDate.Value.ToString("d") : string.Empty, index, 13, color); // дата
            exelDoc.SetCellValue(stageName, index, 14, color); // Стадия ЛНД
            exelDoc.SetCellValue(string.Empty, index, 15, color); // Срок действия ЛНД
            exelDoc.SetCellValue(stateName, index, 16, color); // Статус ЛНД 
            exelDoc.SetCellValue(classificationStr, index, 17, color); // Бизнес-процесс 1-го и 2-го уровней
            exelDoc.SetCellValue(string.Empty, index, 18, color); // Комментарии
            //ввод
            exelDoc.SetCellValue(nameRDvv, index, 19, color); // Наименование РД
            exelDoc.SetCellValue(fileNameRDvv, index, 20, color); // Имя файла РД
            exelDoc.SetCellValue(namberRDvv, index, 21, color); // Номер РД
            exelDoc.SetCellValue(regDateRDvv != null ? regDateRDvv.Value.ToString("d") : string.Empty, index, 22, color); // Дата РД
            exelDoc.SetCellValue(typeRDvv, index, 23, color); // Номер РД

            //отмена
            exelDoc.SetCellValue(nameRDotm, index, 29, color); // Наименование РД
            exelDoc.SetCellValue(fileNameRDotm, index, 30, color); // Имя файла РД
            exelDoc.SetCellValue(namberRDotm, index, 31, color); // Номер РД
            exelDoc.SetCellValue(regDateRDotm != null ? regDateRDvv.Value.ToString("d") : string.Empty, index, 32, color); // Дата РД
            exelDoc.SetCellValue(typeRDotm, index, 33, color); // Номер РД
            // изменения
            if (rowsAct.Count() > 0)
            {
                for (int i = 0; i < rowsAct.Count; i++)
                {
                    index++;//переход на новую строку
                    System.Drawing.Color colorIzm = System.Drawing.Color.Green;
                    int numberIzm = int.Parse(rowsAct[i]["Number"].ToString());
                    string nameRDact = string.Empty;
                    string numberRDact = string.Empty;
                    DateTime? regDateRDact = null;
                    string fileNameRDact = string.Empty;
                    string typeRDact = string.Empty;
                    string numberPGizm = string.Empty;
                    string datePDizm = string.Empty;

                    if (hystoryPGSection.Count > 0)
                    {
                        BaseCardSectionRow row = hystoryPGSection.FirstOrDefault(r => int.Parse(r["ChangeNumber"].ToString()) == numberIzm);
                        numberPGizm = row != null ? row["NumberPG"].ToString() : string.Empty;
                        datePDizm = row != null ? ((DateTime?)row["DatePG"]).Value.ToString("yyyy") : string.Empty;
                    }

                    exelDoc.SetCellValue("ПАО \"Самаранефтехимроект\"", index, 1, colorIzm);// Юридическое лицо
                    exelDoc.SetCellValue(numberPGizm, index, 3, colorIzm); // Номер в ПГ
                    exelDoc.SetCellValue(datePDizm, index, 4, colorIzm); // Год планирования
                    exelDoc.SetCellValue(string.Format("Изменение {0} к \"{1}\"", numberIzm, name), index, 5, colorIzm); // Наименование ЛНД/ Приложения

                    exelDoc.SetCellValue(numberIzm.ToString(), index, 8, colorIzm); // Регистрационный номер ЛНД

                    exelDoc.SetCellValue("Изменение", index, 12, colorIzm); // Тип документа
                    exelDoc.SetCellValue((DateTime?)(rowsAct[i]["Date"]) != null ? ((DateTime?)(rowsAct[i]["Date"])).Value.ToString("d") : string.Empty, index, 13, colorIzm); // дата
                    exelDoc.SetCellValue(classificationStr, index, 17, colorIzm); // Бизнес-процесс 1-го и 2-го уровней
                    Document RDact = Context.GetObject<Document>(new Guid(rowsAct[i]["RDId"].ToString()));
                    if (RDact != null)
                    {
                        GetDataFromRD(RDact, out nameRDact, out numberRDact, out fileNameRDact, out regDateRDact, out typeRDact);
                    }
                    exelDoc.SetCellValue(nameRDact, index, 19, colorIzm); // Наименование РД
                    exelDoc.SetCellValue(fileNameRDact, index, 20, colorIzm); // Имя файла РД
                    exelDoc.SetCellValue(numberRDact, index, 21, colorIzm); // Номер РД
                    exelDoc.SetCellValue(regDateRDact != null ? regDateRDvv.Value.ToString("d") : string.Empty, index, 22, colorIzm); // Дата РД
                    exelDoc.SetCellValue(typeRDact, index, 23, colorIzm); // Номер РД

                }
            }
            foreach (BaseCardSectionRow row in applicationLNDSection)
            {

                if (row != null)
                {
                    index++;// следующая строка после изменений

                    string namePril = row["Name"] != null ? row["Name"].ToString() : string.Empty;
                    string fileNamePril = row["FileName"] != null ? row["FileName"].ToString() : string.Empty;

                    exelDoc.SetCellValue("ПАО \"Самаранефтехимроект\"", index, 1, color);// Юридическое лицо
                    exelDoc.SetCellValue(owner != null ? owner.FullName : string.Empty, index, 2, color);
                    exelDoc.SetCellValue(namePril, index, 5, color); // Наименование ЛНД/ Приложения
                    exelDoc.SetCellValue(fileNamePril, index, 6, color); // Имя файла ЛНД/приложения/ редакции ЛНД с изм.
                    exelDoc.SetCellValue(name, index, 7, color); // Описание
                    exelDoc.SetCellValue(number, index, 8, color); // Регистрационный номер ЛНД
                    exelDoc.SetCellValue(approvalNumber, index, 9, color); // Номер утверждения ЛНД
                    exelDoc.SetCellValue(version, index, 10, color); // Версия ЛНД
                    exelDoc.SetCellValue(typeLNDstr, index, 11, color); // Вид ЛНД
                    exelDoc.SetCellValue("Приложение", index, 12, color); // Тип документа
                    exelDoc.SetCellValue(commissioningDate != null ? commissioningDate.Value.ToString("d") : string.Empty, index, 13, color); // дата
                    exelDoc.SetCellValue(stageName, index, 14, color); // Стадия ЛНД
                    exelDoc.SetCellValue(string.Empty, index, 15, color); // Срок действия ЛНД
                    exelDoc.SetCellValue(stateName, index, 16, color); // Статус ЛНД 
                    exelDoc.SetCellValue(classificationStr, index, 17, color); // Бизнес-процесс 1-го и 2-го уровней
                    exelDoc.SetCellValue(string.Empty, index, 18, color); // Комментарии

                    //ввод
                    exelDoc.SetCellValue(nameRDvv, index, 19, color); // Наименование РД
                    exelDoc.SetCellValue(fileNameRDvv, index, 20, color); // Имя файла РД
                    exelDoc.SetCellValue(namberRDvv, index, 21, color); // Номер РД
                    exelDoc.SetCellValue(regDateRDvv != null ? regDateRDvv.Value.ToString("d") : string.Empty, index, 22, color); // Дата РД
                    exelDoc.SetCellValue(typeRDvv, index, 23, color); // Номер РД

                    //отмена
                    exelDoc.SetCellValue(nameRDotm, index, 29, color); // Наименование РД
                    exelDoc.SetCellValue(fileNameRDotm, index, 30, color); // Имя файла РД
                    exelDoc.SetCellValue(namberRDotm, index, 31, color); // Номер РД
                    exelDoc.SetCellValue(regDateRDotm != null ? regDateRDvv.Value.ToString("d") : string.Empty, index, 32, color); // Дата РД
                    exelDoc.SetCellValue(typeRDotm, index, 33, color); // Номер РД


                }
            }

            return index;

        }
        private void GetDataFromRD(Document rd, out string name, out string number, out string fileName, out DateTime? dateRD, out string typeRD)
        {
            BaseCardSectionRow for_LNDSection = rd.GetSectionRow("For_LND");
            name = rd.MainInfo.Name;
            number = rd.Numbers[0].Number;
            dateRD = (DateTime?)rd.MainInfo["RegDate"];
            fileName = for_LNDSection != null ? (for_LNDSection["FileNameRD"] != null ? for_LNDSection["FileNameRD"].ToString() : string.Empty) : string.Empty;
            typeRD = rd.SystemInfo.CardKind.Name;
        }


        public virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    // Класс  для работы с документом Excel
    public class ExcelDocument
    {
        #region Properties
        private Excel.Application _application = null;
        private Excel.Workbook _workBook = null;
        private Excel.Worksheet _workSheet = null;
        private object _missingObj = System.Reflection.Missing.Value;
        #endregion

        //КОНСТРУКТОР
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

        // ВСТАВКА ЗНАЧЕНИЯ В ЯЧЕЙКУ
        /// <summary>
        /// Установка значений в ячейку
        /// </summary>
        /// <param name="cellValue">Значение ячейки</param>
        /// <param name="rowIndex">Индекс строки</param>
        /// <param name="columnIndex">Индекс столбца</param>
        public void SetCellValue(string cellValue, int rowIndex, int columnIndex, System.Drawing.Color color = default(System.Drawing.Color))
        {
            Excel.Range rng2 = _workSheet.Cells[rowIndex, columnIndex] as Excel.Range;
            rng2.Value2 = cellValue;
            rng2.Font.Color = color;
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
