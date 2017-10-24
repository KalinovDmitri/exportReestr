﻿using System;
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
//using System.Windows.Shapes;

using DocsVision.Platform.ObjectManager;
using DocsVision.Platform.ObjectModel;
using DocsVisionContext;


using DocsVision.BackOffice.ObjectModel;
using GalaSoft.MvvmLight.CommandWpf;
using GalaSoft.MvvmLight;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Threading;

using ExportLNDreestr.AdditionalClasses;
using Newtonsoft.Json;

namespace ExportLNDreestr
{
    public class ViewModel : INotifyPropertyChanged //ViewModelBase
    {
        const string PathQueryLNDO = @"Sourse\AllLNDO.txt";
        const string PathQueryLNDK = @"Sourse\AllLNDK.txt";
        const string PathWithTemplateLNDo = @"\Sourse\РЛО.xlsx";
        const string PathWithTemplateLNDk = @"\Sourse\РЛК.xlsx";


        #region Fields
        private DocsVisionContext.DocsVisionContext dvContext = null;
        private UserSession Session = null;
        private string _log;
        private bool _isEnableButtonRLO = true;
        private bool _isEnableButtonRLK = true;
        private bool _isEnableButtonCancel = false;
        private int _progress;
        private int _maxProgress = 10;//первичное значение прогрессбара
        private string _userName;
        private string _password;
        private string _serverName;
        private string _connectionString;
        public event PropertyChangedEventHandler PropertyChanged;
        CancellationTokenSource cancelTokenSource;
        #endregion Fields
        #region PropertiesForView
        public string LogBox
        {
            get { return _log; }
            set
            {
                _log += string.Format("{0}|| {1}\n", DateTime.Now.ToString("u"), value);
                OnPropertyChanged(nameof(LogBox));
            }
        }
        public bool IsEnableButtonRLO
        {
            get { return _isEnableButtonRLO; }
            set
            {
                if (_isEnableButtonRLO == value) return;
                _isEnableButtonRLO = value;
                
                OnPropertyChanged(nameof(IsEnableButtonRLO));
            }
        }
        public bool IsEnableButtonRLK
        {
            get { return _isEnableButtonRLK; }
            set
            {
                if (_isEnableButtonRLK == value) return;
                _isEnableButtonRLK = value;

                OnPropertyChanged(nameof(IsEnableButtonRLK));
            }
        }
        public bool IsEnableButtonCancel
        {
            get { return _isEnableButtonCancel; }
            set
            {
                if (_isEnableButtonCancel == value) return;
                _isEnableButtonCancel = value;

                OnPropertyChanged(nameof(IsEnableButtonCancel));
            }
        }
        public int MaxProgress
        {
            get { return _maxProgress; }
            set
            {
                if (_maxProgress == value) return;
                _maxProgress = value;
                OnPropertyChanged(nameof(MaxProgress));
            }
        }
        public int Progress
        {
            get { return _progress; }
            set
            {
                if (_progress == value) return;
                _progress = value;
                OnPropertyChanged(nameof(Progress));
            }
        }
        public string ServerName
        {
            get { return _serverName; }
            set
            {
                if (_serverName == value) return;
                _serverName = value;
                ConnectionString = string.Format("http://{0}/docsvision/StorageServer/StorageServerService.asmx", value);
                OnPropertyChanged(nameof(ServerName));
            }
        }
        public string ConnectionString
        {
            get { return _connectionString; }
            set
            {

                if (_connectionString == value) return;
                _connectionString = value;
                OnPropertyChanged(nameof(ConnectionString));

            }
        }
        public string UserName
        {
            get { return _userName; }
            set
            {
                if (_userName == value) return;
                _userName = value;
                OnPropertyChanged(nameof(UserName));
            }
        }
        public string Password
        {
            get { return _password; }
            set
            {
                if (_password == value) return;
                _password = value;
                OnPropertyChanged(nameof(Password));
            }
        }
        #endregion PropertiesForView
        #region Properties
        private CardData _refBaseUniversalCD = null;
        private CardData RefBaseUniversalCD
        {
            get
            {
                if (_refBaseUniversalCD == null)
                {
                    _refBaseUniversalCD = Session.CardManager.GetDictionaryData(Guid.Parse("4538149D-1FC7-4D41-A104-890342C6B4F8"));
                }
                return _refBaseUniversalCD;
            }
        }

        private CardData _refState = null;
        private CardData RefState
        {
            get
            {
                if (_refState == null)
                {
                    _refState = Session.CardManager.GetDictionaryData(new Guid("443F55F0-C8AB-4DD3-BCBD-5328C7C9D385"));
                }
                return _refState;
            }

        }
        private CardData _refKinds = null;
        private CardData RefKinds
        {
            get
            {
                if (_refKinds == null)
                {
                    _refKinds = Session.CardManager.GetCardData(new Guid("8F704E7D-A123-4917-94B4-F3B851F193B2"));
                }
                return _refKinds;
            }

        }
        private SectionData _sectionUnitsFromRefStaff = null;
        private SectionData SectionUnitsFromRefStaff
        {
            get
            {
                if (_sectionUnitsFromRefStaff == null)
                {
                    CardData RefStaffCD = Session.CardManager.GetDictionaryData(Guid.Parse("6710B92A-E148-4363-8A6F-1AA0EB18936C"));
                    _sectionUnitsFromRefStaff = RefStaffCD.Sections[Guid.Parse("7473F07F-11ED-4762-9F1E-7FF10808DDD1")];
                }
                return _sectionUnitsFromRefStaff;
            }
        }
        #endregion Properties
        #region Commands
        private ICommand _exportRO;
        public ICommand ExportRO
        {
            get
            {
                return _exportRO ?? (_exportRO = new RelayCommand(StartExportLNDO));
            }
        }
        private ICommand _exportRK;
        public ICommand ExportRK
        {
            get
            {
                return _exportRK ?? (_exportRK = new RelayCommand(StartExportLNDK));
            }
        }
        private ICommand _cancel;
        public ICommand CancelCommand
        {
            get
            {
                return _cancel ?? (_cancel = new RelayCommand(Cancel));

            }
        }
        private ICommand _saveConnectionSettings;
        public ICommand SaveConnectionSettings
        {
            get
            {
                return _saveConnectionSettings ?? (_saveConnectionSettings = new RelayCommand(SaveSettings));

            }
        }


        #endregion Commands
        #region Constructors
        public ViewModel()
        {
            using (StreamReader stream = new StreamReader(Directory.GetCurrentDirectory() + @"\Sourse\ConnectionSettings.json"))
            {
                string str = stream.ReadToEnd();
                ConnectionSettings connect = JsonConvert.DeserializeObject<ConnectionSettings>(str);
                ServerName = connect.servername;
                UserName = connect.username;
                Password = connect.password;
            }

        }
        public ViewModel(UserSession _session)
        {
            Session = _session;
        }
        #endregion Constructors
        #region Methods
        private bool InicializeContext()
        {
            bool IsConnected = false;
            if (dvContext == null)
            {
                try
                {
                    dvContext = DocsVisionContextFactory.CreateContext(ConnectionString, "docsvision", UserName, Password);
                    Session = dvContext.CurrentSession;
                    IsConnected = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            return IsConnected;
        }
        private void SaveSettings()
        {
            ConnectionSettings setings = new ConnectionSettings
            {
                password = Password,
                username = UserName,
                servername = ServerName
            };

            using (StreamWriter stream = new StreamWriter(Directory.GetCurrentDirectory() + @"\Sourse\ConnectionSettings.json", false))
            {
                string ser = JsonConvert.SerializeObject(
                    new ConnectionSettings
                    {
                        password = Password,
                        username = UserName,
                        servername = ServerName
                    });

                stream.WriteLine(ser);
                stream.Flush();
                stream.Close();
            }
            MessageBox.Show("Настройки успешно сохранены", "Сообщение", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        
        public void Cancel()
        {
            if (cancelTokenSource != null)
            {
                LogBox = "Поступила команда отмены операции";
                cancelTokenSource.Cancel();
                IsEnableButtonRLO = true;
                IsEnableButtonRLK = true;
                IsEnableButtonCancel = false;
            }
        }
        public void StartExportLNDO()
        {
            if (InicializeContext())
            {
                LogBox = "Соединенияе с сервером установлено";
            }
            else
            {
                return;
            }
            cancelTokenSource = new CancellationTokenSource();
            CancellationToken token = cancelTokenSource.Token;
            System.Threading.Tasks.Task.Run(() => ExportLND(GetFullPathToCopyTemplate(PathWithTemplateLNDo), PathQueryLNDO, token, ExcelExportCurrentCardFromCardDataO));
        }
        public void StartExportLNDK()
        {
            if (InicializeContext())
            {
                LogBox = "Соединенияе с сервером установлено";
            }
            else
            {
                return;
            }
            cancelTokenSource = new CancellationTokenSource();
            CancellationToken token = cancelTokenSource.Token;
            System.Threading.Tasks.Task.Run(() => ExportLND(GetFullPathToCopyTemplate(PathWithTemplateLNDk), PathQueryLNDK, token, ExcelExportCurrentCardFromCardDataK));
        }

        delegate int ExportCardDate(ExcelDocument excelDoc, Guid CardID, int IndexRow);
   
        /*private void ExportLNDO(CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            
            IsEnableButtonRLO = false;
            IsEnableButtonRLK = false;
            IsEnableButtonCancel = true;
            string filePath = Directory.GetCurrentDirectory() + @"\Sourse\РЛО.xlsx";
            StreamReader reader = new StreamReader(@"Sourse\AllLNDO.txt");
            string queryXML = reader.ReadToEnd();
            reader.Close();
            LogBox = "Запуск выгрузки реестра ЛНД Общества. Поиск Карточек ЛНД...";
           

            List<Guid> IDs = new List<Guid>();
            CardDataCollection coll = Session.CardManager.FindCards(queryXML);
            LogBox = string.Format("Найдено {0} карточек ЛНД", coll.Count.ToString());

            foreach (CardData el in coll)
            {
                if (!Equals(el.Id, Guid.Empty))
                    IDs.Add(el.Id);
            }

            int counter =  IDs.Count;
            //LogBox = string.Format("Преобразовано в Guid {0} объектов", counter.ToString());

            ExcelDocument exelDoc = new ExcelDocument(filePath);
            exelDoc.Visible = true;
            int indexRow = 5;
            MaxProgress = counter;

            for (int i = 0; i < counter; i++)
            {
                if (!token.IsCancellationRequested)
                {
                    LogBox = string.Format("Начата обработка карточки с ID: {0} {1} из {2}", IDs[i].ToString().ToUpper(), i + 1, counter);

                    indexRow = ExcelExportCurrentCardFromCardDataO(exelDoc, IDs[i], indexRow) + 1;
                    LogBox = string.Format("Завершена обработка карточки с ID: {0} {1} из {2}", IDs[i].ToString().ToUpper(), i + 1, counter);
                    Progress = i + 1;
                }
                else
                {
                    if (exelDoc != null)
                    {
                        exelDoc.Close();
                        IsEnableButtonRLO = true;
                        IsEnableButtonRLK = true;
                        IsEnableButtonCancel = false;
                        break;
                    }

                }
            }
            
            //Session.Close();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            Progress = 0;
            MessageBox.Show("Выгрузка завершена");
            IsEnableButtonRLO = true;
            IsEnableButtonRLK = true;
            IsEnableButtonCancel = false;
        }*/
        private void ExportLND(string pathWithTemplate, string pathQuery, CancellationToken token, ExportCardDate export)
        {
                 
            token.ThrowIfCancellationRequested();

            IsEnableButtonRLO = false;
            IsEnableButtonRLK = false;
            IsEnableButtonCancel = true;
         
            StreamReader reader = new StreamReader(pathQuery);
            string queryXML = reader.ReadToEnd();
            reader.Close();
            LogBox = "Запуск выгрузки реестра ЛНД. Поиск Карточек ЛНД...";

            List<Guid> IDs = new List<Guid>();
            CardDataCollection coll = Session.CardManager.FindCards(queryXML);
            LogBox = string.Format("Поиск завершен. Найдено {0} карточек ЛНД", coll.Count.ToString());
            
            foreach (CardData el in coll)
            {
                if (!Equals(el.Id, Guid.Empty))
                    IDs.Add(el.Id);
            }
            int counter = IDs.Count;
            ExcelDocument exelDoc = new ExcelDocument(pathWithTemplate);
            exelDoc.Visible = true;
            int indexRow = 9;
            MaxProgress = counter;

            for (int i = 0; i < counter; i++)
            {
                if (!token.IsCancellationRequested)
                {
                    LogBox = string.Format("Начата обработка карточки с ID: {0} {1} из {2}", IDs[i].ToString().ToUpper(), i + 1, counter);

                    indexRow = export(exelDoc, IDs[i], indexRow) + 1;
                    LogBox = string.Format("Завершена обработка карточки с ID: {0} {1} из {2}", IDs[i].ToString().ToUpper(), i + 1, counter);
                    Progress = i + 1;
                }
                else
                {
                    if (exelDoc != null)
                    {
                        exelDoc.Close();
                        break;
                    }

                }
            }

            //Session.Close();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            Progress = 0;
            MessageBox.Show("Выгрузка завершена");
            IsEnableButtonRLO = true;
            IsEnableButtonRLK = true;
            IsEnableButtonCancel = false;
        }

        private int ExcelExportCurrentCardFromCardDataK(ExcelDocument exelDoc, Guid guid, int indexRow)
        {
            exelDoc.SetCellValue("раз", indexRow, 1);
            return indexRow;
            //throw new NotImplementedException();
            
        }


        private int ExcelExportCurrentCardFromCardDataO(ExcelDocument exelDoc, Guid CardId, int index)
        {
            CardManager CM = Session.CardManager;
            CardData CardDataLND = CM.GetCardData(CardId);
           
            RowData additionalPropertiesLNDSection = CardDataLND.Sections[CardDataLND.Type.Sections["AdditionalPropertiesLND"].Id].FirstRow;
            RowData systemLNDSection = CardDataLND.Sections[CardDataLND.Type.Sections["System"].Id].FirstRow;
            RowData mainInfoLND = CardDataLND.Sections[CardDataLND.Type.Sections["MainInfo"].Id].FirstRow;
      
            RowDataCollection hystoryPGSection = CardDataLND.Sections[CardDataLND.Type.Sections["HystoryPG"].Id].Rows;
            RowDataCollection hystoryLNDSection = CardDataLND.Sections[CardDataLND.Type.Sections["HystoryLND"].Id].Rows;
            RowDataCollection applicationLNDSection = CardDataLND.Sections[CardDataLND.Type.Sections["ApplicationLND"].Id].Rows;

            System.Drawing.Color color = default(System.Drawing.Color);

            string name = mainInfoLND["Name"]!=null ? mainInfoLND["Name"].ToString() : string.Empty;
            string numberPG = string.Empty;
            string datePD = string.Empty;
            Guid stateID = systemLNDSection["State"] != null ? new Guid(systemLNDSection["State"].ToString()) : Guid.Empty;

            RowData statesSprState = RefState.Sections[new Guid("521B4477-DD10-4F57-A453-09C70ADB7799")].GetRow(stateID);
            string owner = GetFullNameOfUnit(additionalPropertiesLNDSection["Owner"].ToString());
            if (hystoryPGSection.Count > 0)
            {
                RowData row = hystoryPGSection.FirstOrDefault(r => int.Parse(r["ChangeNumber"].ToString()) == 0);
                if (row != null)
                {
                    numberPG = row["NumberPG"] != null ? row["NumberPG"].ToString() : string.Empty;
                    numberPG = row["NumberPG"]?.ToString() ?? string.Empty;
                    datePD = row["DatePG"] != null ? ((DateTime?)row["DatePG"]).Value.ToString("yyyy") : string.Empty;
                }
            }

            string fileName = additionalPropertiesLNDSection["FileName"] != null ? additionalPropertiesLNDSection["FileName"].ToString() : string.Empty;

            string сotent = mainInfoLND["Content"] != null ? mainInfoLND["Content"].ToString() : string.Empty;
            string number = additionalPropertiesLNDSection["Number"] != null ? additionalPropertiesLNDSection["Number"].ToString() : string.Empty;
            string approvalNumber = additionalPropertiesLNDSection["ApprovalNumber"] != null ? additionalPropertiesLNDSection["ApprovalNumber"].ToString() : string.Empty;
            string version = additionalPropertiesLNDSection["Version"] != null ? additionalPropertiesLNDSection["Version"].ToString() : string.Empty;

            string typeLNDid = additionalPropertiesLNDSection["TypeLND"] != null ? additionalPropertiesLNDSection["TypeLND"].ToString() : string.Empty;
            string typeLNDstr = GetNameOfRefBaseUniversal(typeLNDid);
           
            DateTime? commissioningDate = (DateTime?)additionalPropertiesLNDSection["CommissioningDate"];
            string stageName = string.Empty;
            string stateName = string.Empty;

            switch (statesSprState["DefaultName"].ToString())
            {
                case "Drafting":
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
            string classificationStr = GetNameOfRefBaseUniversal(classificationId);

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
            List<RowData> rowsAct = new List<RowData>();
            
            if (hystoryLNDSection.Count > 0)
            {
                RowData rowVV = hystoryLNDSection.FirstOrDefault(r => GetNameOfRefBaseUniversal(r["Type"].ToString())== "РД Общества о вводе");
                if (rowVV != null)
                {
                    Guid rdOvv = new Guid(rowVV["RDId"].ToString());
                    if (rdOvv != Guid.Empty)
                        GetDataFromRDCardData(rdOvv, out nameRDvv, out namberRDvv, out fileNameRDvv, out regDateRDvv, out typeRDvv);
                }
                
                RowData rowOtm = hystoryLNDSection.FirstOrDefault(r => GetNameOfRefBaseUniversal(r["Type"].ToString()) == "РД Общества об отмене");
                if (rowOtm != null)
                {
                    Guid rdOtm = new Guid(rowOtm["RDId"].ToString());
                    if (rdOtm != Guid.Empty)
                        GetDataFromRDCardData(rdOtm, out nameRDotm, out namberRDotm, out fileNameRDotm, out regDateRDotm, out typeRDotm);

                }

                rowsAct = (hystoryLNDSection.Where(r => GetNameOfRefBaseUniversal(r["Type"].ToString()) == "РД Общества об актуализации").OrderBy(cr => int.Parse(cr["Number"].ToString()))).ToList<RowData>();
            }

            exelDoc.SetCellValue("ПАО \"Самаранефтехимроект\"", index, 1, color);// Юридическое лицо
            exelDoc.SetCellValue(owner, index, 2, color);//Владелец
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
                        RowData row = hystoryPGSection.FirstOrDefault(r => int.Parse(r["ChangeNumber"].ToString()) == numberIzm);
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
                    Guid RDact = Guid.Empty; //new Guid();//Context.GetObject<Document>(new Guid(rowsAct[i]["RDId"].ToString()));
                    if (Guid.TryParse(rowsAct[i]["RDId"].ToString(),out RDact))
                    {
                        GetDataFromRDCardData(RDact, out nameRDact, out numberRDact, out fileNameRDact, out regDateRDact, out typeRDact);
                    }
                    exelDoc.SetCellValue(nameRDact, index, 19, colorIzm); // Наименование РД
                    exelDoc.SetCellValue(fileNameRDact, index, 20, colorIzm); // Имя файла РД
                    exelDoc.SetCellValue(numberRDact, index, 21, colorIzm); // Номер РД
                    exelDoc.SetCellValue(regDateRDact != null ? regDateRDvv.Value.ToString("d") : string.Empty, index, 22, colorIzm); // Дата РД
                    exelDoc.SetCellValue(typeRDact, index, 23, colorIzm); // Номер РД

                }
            }
            foreach (RowData row in applicationLNDSection)
            {

                if (row != null)
                {
                    index++;// следующая строка после изменений

                    string namePril = row["Name"] != null ? row["Name"].ToString() : string.Empty;
                    string fileNamePril = row["FileName"] != null ? row["FileName"].ToString() : string.Empty;

                    exelDoc.SetCellValue("ПАО \"Самаранефтехимроект\"", index, 1, color, true);// Юридическое лицо
                    exelDoc.SetCellValue(owner, index, 2, color);//Владелец
                    exelDoc.SetCellValue(namePril, index, 5, color, true); // Наименование ЛНД/ Приложения
                    exelDoc.SetCellValue(fileNamePril, index, 6, color, true); // Имя файла ЛНД/приложения/ редакции ЛНД с изм.
                    exelDoc.SetCellValue(name, index, 7, color, true); // Описание
                    exelDoc.SetCellValue(number, index, 8, color, true); // Регистрационный номер ЛНД
                    exelDoc.SetCellValue(approvalNumber, index, 9, color, true); // Номер утверждения ЛНД
                    exelDoc.SetCellValue(version, index, 10, color, true); // Версия ЛНД
                    exelDoc.SetCellValue(typeLNDstr, index, 11, color, true); // Вид ЛНД
                    exelDoc.SetCellValue("Приложение", index, 12, color, true); // Тип документа
                    exelDoc.SetCellValue(commissioningDate != null ? commissioningDate.Value.ToString("d") : string.Empty, index, 13, color, true); // дата
                    exelDoc.SetCellValue(stageName, index, 14, color, true); // Стадия ЛНД
                    exelDoc.SetCellValue(string.Empty, index, 15, color, true); // Срок действия ЛНД
                    exelDoc.SetCellValue(stateName, index, 16, color, true); // Статус ЛНД 
                    exelDoc.SetCellValue(classificationStr, index, 17, color, true); // Бизнес-процесс 1-го и 2-го уровней
                    exelDoc.SetCellValue(string.Empty, index, 18, color, true); // Комментарии

                    //ввод
                    exelDoc.SetCellValue(nameRDvv, index, 19, color, true); // Наименование РД
                    exelDoc.SetCellValue(fileNameRDvv, index, 20, color, true); // Имя файла РД
                    exelDoc.SetCellValue(namberRDvv, index, 21, color, true); // Номер РД
                    exelDoc.SetCellValue(regDateRDvv != null ? regDateRDvv.Value.ToString("d") : string.Empty, index, 22, color, true); // Дата РД
                    exelDoc.SetCellValue(typeRDvv, index, 23, color, true); // Номер РД

                    //отмена
                    exelDoc.SetCellValue(nameRDotm, index, 29, color, true); // Наименование РД
                    exelDoc.SetCellValue(fileNameRDotm, index, 30, color, true); // Имя файла РД
                    exelDoc.SetCellValue(namberRDotm, index, 31, color, true); // Номер РД
                    exelDoc.SetCellValue(regDateRDotm != null ? regDateRDvv.Value.ToString("d") : string.Empty, index, 32, color, true); // Дата РД
                    exelDoc.SetCellValue(typeRDotm, index, 33, color, true); // Номер РД

                }
            }

            return index;

        }
        public string GetFullPathToCopyTemplate(string pathToTemplate)
        {
            string fullPathToTemplate = string.Format("{0}{1}", Directory.GetCurrentDirectory(), pathToTemplate);
            string fileName = "Реестр ЛНД";
            string fullPathDest = string.Format("{0}{1}.xlsx", Path.GetTempPath(), fileName);
            if (File.Exists(fullPathToTemplate))
            {
                while (File.Exists(fullPathDest))
                {
                    try
                    {
                        File.Delete(fullPathDest);
                    }
                    catch
                    {
                        fileName += "1";
                        fullPathDest = string.Format("{0}{1}.xlsx", Path.GetTempPath(), fileName);
                    }
                }
                  
                
                File.Copy(fullPathToTemplate, fullPathDest);

            }
          


            return fullPathDest;
        }
        /// <summary>
        /// Возврящает полное название подразделения из справочника сотрудников
        /// </summary>
        /// <param name="UnitIDstr">Идентификатор подразделиения справочника сотрудников</param>
        /// <returns>Полное название подразделения справочника сотрдуников</returns>
        public string GetFullNameOfUnit(string UnitIDstr)
        {
            string fullName = string.Empty;
            Guid UnitID = Guid.Empty;
            if (Guid.TryParse(UnitIDstr,out UnitID))
            {
                RowData unit = SectionUnitsFromRefStaff.GetAllRows().FirstOrDefault(r => r.Id == UnitID);
                if (unit != null)
                    fullName = unit["FullName"] != null ? unit["FullName"].ToString() : string.Empty;
            }

            return fullName;
        }
        /// <summary>
        /// Метод получает Наименование строки Конструктора справочников
        /// </summary>
        /// <param name="RowIdStr">Идентификатор строки Конструктора справочников</param>
        /// <returns>Наименование строки Конструктора справочников</returns>
        public string GetNameOfRefBaseUniversal(string RowIdStr)
        {
            string name = string.Empty;
            Guid RowId = Guid.Empty;
            if (Guid.TryParse(RowIdStr, out RowId))
            {
                SectionData section = RefBaseUniversalCD.Sections[new Guid("1B1A44FB-1FB1-4876-83AA-95AD38907E24")];
                RowData row = section.GetAllRows().FirstOrDefault(r => r.Id == RowId);
                if (row != null)
                    name = row["Name"] != null ? row["Name"].ToString() : string.Empty;
            }
            return name;
        }
        /*private void GetDataFromRD(Document rd, out string name, out string number, out string fileName, out DateTime? dateRD, out string typeRD)
        {
            BaseCardSectionRow for_LNDSection = rd.GetSectionRow("For_LND");
            name = rd.MainInfo.Name;
            number = rd.Numbers[0].Number;
            dateRD = (DateTime?)rd.MainInfo["RegDate"];
            fileName = for_LNDSection != null ? (for_LNDSection["FileNameRD"] != null ? for_LNDSection["FileNameRD"].ToString() : string.Empty) : string.Empty;
            typeRD = rd.SystemInfo.CardKind.Name;
        }*/
        /// <summary>
        /// Возвращяет набор данных из РД для заполнения реестра
        /// </summary>
        /// <param name="rd">Guid РД</param>
        /// <param name="name">Наименование РД</param>
        /// <param name="number">Номер РД</param>
        /// <param name="fileName">Имя файла РД</param>
        /// <param name="dateRD">Дата РД</param>
        /// <param name="typeRD">Тип Рд</param>
        private void GetDataFromRDCardData(Guid rd, out string name, out string number, out string fileName, out DateTime? dateRD, out string typeRD)
        {
            CardData rdCD = Session.CardManager.GetCardData(rd);

            RowData for_LNDSection = rdCD.Sections[rdCD.Type.Sections["For_LND"].Id].FirstRow;
            RowData MainInfoSection = rdCD.Sections[rdCD.Type.Sections["MainInfo"].Id].FirstRow;
            RowData NumbersSection = rdCD.Sections[rdCD.Type.Sections["Numbers"].Id].FirstRow;
            RowData SystemSection = rdCD.Sections[rdCD.Type.Sections["System"].Id].FirstRow;

            name = MainInfoSection["Name"].ToString();
            number = NumbersSection["Number"].ToString();
            dateRD = (DateTime?)MainInfoSection["RegDate"];
            fileName = for_LNDSection != null ? (for_LNDSection["FileNameRD"] != null ? for_LNDSection["FileNameRD"].ToString() : string.Empty) : string.Empty;

            Guid typeID = Guid.Parse(SystemSection["Kind"].ToString());
            RowData kinde = RefKinds.Sections[new Guid("C7BA000C-6203-4D7F-8C6B-5CB6F1E6F851")].GetRow(typeID);

            typeRD = kinde["Name"].ToString();
        }
        
        public virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        #endregion Methods
    }

  

}
