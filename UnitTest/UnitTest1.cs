using Docsvision.DocumentsManagement;
using DocsVision.BackOffice.WinForms;
using DocsVision.Platform.ObjectManager;
using DocsVision.Platform.ObjectManager.SearchModel;
using DocsVision.Platform.ObjectManager.SystemCards;
using DocsVision.Platform.ObjectModel;
using DocsVision.Platform.ObjectModel.Mapping;
using DocsVision.Platform.ObjectModel.Persistence;
using DocsVision.Platform.ObjectModel.Search;
using DocsVision.Platform.Security.AccessControl;
using DocsVision.Platform.SystemCards.ObjectModel.Mapping;
using DocsVision.Platform.SystemCards.ObjectModel.Services;
using DocsVision.Workflow;
using DocsVisionContext;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using ExportLNDreestr;
using System.Windows;
using ExportLNDreestr.AdditionalClasses;
using Newtonsoft.Json;

namespace UnitTest
{
    [TestClass]
    public class TestModuls
    {
        private DocsVisionContext.DocsVisionContext dvContext = null;
        private ObjectContext Context = null;
        private UserSession Session = null;
        private ViewModel VievModelTst = null;
        private const string ServiceUrl = "http://sam-dvsv-01/docsvision/StorageServer/StorageServerService.asmx";
        private const string BaseName = "docsvision";
        private const string UserName = @"rosneft\dv_services";
        private const string Password = "1pEI5LVd";
        /// <summary>
        /// Инициализация тестов
        /// </summary>
        [TestInitialize]
        public void TestInicialize()
        {

            dvContext = DocsVisionContextFactory.CreateContext(ServiceUrl, BaseName, UserName, Password);
            Context = dvContext.CurrentContext;
            Session = dvContext.CurrentSession;
            VievModelTst = new ViewModel(Session);
        }
        /// <summary>
        /// Закрытие сессии и контекста
        /// </summary>
        [TestCleanup]
        public void TestCleanup()
        {
            Context.Dispose();
            Session.Close();
        }

        [TestMethod]
        public void TestMethod1()
        {
            ViewModel test = new ViewModel(Session);
            //string name = test.GetNameOfRefBaseUniversal("CC4889D8-0E84-4DA1-9E69-2CB19DE6F217");

           // Assert.IsNotNull(name);
        }
        [TestMethod]
        public void TestGetFullNameOfUnit()
        {
            ViewModel test = new ViewModel(Session);
           
            //string name = test.GetFullNameOfUnit("5AF2E7CE-B458-46BD-8D9E-965E754CF94B");

           // Assert.IsNotNull(name, "Значение присуствует");

        }

        [TestMethod]
        public void TryDeserialize()
        {
                      
            using (StreamReader stream = new StreamReader(Directory.GetCurrentDirectory() + @"\Sourse\ConnectionSettings.json"))
            {
                string str = stream.ReadToEnd();
                ConnectionSettings connect = JsonConvert.DeserializeObject<ConnectionSettings>(str);


            }
        }

        [TestMethod]
        public void TrySerialization()
        {
            ConnectionSettings set = new ConnectionSettings
            {
                password = "1",
                username = "polizuk",
                servername = "servak"
            };

            using (StreamWriter stream = new StreamWriter(Directory.GetCurrentDirectory() + @"\Sourse\ConnectionSettings1646.json",false))
            {
                //stream.
                string ser = JsonConvert.SerializeObject(set);
                stream.WriteLine(ser);
                stream.Flush();
                stream.Close();
               // MessageBox.Show(connect.username);

            }
        }
        [TestMethod]
        public void TestGetFullPathToCopyTemplate()
        {
            ViewModel test = new ViewModel(Session);
            //string filePath = test.GetFullPathToCopyTemplate(Directory.GetCurrentDirectory() + @"\Sourse\РЛК.xlsx");
            //Assert.IsTrue(File.Exists(filePath));
        }
        [TestMethod]
        public void Test()
        {
            //string NameRD = VievModelTst.GetAltDecriptionRDCardData(Guid.Parse("2991C5CC-9FAD-E711-80D3-00155D2C0A31"));
        }
    }
}
