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

namespace UnitTest
{
    [TestClass]
    public class TestModuls
    {
        private DocsVisionContext.DocsVisionContext dvContext = null;
        private ObjectContext Context = null;
        private UserSession Session = null;
        
        /// <summary>
        /// Инициализация тестов
        /// </summary>
        [TestInitialize]
        public void TestInicialize()
        {
            dvContext = DocsVisionContextFactory.CreateDefault();
            Context = dvContext.CurrentContext;
            Session = dvContext.CurrentSession;
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
            string name = test.GetNameOfRefBaseUniversal("CC4889D8-0E84-4DA1-9E69-2CB19DE6F217");

            Assert.IsNotNull(name);
        }
        [TestMethod]
        public void TestGetFullNameOfUnit()
        {
            ViewModel test = new ViewModel(Session);

            string name = test.GetFullNameOfUnit("5AF2E7CE-B458-46BD-8D9E-965E754CF94B");

            Assert.IsNotNull(name, "Значение присуствует");

        }
    }
}
