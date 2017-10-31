using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel.Design;

using DocsVision.ApprovalDesigner.ObjectModel.Mapping;
using DocsVision.BackOffice.ObjectModel;
using DocsVision.BackOffice.ObjectModel.Mapping;
using DocsVision.BackOffice.ObjectModel.Services;

using DocsVision.Platform.Data.Metadata;
using DocsVision.Platform.ObjectManager;
using DocsVision.Platform.ObjectModel;
using DocsVision.Platform.ObjectModel.Mapping;
using DocsVision.Platform.ObjectModel.Persistence;
using DocsVision.Platform.SystemCards.ObjectModel.Mapping;
using DocsVision.Platform.SystemCards.ObjectModel.Services;

using DocsVision.Platform.Data.Metadata.CardModel;

using DocsVision.Workflow.Objects;
using DocsVision.Workflow.Runtime;
using System.Globalization;


namespace DocsVisionContext
{

    public class DocsVisionContext
    {
        #region Constants and fields

        private const string NotInitializedMessage = "DocsVisionContext not initialized.";

        private readonly string ServiceUrl;
        private readonly string BaseName;
        private readonly string UserName;
        private readonly string Password;

        private readonly SessionManager Manager;

        private bool Initialized = false;
        private UserSession Session;
        private ObjectContext Context;
        private Library WorkflowLibrary;
        #endregion

        #region Properties

        public bool IsInitialized => Initialized;

        public UserSession CurrentSession => Session;

        public ObjectContext CurrentContext => Context;

        public Library CurrentLibrary => WorkflowLibrary;
        #endregion

        #region Constructors

        private DocsVisionContext()
        {
            Manager = ComFactory.CreateSessionManager();
        }

        public DocsVisionContext(string serviceUrl, string baseName, string userName, string password) : this()
        {
            ServiceUrl = serviceUrl;
            BaseName = baseName;
            UserName = userName;
            Password = password;
            Initialize();
        }
        #endregion

        #region Public class methods

        public void Initialize()
        {
            if (!Initialized)
            {
                Manager.Connect(ServiceUrl, BaseName, UserName, Password);

                CreateSession();
                CreateLibrary();
                CreateContext();

                Initialized = true;
            }
        }

        public TObject GetObject<TObject>(Guid objectID) where TObject : ObjectBase
        {
            if (!Initialized)
            {
                throw new InvalidOperationException(NotInitializedMessage);
            }

            ObjectRef<TObject> objectRef = new ObjectRef<TObject>(objectID);

            var parameters = new Dictionary<string, object>();

            return Context.GetObject(objectRef, parameters) as TObject;
        }

        public TService GetService<TService>()
        {
            if (!Initialized)
            {
                throw new InvalidOperationException(NotInitializedMessage);
            }

            return Context.GetService<TService>();
        }

        public ProcessInfo GetProcessInfo(Guid processID)
        {
            if (!Initialized)
            {
                throw new InvalidOperationException(NotInitializedMessage);
            }

            Process dataSource = WorkflowLibrary.GetProcess(processID);

            return new ProcessInfo(dataSource, WorkflowLibrary, Session);
        }
        #endregion

        #region Private class methods

        private void CreateSession()
        {
            Session = Manager.CreateSession();
        }

        private void CreateLibrary()
        {
            WorkflowLibrary = new Library(Session);
        }

        private void CreateContext()
        {
            ServiceContainer container = new ServiceContainer();
            container.AddService(typeof(UserSession), Session);

            ObjectContext context = new ObjectContext(container);

            var mapperRegistry = context.GetService<IObjectMapperFactoryRegistry>();
            mapperRegistry.RegisterFactory(typeof(SystemCardsMapperFactory));
            mapperRegistry.RegisterFactory(typeof(BackOfficeMapperFactory));
            mapperRegistry.RegisterFactory(typeof(ApprovalDesignerMapperFactory));
            //mapperRegistry.RegisterFactory(typeof(ACSGroup.NormativeDocumentManagement.ObjectModel.Mapping.NormativeDocumentManagementMapperFactory));


            var serviceRegistry = context.GetService<IServiceFactoryRegistry>();
            serviceRegistry.RegisterFactory(typeof(SystemCardsServiceFactory));
            serviceRegistry.RegisterFactory(typeof(BackOfficeServiceFactory));
            serviceRegistry.RegisterFactory(typeof(ApprovalDesignerServiceFactory));
            // serviceRegistry.RegisterFactory(typeof(ACSGroup.NormativeDocumentManagement.ObjectModel.Services.NormativeDocumentManagementServiceFactory));

            IMetadataProvider metaProvider = DocsVisionObjectFactory.CreateMetadataProvider(Session);
            IMetadataManager metaManager = DocsVisionObjectFactory.CreateMetadataManager(metaProvider, Session);
            context.AddService(metaManager);
            context.AddService(metaProvider);

            IPersistentStore store = DocsVisionObjectFactory.CreatePersistentStore(Session, null);
            context.AddService(store);

            Context = context;
        }
        #endregion
    }

    public static class DocsVisionContextFactory
    {
        #region Constants

        private const string ServiceUrl = "http://sam-dvsv-01/docsvision/StorageServer/StorageServerService.asmx";
        private const string BaseName = "docsvision";
        private const string UserName = @"";
        private const string Password = "";

        private const string ServiceUrlTEST = "http://SAM-TEST-VDBSV/DocsVision/StorageServer/StorageServerService.asmx";
        private const string BaseNameTEST = "docsvision_in54";
        private const string UserNameTEST = @"";
        private const string PasswordTEST = "";

        #endregion

        #region Public class methods

        public static DocsVisionContext CreateDefault()
        {
            return new DocsVisionContext(ServiceUrl, BaseName, UserName, Password);

        }
        public static DocsVisionContext CreateTestSRVDefault()
        {
            return new DocsVisionContext(ServiceUrlTEST, BaseNameTEST, UserNameTEST, PasswordTEST);

        }
        public static DocsVisionContext CreateTestSRVContext(string userName, string password)
        {
            return new DocsVisionContext(ServiceUrlTEST, BaseNameTEST, userName, password);

        }
        public static DocsVisionContext CreateContext(string serviceUrl, string baseName, string userName, string password)
        {
            return new DocsVisionContext(serviceUrl, baseName, userName, password);
        }
        #endregion
    }

    public static class BaseCardExtensions
    {
        public static BaseCardSectionRow GetSectionRow(this BaseCard card, string sectionAlias)
        {
            if (card == null)
            {
                throw new ArgumentNullException("card", "Card can't be null.");
            }
            if (string.IsNullOrEmpty(sectionAlias))
            {
                throw new ArgumentNullException("sectionAlias", "Section alias can't be null or empty.");
            }

            CardSection section = card.CardType.GetSectionByAlias(sectionAlias);
            if (section != null)
            {
                IList sectionData = card.GetSection(section.Id);
                if (sectionData.Count == 0)
                {
                    sectionData.Add(new BaseCardSectionRow());
                }
                return sectionData[0] as BaseCardSectionRow;
            }
            return null;
        }
        /// <summary>
        /// Возвращает секцию с указанным псевдонимом в виде коллекции строк этой секции
        /// </summary>
        /// <param name="card">Карточка, для которой запрашивается получение секции</param>
        /// <param name="sectionAlias">Псевдоним запрашиваемой секции</param>
        /// <returns>Коллекция строк, представляющих секцию в случае успешности её получения; иначе null</returns>
        public static IList GetSection(this BaseCard card, string sectionAlias)
        {
            if (card == null)
            {
                throw new ArgumentNullException("card");
            }
            if (string.IsNullOrEmpty(sectionAlias))
            {
                throw new ArgumentOutOfRangeException("sectionAlias");
            }

            CardSection section = card.CardType.GetSectionByAlias(sectionAlias);
            if (section != null)
            {
                return card.GetSection(section.Id);
            }

            return null;
        }
        public static T GetValueFromCard<T>(this BaseCard card,string sectionAlias, string fieldAlias)
        {
            T result=default(T);
            BaseCardSectionRow sectionRow = card.GetSectionRow(sectionAlias);
            if (sectionRow != null)
            {
                object value = sectionRow[fieldAlias];
                if (value is T)
                {
                    result = (T)value;
                }
                else if (value is IConvertible)
                {
                    IConvertible converteble = value as IConvertible;
                    result = (T)converteble.ToType(typeof(T), CultureInfo.InvariantCulture);
                    //result = (T)(object)sectionRow[fieldAlias];
                }
                else if(value!=null)
                {
                    string message = string.Format("Can't convert value of type '{0}' to type '{1}'", value.GetType().Name, typeof(T).Name);
                    throw new InvalidOperationException(message);
                }
            }
            return result;
        }

        public static IList GetSectionRows(this BaseCard card, string sectionAlias)
        {
            if (card == null)
            {
                throw new ArgumentNullException("card", "Card can't be null.");
            }
            if (string.IsNullOrEmpty(sectionAlias))
            {
                throw new ArgumentNullException("sectionAlias", "Section alias can't be null or empty.");
            }

            CardSection section = card.CardType.GetSectionByAlias(sectionAlias);
            if (section != null)
            {
                IList sectionData = card.GetSection(section.Id);
                if (sectionData.Count == 0)
                {
                    sectionData.Add(new BaseCardSectionRow());
                }
                return sectionData;
            }
            return null;
        }
    }
}
