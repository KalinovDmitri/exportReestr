using System;
using System.Collections.Generic;
using System.ComponentModel.Design;

using DocsVision.BackOffice.ObjectModel.Mapping;
using DocsVision.BackOffice.ObjectModel.Services;

using DocsVision.Platform.Data.Metadata;
using DocsVision.Platform.ObjectManager;
using DocsVision.Platform.ObjectModel;
using DocsVision.Platform.ObjectModel.Mapping;
using DocsVision.Platform.ObjectModel.Persistence;
using DocsVision.Platform.SystemCards.ObjectModel.Mapping;
using DocsVision.Platform.SystemCards.ObjectModel.Services;

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
      
        #endregion

        #region Properties

        public bool IsInitialized => Initialized;

        public UserSession CurrentSession => Session;

        public ObjectContext CurrentContext => Context;

      
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
             
                CreateContext();

                Initialized = true;
            }
        }


        #endregion

        #region Private class methods

        private void CreateSession()
        {
            Session = Manager.CreateSession();
        }

        private void CreateContext()
        {
            ServiceContainer container = new ServiceContainer();
            container.AddService(typeof(UserSession), Session);

            ObjectContext context = new ObjectContext(container);

            var mapperRegistry = context.GetService<IObjectMapperFactoryRegistry>();
            mapperRegistry.RegisterFactory(typeof(SystemCardsMapperFactory));
            mapperRegistry.RegisterFactory(typeof(BackOfficeMapperFactory));
        

            var serviceRegistry = context.GetService<IServiceFactoryRegistry>();
            serviceRegistry.RegisterFactory(typeof(SystemCardsServiceFactory));
            serviceRegistry.RegisterFactory(typeof(BackOfficeServiceFactory));
     

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
        #region Public class methods

        public static DocsVisionContext CreateContext(string serviceUrl, string baseName, string userName, string password)
        {
            return new DocsVisionContext(serviceUrl, baseName, userName, password);
        }
        #endregion
    }

  
}
