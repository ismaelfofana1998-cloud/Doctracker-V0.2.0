#pragma warning disable 414
namespace Doctracker.AddIn
{
    [Microsoft.VisualStudio.Tools.Applications.Runtime.StartupObjectAttribute(0)]
    [global::System.Security.Permissions.PermissionSetAttribute(
        global::System.Security.Permissions.SecurityAction.Demand, Name = "FullTrust")]
    public sealed partial class ThisAddIn : Microsoft.Office.Tools.AddInBase
    {
        internal Microsoft.Office.Tools.CustomTaskPaneCollection CustomTaskPanes;
        internal Microsoft.Office.Tools.SmartTagCollection VstoSmartTags;
        private global::System.Object missing = global::System.Type.Missing;
        internal Microsoft.Office.Interop.Excel.Application Application;

        public ThisAddIn(
            global::Microsoft.Office.Tools.Excel.ApplicationFactory factory,
            global::System.IServiceProvider serviceProvider)
            : base(factory, serviceProvider, "AddIn", "ThisAddIn")
        {
            Globals.Factory = factory;
        }

        protected override void Initialize()
        {
            base.Initialize();
            Application = GetHostItem<Microsoft.Office.Interop.Excel.Application>(
                typeof(Microsoft.Office.Interop.Excel.Application), "Application");
            Globals.ThisAddIn = this;
            global::System.Windows.Forms.Application.EnableVisualStyles();
            InitializeCachedData();
            InitializeControls();
            InitializeComponents();
            InitializeData();
        }

        protected override void FinishInitialization()
        {
            InternalStartup();
            OnStartup();
        }

        protected override void InitializeDataBindings()
        {
            BeginInitialization();
            BindToData();
            EndInitialization();
        }

        private void InitializeCachedData()
        {
            if (DataHost != null && DataHost.IsCacheInitialized)
            {
                DataHost.FillCachedData(this);
            }
        }

        private void InitializeData() { }
        private void BindToData() { }
        private void InitializeComponents() { }

        private void BeginInitialization()
        {
            BeginInit();
            CustomTaskPanes.BeginInit();
            VstoSmartTags.BeginInit();
        }

        private void EndInitialization()
        {
            VstoSmartTags.EndInit();
            CustomTaskPanes.EndInit();
            EndInit();
        }

        private void InitializeControls()
        {
            CustomTaskPanes = Globals.Factory.CreateCustomTaskPaneCollection(
                null, null, "CustomTaskPanes", "CustomTaskPanes", this);
            VstoSmartTags = Globals.Factory.CreateSmartTagCollection(
                null, null, "VstoSmartTags", "VstoSmartTags", this);
        }

        protected override void OnShutdown()
        {
            VstoSmartTags.Dispose();
            CustomTaskPanes.Dispose();
            base.OnShutdown();
        }
    }

    internal sealed partial class Globals
    {
        private Globals() { }
        private static ThisAddIn thisAddIn;
        private static global::Microsoft.Office.Tools.Excel.ApplicationFactory factory;
        private static ThisRibbonCollection ribbonCollection;

        internal static ThisAddIn ThisAddIn
        {
            get { return thisAddIn; }
            set
            {
                if (thisAddIn != null) throw new global::System.NotSupportedException();
                thisAddIn = value;
            }
        }

        internal static global::Microsoft.Office.Tools.Excel.ApplicationFactory Factory
        {
            get { return factory; }
            set
            {
                if (factory != null) throw new global::System.NotSupportedException();
                factory = value;
            }
        }

        internal static ThisRibbonCollection Ribbons
        {
            get
            {
                if (ribbonCollection == null)
                {
                    ribbonCollection = new ThisRibbonCollection(factory.GetRibbonFactory());
                }
                return ribbonCollection;
            }
        }
    }

    internal sealed partial class ThisRibbonCollection :
        Microsoft.Office.Tools.Ribbon.RibbonCollectionBase
    {
        internal ThisRibbonCollection(global::Microsoft.Office.Tools.Ribbon.RibbonFactory factory)
            : base(factory) { }
    }
}
