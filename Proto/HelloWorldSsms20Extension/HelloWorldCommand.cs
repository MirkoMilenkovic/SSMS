using EnvDTE;
using EnvDTE80;
using Microsoft.SqlServer.Management.Smo.RegSvrEnum;
using Microsoft.SqlServer.Management.UI.VSIntegration;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using Task = System.Threading.Tasks.Task;

namespace HelloWorldSsms20Extension
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class HelloWorldCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("d12ce708-8b98-482d-8c12-f3f85a9ee0bd");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        private DTE2 _dte;

        private bool _logicInitialized = false;

        /// <summary>
        /// Initializes a new instance of the <see cref="HelloWorldCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private HelloWorldCommand(
            AsyncPackage package,
            OleMenuCommandService commandService,
            DTE2 dte)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
            _dte = dte;
        }

        private void InitializeLogic()
        {
            // Ensure UI thread.
            // touching Commands requires the UI thread.
            // this should also ensure thread safety
            ThreadHelper.ThrowIfNotOnUIThread();

#warning Mirko multiple init allowed
/*
            if (_logicInitialized)
            {
                return;
            };
*/

            // perform initialization

            Command command = _dte.Commands.Item("Query.Execute");

            CommandEvents queryExecuteEvent = _dte.Events.get_CommandEvents(
                command.Guid, 
                command.ID);
            queryExecuteEvent.BeforeExecute += this.QueryExecute_BeforeExecute;

            _logicInitialized = true;
        }

        private void QueryExecute_BeforeExecute(
            string Guid, 
            int ID, 
            object CustomIn, 
            object CustomOut, 
            ref bool CancelDefault)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                VsShellUtilities.ShowMessageBox(
                this.package,
                "works",
                "handler",
                OLEMSGICON.OLEMSGICON_INFO,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);

                string queryText = GetQueryText();

                if (string.IsNullOrWhiteSpace(queryText))
                    return;

                // Get Current Connection Information
                UIConnectionInfo connInfo = ServiceCache.ScriptFactory.CurrentlyActiveWndConnectionInfo.UIConnectionInfo;

                return;

                /*
                var queryItem = new QueryItem()
                {
                    Query = queryText,
                    Server = connInfo.ServerName,
                    Username = connInfo.UserName,
                    Database = connInfo.AdvancedOptions["DATABASE"],
                    ExecutionDateUtc = DateTime.UtcNow
                };

                _logger.LogInformation("Enqueued {@quetyItem}", queryItem.Query);

                itemsQueue.Enqueue(queryItem);

                Task.Delay(1000).ContinueWith((t) => this.SavePendingItems());
                */
            }
            catch (Exception ex)
            {
                // _logger.LogError("Error on BeforeExecute tracking", ex);
            }
        }

        private string GetQueryText()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            Document document = _dte.ActiveDocument;
            if (document == null)
                return null;

            var textDocument = (TextDocument)document.Object("TextDocument");
            string queryText = textDocument.Selection.Text;

            if (string.IsNullOrEmpty(queryText))
            {
                EditPoint startPoint = textDocument.StartPoint.CreateEditPoint();
                queryText = startPoint.GetText(textDocument.EndPoint);
            }

            return queryText;
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static HelloWorldCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(
            AsyncPackage package,
            DTE2 dte)
        {
            // Switch to the main thread - the call to AddCommand in HelloWorldCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new HelloWorldCommand(
                package,
                commandService,
                dte);

            Instance.InitializeLogic();
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            InitializeLogic();

            string message = string.Format(CultureInfo.CurrentCulture, "Inside {0}.MenuItemCallback()", this.GetType().FullName);
            string title = "HelloWorldCommand";

            string fileFullName = null;
            if (_dte.ActiveDocument != null)
            {
                fileFullName = _dte.ActiveDocument.FullName;
            }

            Microsoft.SqlServer.Management.Smo.RegSvrEnum.UIConnectionInfo connInfo = ServiceCache.ScriptFactory.CurrentlyActiveWndConnectionInfo.UIConnectionInfo;

            // Show a message box to prove we were here
            VsShellUtilities.ShowMessageBox(
                this.package,
                fileFullName,
                title,
                OLEMSGICON.OLEMSGICON_INFO,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }
    }
}
