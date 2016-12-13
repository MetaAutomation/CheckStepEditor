////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
//  MetaAutomation (C) 2016 by Matt Griscom.
//
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using System.Text;

namespace CheckStepEditor
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class AddCheckStepCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("5fa8d07c-1c90-4ffc-8629-240142098f29");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="AddCheckStepCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private AddCheckStepCommand(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new OleMenuCommand(this.MenuItemCallback, menuCommandID);
                menuItem.BeforeQueryStatus += MenuItem_BeforeQueryStatus;
                commandService.AddCommand(menuItem);
            }
        }

        private DTE2 _dte = null;
        DTE2 GetDTE()
        {
            if (_dte == null)
            {
                _dte = Package.GetGlobalService(typeof(SDTE)) as DTE2;
            }
            return _dte;
        }

        private void MenuItem_BeforeQueryStatus(object sender, EventArgs e)
        {
            OleMenuCommand addCheckStepCommand = (OleMenuCommand)sender;

            // hidden by default
            addCheckStepCommand.Visible = false;
            addCheckStepCommand.Enabled = false;
            Document activeDoc = GetDTE().ActiveDocument;
            if (activeDoc != null && activeDoc.ProjectItem != null && activeDoc.ProjectItem.ContainingProject != null)
            {
                string lang = activeDoc.Language;
                if (activeDoc.Language.Equals("CSharp"))
                {
                    // show command if active document is a csharp file.
                    addCheckStepCommand.Visible = true;
                    addCheckStepCommand.Enabled = true;
                }
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static AddCheckStepCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private System.IServiceProvider ServiceProvider
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
        public static void Initialize(Package package)
        {
            Instance = new AddCheckStepCommand(package);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            try
            {
                EnvDTE._DTE theDTE = DteResources.Instance.GetDteInstance();
                EnvDTE.TextDocument activeTextDocument = (EnvDTE.TextDocument)theDTE.ActiveDocument.Object("TextDocument");
                TextSelection selection = activeTextDocument.Selection;

                this.SubstituteStringFromSelection_AddCheckStep(activeTextDocument);
            }
            catch (NullReferenceException)
            {
            }
            catch (Exception ex)
            {
                VsShellUtilities.ShowMessageBox(this.ServiceProvider, string.Format("Exception type='{0}', Message='{1}'", ex.GetType().ToString(), ex.Message), "Exception",
                    OLEMSGICON.OLEMSGICON_INFO,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
            }
        }

        /// <summary>
        /// Put the selected lines of code inside a check step.
        /// This method does all modification of the text in the editor, with one string substitution
        /// </summary>
        /// <param name="document"></param>
        /// <param name="selection"></param>
        private void SubstituteStringFromSelection_AddCheckStep(TextDocument document)
        {
            EditPoint earlierPoint = null, laterPoint = null;
            EditingTools.Instance.GetEditPointsForLinesToCheck(document, out earlierPoint, out laterPoint);
            string indentSpaces = EditingTools.Instance.GetIndentSpaces(document);
            string linesNoTabs = earlierPoint.GetText(laterPoint).Replace("\t", indentSpaces);
            string[] lines = linesNoTabs.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
            StringBuilder resultingText = new StringBuilder();
            bool firstLine = true;
            string originalIndent = string.Empty;

            foreach (string line in lines)
            {
                if (!firstLine)
                {
                    resultingText.Append(Environment.NewLine);
                }

                firstLine = false;

                if (!String.IsNullOrWhiteSpace(line))
                {
                    string totalIndentSpaces = "";
                    int tabbedSpacesCounter = 0;

                    // Found how many tabbed spaces precede the line text
                    while ((line.Length >= ((tabbedSpacesCounter + 1) * document.TabSize)) && (line.Substring(tabbedSpacesCounter * document.TabSize, document.TabSize) == indentSpaces))
                    {
                        totalIndentSpaces += indentSpaces;
                        tabbedSpacesCounter++;
                    }

                    // find the minimum tab count, but skipping any zero-tab line
                    if (originalIndent.Length == 0)
                    {
                        originalIndent = totalIndentSpaces;
                    }
                    else
                    {
                        if (totalIndentSpaces.Length < originalIndent.Length)
                        {
                            originalIndent = totalIndentSpaces;
                        }
                    }

                    // Start the resulting line with a tab
                    resultingText.Append(indentSpaces);
                }

                // Add the rest of the text
                resultingText.Append(line);
            }

            StringBuilder stringToInsert = new StringBuilder();

            // strings that come before user-selected lines
            foreach (string beforeSelectedCode in StringResources.Instance.StringsBeforeSelectedCode)
            {
                stringToInsert.Append(originalIndent);
                stringToInsert.Append(beforeSelectedCode);
                stringToInsert.Append(Environment.NewLine);
            }

            stringToInsert.Append(resultingText);

            // strings that come after user-selected lines
            foreach (string afterSelectedCode in StringResources.Instance.StringsAfterSelectedCode)
            {
                stringToInsert.Append(originalIndent);
                stringToInsert.Append(afterSelectedCode);
                stringToInsert.Append(Environment.NewLine);
            }

            // Do replacement in Visual Studio text buffer
            earlierPoint.ReplaceText(laterPoint, stringToInsert.ToString(), (int)vsEPReplaceTextOptions.vsEPReplaceTextAutoformat);
        }
    }
}
