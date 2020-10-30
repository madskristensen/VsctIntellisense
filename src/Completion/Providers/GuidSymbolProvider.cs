using System;
using System.Collections.Generic;
using System.Xml;
using System.Xml.XPath;
using Microsoft.VisualStudio.Language.Intellisense.AsyncCompletion.Data;

namespace VsctCompletion.Completion.Providers
{
    public class GuidSymbolProvider : ICompletionProvider
    {
        public IEnumerable<CompletionItem> GetCompletions(XmlDocument doc, XPathNavigator navigator, Func<string, CompletionItem> CreateCompletionItem)
        {
            XmlNodeList guids = doc.SelectNodes("//GuidSymbol");

            foreach (XmlNode node in guids)
            {
                XmlAttribute name = node.Attributes["name"];

                if (name != null)
                {
                    yield return CreateCompletionItem(name.Value);
                }
            }

            if (navigator.LocalName.Equals("parent", StringComparison.OrdinalIgnoreCase))
            {
                yield return CreateCompletionItem("guidSHLMainMenu");
            }
            else if (navigator.LocalName.Equals("icon", StringComparison.OrdinalIgnoreCase))
            {
                yield return CreateCompletionItem("ImageCatalogGuid");
            }
            else if (navigator.LocalName.Equals("VisibilityItem", StringComparison.OrdinalIgnoreCase))
            {
                yield return CreateCompletionItem("GUID_TextEditorFactory");
                yield return CreateCompletionItem("UICONTEXT_BulkFileOperation");
                yield return CreateCompletionItem("UICONTEXT_CloudDebugging");
                yield return CreateCompletionItem("UICONTEXT_CodeWindow");
                yield return CreateCompletionItem("UICONTEXT_DataSourceWindowAutoVisible");
                yield return CreateCompletionItem("UICONTEXT_DataSourceWindowSupported");
                yield return CreateCompletionItem("UICONTEXT_DataSourceWizardSuppressed");
                yield return CreateCompletionItem("UICONTEXT_Debugging");
                yield return CreateCompletionItem("UICONTEXT_DesignMode");
                yield return CreateCompletionItem("UICONTEXT_Dragging");
                yield return CreateCompletionItem("UICONTEXT_EmptySolution");
                yield return CreateCompletionItem("UICONTEXT_FirstLaunchSetup");
                yield return CreateCompletionItem("UICONTEXT_FullScreenMode");
                yield return CreateCompletionItem("UICONTEXT_FullSolutionLoading");
                yield return CreateCompletionItem("UICONTEXT_HistoricalDebugging");
                yield return CreateCompletionItem("UICONTEXT_NoSolution");
                yield return CreateCompletionItem("UICONTEXT_NotBuildingAndNotDebugging");
                yield return CreateCompletionItem("UICONTEXT_OsWindows8OrHigher");
                yield return CreateCompletionItem("UICONTEXT_ProjectCreating");
                yield return CreateCompletionItem("UICONTEXT_ProjectRetargeting");
                yield return CreateCompletionItem("UICONTEXT_RepositoryOpen");
                yield return CreateCompletionItem("UICONTEXT_SolutionBuilding");
                yield return CreateCompletionItem("UICONTEXT_SolutionClosing");
                yield return CreateCompletionItem("UICONTEXT_SolutionExists");
                yield return CreateCompletionItem("UICONTEXT_SolutionExistsAndFullyLoaded");
                yield return CreateCompletionItem("UICONTEXT_SolutionExistsAndNotBuildingAndNotDebugging");
                yield return CreateCompletionItem("UICONTEXT_SolutionHasMultipleProjects");
                yield return CreateCompletionItem("UICONTEXT_SolutionHasSingleProject");
                yield return CreateCompletionItem("UICONTEXT_SolutionOpening");
                yield return CreateCompletionItem("UICONTEXT_SolutionOrProjectUpgrading");
                yield return CreateCompletionItem("UICONTEXT_SynchronousSolutionOperation");
                yield return CreateCompletionItem("UICONTEXT_ToolboxInitialized");
            }
        }
    }
}
