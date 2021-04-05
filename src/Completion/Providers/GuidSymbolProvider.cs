using System;
using System.Collections.Generic;
using System.Xml;
using System.Xml.XPath;
using Microsoft.VisualStudio.Language.Intellisense.AsyncCompletion.Data;

namespace VsctCompletion.Completion.Providers
{
    public class GuidSymbolProvider : ICompletionProvider
    {
        public IEnumerable<CompletionItem> GetCompletions(XmlDocument doc, XPathNavigator navigator, Func<string, string, CompletionItem> CreateCompletionItem)
        {
            XmlNodeList guids = doc.SelectNodes("//GuidSymbol");

            foreach (XmlNode node in guids)
            {
                XmlAttribute name = node.Attributes["name"];

                if (name != null)
                {
                    // Only show results for VSGlobals in <parent> nodes
                    if (navigator.LocalName == "Parent" || name.Value != "VSGlobals")
                    {
                        yield return CreateCompletionItem(name.Value, "Guid");
                    }
                }
            }

            if (navigator.LocalName.Equals("parent", StringComparison.OrdinalIgnoreCase))
            {
                yield return CreateCompletionItem("guidSHLMainMenu", "Guid");
            }
            else if (navigator.LocalName.Equals("icon", StringComparison.OrdinalIgnoreCase))
            {
                yield return CreateCompletionItem("ImageCatalogGuid", "KnownMonikers");
            }
            else if (navigator.LocalName.Equals("VisibilityItem", StringComparison.OrdinalIgnoreCase))
            {
                yield return CreateCompletionItem("GUID_TextEditorFactory", "Guid");
                yield return CreateCompletionItem("UICONTEXT_BulkFileOperation", "Guid");
                yield return CreateCompletionItem("UICONTEXT_CloudDebugging", "Guid");
                yield return CreateCompletionItem("UICONTEXT_CodeWindow", "Guid");
                yield return CreateCompletionItem("UICONTEXT_DataSourceWindowAutoVisible", "Guid");
                yield return CreateCompletionItem("UICONTEXT_DataSourceWindowSupported", "Guid");
                yield return CreateCompletionItem("UICONTEXT_DataSourceWizardSuppressed", "Guid");
                yield return CreateCompletionItem("UICONTEXT_Debugging", "Guid");
                yield return CreateCompletionItem("UICONTEXT_DesignMode", "Guid");
                yield return CreateCompletionItem("UICONTEXT_Dragging", "Guid");
                yield return CreateCompletionItem("UICONTEXT_EmptySolution", "Guid");
                yield return CreateCompletionItem("UICONTEXT_FirstLaunchSetup", "Guid");
                yield return CreateCompletionItem("UICONTEXT_FullScreenMode", "Guid");
                yield return CreateCompletionItem("UICONTEXT_FullSolutionLoading", "Guid");
                yield return CreateCompletionItem("UICONTEXT_HistoricalDebugging", "Guid");
                yield return CreateCompletionItem("UICONTEXT_NoSolution", "Guid");
                yield return CreateCompletionItem("UICONTEXT_NotBuildingAndNotDebugging", "Guid");
                yield return CreateCompletionItem("UICONTEXT_OsWindows8OrHigher", "Guid");
                yield return CreateCompletionItem("UICONTEXT_ProjectCreating", "Guid");
                yield return CreateCompletionItem("UICONTEXT_ProjectRetargeting", "Guid");
                yield return CreateCompletionItem("UICONTEXT_RepositoryOpen", "Guid");
                yield return CreateCompletionItem("UICONTEXT_SolutionBuilding", "Guid");
                yield return CreateCompletionItem("UICONTEXT_SolutionClosing", "Guid");
                yield return CreateCompletionItem("UICONTEXT_SolutionExists", "Guid");
                yield return CreateCompletionItem("UICONTEXT_SolutionExistsAndFullyLoaded", "Guid");
                yield return CreateCompletionItem("UICONTEXT_SolutionExistsAndNotBuildingAndNotDebugging", "Guid");
                yield return CreateCompletionItem("UICONTEXT_SolutionHasMultipleProjects", "Guid");
                yield return CreateCompletionItem("UICONTEXT_SolutionHasSingleProject", "Guid");
                yield return CreateCompletionItem("UICONTEXT_SolutionOpening", "Guid");
                yield return CreateCompletionItem("UICONTEXT_SolutionOrProjectUpgrading", "Guid");
                yield return CreateCompletionItem("UICONTEXT_SynchronousSolutionOperation", "Guid");
                yield return CreateCompletionItem("UICONTEXT_ToolboxInitialized", "Guid");
            }
        }
    }
}
