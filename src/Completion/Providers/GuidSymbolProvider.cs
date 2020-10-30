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
                yield return CreateCompletionItem("UICONTEXT_CodeWindow");
                yield return CreateCompletionItem("UICONTEXT_Debugging");
                yield return CreateCompletionItem("UICONTEXT_DesignMode");
                yield return CreateCompletionItem("UICONTEXT_Dragging");
                yield return CreateCompletionItem("UICONTEXT_EmptySolution");
                yield return CreateCompletionItem("UICONTEXT_FullScreenMode");
                yield return CreateCompletionItem("UICONTEXT_NoSolution");
                yield return CreateCompletionItem("UICONTEXT_NotBuildingAndNotDebugging");
                yield return CreateCompletionItem("UICONTEXT_SolutionBuilding");
                yield return CreateCompletionItem("UICONTEXT_SolutionExists");
                yield return CreateCompletionItem("UICONTEXT_SolutionExistsAndNotBuildingAndNotDebugging");
                yield return CreateCompletionItem("UICONTEXT_SolutionHasMultipleProjects");
                yield return CreateCompletionItem("UICONTEXT_SolutionHasSingleProject");
                yield return CreateCompletionItem("UICONTEXT_ToolboxInitialized");
            }
        }
    }
}
