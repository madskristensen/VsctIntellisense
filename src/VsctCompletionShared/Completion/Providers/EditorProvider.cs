using System;
using System.Collections.Generic;
using System.Xml;
using System.Xml.XPath;
using Microsoft.VisualStudio.Language.Intellisense.AsyncCompletion.Data;

namespace VsctCompletion.Completion.Providers
{
    public class EditorProvider : ICompletionProvider
    {
        private static List<CompletionItem> _items;

        public IEnumerable<CompletionItem> GetCompletions(XmlDocument doc, XPathNavigator navigator, Func<string, string, CompletionItem> CreateCompletionItem)
        {
            if (_items == null)
            {
                _items = new List<CompletionItem>
                {
                    CreateCompletionItem("GUID_TextEditorFactory", "Guid"),
                    CreateCompletionItem("guidVSStd97", "Guid"),
                    CreateCompletionItem("guidVSStd2K", "Guid"),
                };
            }

            return _items;
        }
    }
}
