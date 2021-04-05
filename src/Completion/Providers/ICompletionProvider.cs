using System;
using System.Collections.Generic;
using System.Xml;
using System.Xml.XPath;
using Microsoft.VisualStudio.Language.Intellisense.AsyncCompletion.Data;

namespace VsctCompletion.Completion.Providers
{
    internal interface ICompletionProvider
    {
        IEnumerable<CompletionItem> GetCompletions(XmlDocument doc, XPathNavigator navigator, Func<string, string, CompletionItem> CreateCompletionItem);
    }
}
