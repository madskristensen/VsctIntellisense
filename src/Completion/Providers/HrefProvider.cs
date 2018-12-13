using System;
using System.Collections.Generic;
using System.Xml;
using System.Xml.XPath;
using Microsoft.VisualStudio.Language.Intellisense.AsyncCompletion.Data;

namespace VsctCompletion.Completion.Providers
{
    public class HrefProvider : ICompletionProvider
    {
        public IEnumerable<CompletionItem> GetCompletions(XmlDocument doc, XPathNavigator navigator, Func<string, CompletionItem> CreateCompletionItem)
        {
            if (navigator.LocalName == "Extern")
            {
                yield return CreateCompletionItem("stdidcmd.h");
                yield return CreateCompletionItem("vsshlids.h");
            }
            else if (navigator.LocalName == "Include")
            {
                yield return CreateCompletionItem("KnownImageIds.vsct");
            }
        }
    }
}
