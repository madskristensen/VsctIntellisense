using System;
using System.Collections.Generic;
using System.Xml;
using System.Xml.XPath;
using Microsoft.VisualStudio.Language.Intellisense.AsyncCompletion.Data;

namespace VsctCompletion.Completion.Providers
{
    public class HrefProvider : ICompletionProvider
    {
        public IEnumerable<CompletionItem> GetCompletions(XmlDocument doc, XPathNavigator navigator, Func<string, string, CompletionItem> CreateCompletionItem)
        {
            if (navigator.LocalName == "Extern")
            {
                yield return CreateCompletionItem("stdidcmd.h", "Def");
                yield return CreateCompletionItem("vsshlids.h", "Def");
            }
            else if (navigator.LocalName == "Include")
            {
                yield return CreateCompletionItem("KnownImageIds.vsct", "KnownMonikers");
                yield return CreateCompletionItem("VSGlobals.vsct", "Aliases");
            }
        }
    }
}
