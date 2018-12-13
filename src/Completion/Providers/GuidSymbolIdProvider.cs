using System;
using System.Collections.Generic;
using System.Xml;
using System.Xml.XPath;
using Microsoft.VisualStudio.Language.Intellisense.AsyncCompletion;
using Microsoft.VisualStudio.Language.Intellisense.AsyncCompletion.Data;
using Microsoft.VisualStudio.Text.Adornments;

namespace VsctCompletion.Completion.Providers
{
    public class GuidSymbolIdProvider : ICompletionProvider
    {
        private readonly List<CompletionItem> _monikerItems;
        private readonly IAsyncCompletionSource _source;
        private readonly ImageElement _icon;

        public GuidSymbolIdProvider(List<CompletionItem> monikerItems, IAsyncCompletionSource source, ImageElement icon)
        {
            _monikerItems = monikerItems;
            _source = source;
            _icon = icon;
        }

        public IEnumerable<CompletionItem> GetCompletions(XmlDocument doc, XPathNavigator navigator, Func<string, CompletionItem> CreateItem)
        {
            var list = new List<CompletionItem>();

            string guid = navigator.GetAttribute("guid", "");

            if (guid == "ImageCatalogGuid")
            {
                list.AddRange(_monikerItems);
            }
            else if (guid == "guidSHLMainMenu")
            {
                list.AddRange(KnownIds.GetItems(_source, _icon));
            }
            else
            {
                XmlNodeList ids = doc.SelectNodes("//GuidSymbol[@name='" + guid + "']//IDSymbol");

                foreach (XmlNode symbol in ids)
                {
                    XmlAttribute name = symbol.Attributes["name"];

                    if (name != null)
                    {
                        list.Add(CreateItem(name.Value));
                    }
                }
            }

            return list;
        }
    }
}
