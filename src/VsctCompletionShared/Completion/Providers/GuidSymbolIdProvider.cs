﻿using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using System.Xml;
using System.Xml.XPath;
using Microsoft.VisualStudio.Language.Intellisense.AsyncCompletion;
using Microsoft.VisualStudio.Language.Intellisense.AsyncCompletion.Data;
using Microsoft.VisualStudio.Text.Adornments;

namespace VsctCompletion.Completion.Providers
{
    public class GuidSymbolIdProvider : ICompletionProvider
    {
        private readonly IAsyncCompletionSource _source;
        private readonly ImageElement _icon;
        private static IEnumerable<CompletionItem> _knownIds, _knownMonikers;

        public GuidSymbolIdProvider(IAsyncCompletionSource source, ImageElement icon)
        {
            _source = source;
            _icon = icon;
        }

        public IEnumerable<CompletionItem> GetCompletions(XmlDocument doc, XPathNavigator navigator, Func<string, string, CompletionItem> CreateItem)
        {
            var list = new List<CompletionItem>();

            var guid = navigator.GetAttribute("guid", "");

            if (guid == "ImageCatalogGuid")
            {
                list.AddRange(GetKnownMonikers(_source, _icon));
            }
            else if (guid == "guidSHLMainMenu")
            {
                list.AddRange(GetKnownIds(_source, _icon));
            }
            else
            {
                XmlNodeList ids = doc.SelectNodes("//GuidSymbol[@name='" + guid + "']//IDSymbol");

                foreach (XmlNode symbol in ids)
                {
                    XmlAttribute name = symbol.Attributes["name"];

                    if (name != null)
                    {
                        CompletionItem item;
                        if (guid.StartsWith("VS", StringComparison.Ordinal))
                        {
                            var isMenu = name.Value.Trim('.').Count(c => c == '.') % 2 == 0;
                            var typeName = isMenu ? "<Menu>" : "<Group>";
                            item = CreateItem(name.Value, typeName);
                            item.Properties.AddProperty("IsGlobal", true);
                        }
                        else {
                            item = CreateItem(name.Value, "");
                        }

                        list.Add(item);
                    }
                }
            }

            return list;
        }

        private IEnumerable<CompletionItem> GetKnownMonikers(IAsyncCompletionSource source, ImageElement icon)
        {
            if (_knownMonikers == null)
            {
                var list = new List<CompletionItem>();

                foreach (var name in KnownMonikersList.KnownMonikerNames)
                {
                    var item = new CompletionItem(name, source, icon, ImmutableArray.Create<CompletionFilter>(), "Image");
                    item.Properties.AddProperty("knownmoniker", name);
                    list.Add(item);
                }

                _knownMonikers = list;
            }

            return _knownMonikers;
        }

        private IEnumerable<CompletionItem> GetKnownIds(IAsyncCompletionSource source, ImageElement icon)
        {
            if (_knownIds == null)
            {
                _knownIds = KnownIdList.KnownIds.Select(k => new CompletionItem(k, source, icon, ImmutableArray.Create<CompletionFilter>(), "Group/Menu"));
            }

            return _knownIds;
        }
    }
}
