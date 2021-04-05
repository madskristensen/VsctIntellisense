using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.XPath;
using Microsoft.VisualStudio.Core.Imaging;
using Microsoft.VisualStudio.Imaging;
using Microsoft.VisualStudio.Language.Intellisense.AsyncCompletion;
using Microsoft.VisualStudio.Language.Intellisense.AsyncCompletion.Data;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Adornments;
using VsctCompletion.Completion.Providers;

namespace VsctCompletion.Completion
{
    internal class VsctParser
    {
        private readonly IAsyncCompletionSource _source;
        private readonly string _directory;
        private static readonly ImageElement _icon = new(KnownMonikers.TypePublic.ToImageId());

        public VsctParser(IAsyncCompletionSource source, string file)
        {
            _source = source;
            _directory = Path.GetDirectoryName(file);
        }

        public bool IsAttributeAllowed(string attributeName)
        {
            var allowed = new[] { "id", "guid", "package", "href", "editor", "context" };

            return allowed.Contains(attributeName, StringComparer.OrdinalIgnoreCase);
        }

        public bool TryGetCompletionList(SnapshotPoint triggerLocation, string attrName, out IEnumerable<CompletionItem> completions)
        {
            completions = null;

            SnapshotSpan extent = triggerLocation.GetContainingLine().Extent;
            var line = extent.GetText();

            if (TryGetXmlFragment(line, out XPathNavigator navigator))
            {
                XmlDocument vsct = ReadXmlDocument(triggerLocation.Snapshot.GetText());
                IEnumerable<XmlDocument> allDocs = MergeIncludedVsct(vsct);

                var list = new List<CompletionItem>();

                foreach (XmlDocument doc in allDocs)
                {
                    CompletionItem[] comps = GetCompletions(doc, navigator, attrName).ToArray();

                    foreach (CompletionItem item in comps)
                    {
                        if (!list.Any(c => c.DisplayText == item.DisplayText))
                        {
                            list.Add(item);
                        }
                    }
                }

                completions = list.Distinct();
            }

            return completions != null;
        }

        private readonly Dictionary<XmlDocument, DateTime> _docCache = new();

        private IEnumerable<XmlDocument> MergeIncludedVsct(XmlDocument vsct)
        {
            yield return vsct;

            foreach (XmlNode include in vsct.SelectNodes("//CommandTable/Include"))
            {
                var href = include.Attributes["href"]?.Value;

                if (href != null)
                {
                    var localFile = Path.Combine(_directory, href);

                    if (File.Exists(localFile))
                    {
                        DateTime lwt = File.GetLastWriteTimeUtc(localFile);
                        XmlDocument doc = _docCache.FirstOrDefault(d => d.Value == lwt).Key;

                        if (doc == null)
                        {
                            doc = ReadXmlDocument(File.ReadAllText(localFile));
                            _docCache.Add(doc, File.GetLastWriteTimeUtc(localFile));
                        }

                        yield return doc;
                    }
                }
            }
        }

        private bool TryGetXmlFragment(string fragment, out XPathNavigator doc)
        {
            var settings = new XmlReaderSettings
            {
                ConformanceLevel = ConformanceLevel.Fragment,
                Async = true
            };

            // Fixes open elements
            fragment = fragment
                .Replace("/>", ">")
                .Replace(">", "/>");

            try
            {
                var reader = XmlReader.Create(new StringReader(fragment), settings);
                doc = new XPathDocument(reader).CreateNavigator();
                doc.MoveToFirstChild();

                return true;
            }
            catch (Exception)
            {
                doc = null;
                return false;
            }
        }

        private static XmlDocument ReadXmlDocument(string xml)
        {
            var doc = new XmlDocument();

            try
            {
                doc.LoadXml(Regex.Replace(xml, " xmlns(:[^\"]+)?=\"([^\"]+)\"", string.Empty));
                return doc;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private IEnumerable<CompletionItem> GetCompletions(XmlDocument doc, XPathNavigator navigator, string attribute)
        {
            IEnumerable<ICompletionProvider> providers = GetProviders(attribute);
            var list = new List<CompletionItem>();

            foreach (ICompletionProvider provider in providers)
            {
                list.AddRange(provider.GetCompletions(doc, navigator, CreateCompletionItem));
            }

            return list;
        }

        private IEnumerable<ICompletionProvider> GetProviders(string attributeName)
        {
            switch (attributeName.ToLowerInvariant())
            {
                case "guid":
                case "package":
                case "context":
                    return new ICompletionProvider[] { new GuidSymbolProvider() };
                case "id":
                    return new ICompletionProvider[] { new GuidSymbolIdProvider(_source, _icon) };
                case "editor":
                    return new ICompletionProvider[] { new GuidSymbolProvider(), new EditorProvider() };
                case "href":
                    return new ICompletionProvider[] { new HrefProvider() };
            }

            return Enumerable.Empty<ICompletionProvider>();
        }

        private ImmutableArray<CompletionFilter> _filters = ImmutableArray.Create<CompletionFilter>();

        private CompletionItem CreateCompletionItem(string value, string suffix = "")
        {
            return new CompletionItem(value, _source, _icon, _filters, suffix);
        }
    }
}
