using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.IO;
using System.Linq;
using System.Reflection;
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
    internal class VsctParser(IAsyncCompletionSource source, string file)
    {
        private readonly string _directory = Path.GetDirectoryName(file);
        private static readonly ImageElement _icon = new(KnownMonikers.TypePublic.ToImageId());

        // Improved cache: by file path and last write time
        private readonly Dictionary<string, (DateTime LastWrite, XmlDocument Doc)> _docCache = [];

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
                    CompletionItem[] comps = [.. GetCompletions(doc, navigator, attrName)];
                    for (var i = 0; i < comps.Length; i++)
                    {
                        CompletionItem item = comps[i];
                        var exists = false;
                        for (var j = 0; j < list.Count; j++)
                        {
                            if (list[j].DisplayText == item.DisplayText)
                            {
                                exists = true;
                                break;
                            }
                        }
                        if (!exists)
                        {
                            list.Add(item);
                        }
                    }
                }
                completions = list.Distinct();
            }
            return completions != null;
        }

        private IEnumerable<XmlDocument> MergeIncludedVsct(XmlDocument vsct)
        {
            yield return vsct;
            XmlNodeList includes = vsct.SelectNodes("//CommandTable/Include");
            foreach (XmlNode include in includes)
            {
                var href = include.Attributes["href"]?.Value;
                if (href != null)
                {
                    var localFile = Path.Combine(_directory, href);
                    if (!File.Exists(localFile) && href.Equals("VSGlobals.vsct", StringComparison.OrdinalIgnoreCase))
                    {
                        var extDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                        localFile = Path.Combine(extDir, "Resources\\VSGlobals.vsct");
                    }
                    if (File.Exists(localFile))
                    {
                        DateTime lwt = File.GetLastWriteTimeUtc(localFile);
                        if (_docCache.TryGetValue(localFile, out (DateTime LastWrite, XmlDocument Doc) entry) && entry.LastWrite == lwt)
                        {
                            yield return entry.Doc;
                        }
                        else
                        {
                            XmlDocument doc = ReadXmlDocument(File.ReadAllText(localFile));
                            _docCache[localFile] = (lwt, doc);
                            yield return doc;
                        }
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
            // Try to parse as-is, fallback to fixing open elements if needed
            try
            {
                using var reader = XmlReader.Create(new StringReader(fragment), settings);
                doc = new XPathDocument(reader).CreateNavigator();
                _ = doc.MoveToFirstChild();
                return true;
            }
            catch (Exception)
            {
                // Fallback: try to fix open elements
                fragment = fragment.Replace("/>", ">")
                                   .Replace(">", "/>");
                try
                {
                    using var reader = XmlReader.Create(new StringReader(fragment), settings);
                    doc = new XPathDocument(reader).CreateNavigator();
                    _ = doc.MoveToFirstChild();
                    return true;
                }
                catch
                {
                    doc = null;
                    return false;
                }
            }
        }

        // Improved: only remove xmlns if present
        private static XmlDocument ReadXmlDocument(string xml)
        {
            var doc = new XmlDocument();
            try
            {
                if (xml.Contains("xmlns"))
                {
                    xml = Regex.Replace(xml, " xmlns(:[^\"]+)?=\"([^\"]+)\"", string.Empty);
                }
                doc.LoadXml(xml);
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
            return attributeName.ToLowerInvariant() switch
            {
                "guid" or "package" or "context" => [new GuidSymbolProvider()],
                "id" => [new GuidSymbolIdProvider(source, _icon)],
                "editor" => [new GuidSymbolProvider(), new EditorProvider()],
                "href" => [new HrefProvider()],
                _ => [],
            };
        }

        private ImmutableArray<CompletionFilter> _filters = ImmutableArray.Create<CompletionFilter>();

        private CompletionItem CreateCompletionItem(string value, string suffix = "")
        {
            return new CompletionItem(value, source, _icon, _filters, suffix);
        }
    }
}
