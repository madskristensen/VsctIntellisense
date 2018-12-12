using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows;
using System.Xml;
using System.Xml.XPath;
using Microsoft.VisualStudio.Core.Imaging;
using Microsoft.VisualStudio.Imaging;
using Microsoft.VisualStudio.Imaging.Interop;
using Microsoft.VisualStudio.Language.Intellisense.AsyncCompletion;
using Microsoft.VisualStudio.Language.Intellisense.AsyncCompletion.Data;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Adornments;
using Microsoft.VisualStudio.Text.Classification;

namespace VsctCompletion.Completion
{
    internal class VsctParser
    {
        private readonly IAsyncCompletionSource _source;
        private readonly IClassifier _classifier;
        private static readonly ImageElement _icon = new ImageElement(KnownMonikers.TypePublic.ToImageId());

        public VsctParser(IAsyncCompletionSource source, IClassifier classifier)
        {
            _source = source;
            _classifier = classifier;
        }

        public bool TryGetCompletionList(SnapshotPoint triggerLocation, out IEnumerable<CompletionItem> completions)
        {
            completions = null;

            SnapshotSpan extent = triggerLocation.GetContainingLine().Extent;
            string line = extent.GetText();

            if (!TryGetXmlFragment(line, out XPathNavigator doc))
            {
                return false;
            }

            IList<ClassificationSpan> spans = _classifier.GetClassificationSpans(extent);
            ClassificationSpan attrValueSpan = spans.FirstOrDefault(s => s.Span.Start <= triggerLocation && s.Span.End >= triggerLocation && s.ClassificationType.IsOfType("XML Attribute Value"));
            int valueSpanIndex = spans.IndexOf(attrValueSpan);

            if (attrValueSpan == null || valueSpanIndex == 0)
            {
                return false;
            }

            ClassificationSpan attrNameSpan = spans.ElementAt(valueSpanIndex - 3);

            if (!attrNameSpan.ClassificationType.IsOfType("XML Attribute"))
            {
                return false;
            }

            string attrName = attrNameSpan.Span.GetText();

            completions = GetCompletions(ReadXmlDocument(triggerLocation.Snapshot), doc, attrName).ToArray();

            return true;
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

        private static XmlDocument ReadXmlDocument(ITextSnapshot snapshot)
        {
            var doc = new XmlDocument();

            try
            {
                doc.LoadXml(Regex.Replace(snapshot.GetText(), " xmlns(:[^\"]+)?=\"([^\"]+)\"", string.Empty));
                return doc;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private IEnumerable<CompletionItem> GetCompletions(XmlDocument doc, XPathNavigator navigator, string attribute)
        {
            if (attribute == "guid" || attribute == "package" || attribute == "context")
            {
                return GetGuidSymbols(doc, navigator);
            }
            else if (attribute == "id")
            {
                return GetGuidSymbolIds(doc, navigator);
            }
            else if (attribute == "editor")
            {
                return GetEditors(doc, navigator);
            }
            else if (attribute == "href")
            {
                return GetHrefs(navigator);
            }

            return Enumerable.Empty<CompletionItem>();
        }

        private IEnumerable<CompletionItem> GetEditors(XmlDocument doc, XPathNavigator navigator)
        {
            var list = GetGuidSymbols(doc, navigator).ToList();

            list.Add(CreateItem("guidVSStd97"));
            list.Add(CreateItem("guidVSStd2K"));

            return list;
        }

        private IEnumerable<CompletionItem> GetHrefs(XPathNavigator navigator)
        {
            if (navigator.LocalName == "Extern")
            {
                yield return CreateItem("stdidcmd.h");
                yield return CreateItem("vsshlids.h");
            }
            else if (navigator.LocalName == "Include")
            {
                yield return CreateItem("KnownImageIds.vsct");
            }
        }

        private IEnumerable<CompletionItem> GetGuidSymbols(XmlDocument doc, XPathNavigator navigator)
        {
            XmlNodeList guids = doc.SelectNodes("//GuidSymbol");

            foreach (XmlNode node in guids)
            {
                XmlAttribute name = node.Attributes["name"];

                if (name != null)
                {
                    yield return CreateItem(name.Value);
                }
            }

            if (navigator.LocalName.Equals("parent", StringComparison.OrdinalIgnoreCase))
            {
                yield return CreateItem("guidSHLMainMenu");
            }
            else if (navigator.LocalName.Equals("icon", StringComparison.OrdinalIgnoreCase))
            {
                yield return CreateItem("ImageCatalogGuid");
            }
            else if (navigator.LocalName.Equals("ImageAttributes", StringComparison.OrdinalIgnoreCase))
            {
                yield return CreateItem("UICONTEXT_CodeWindow");
                yield return CreateItem("UICONTEXT_Debugging");
                yield return CreateItem("UICONTEXT_DesignMode");
                yield return CreateItem("UICONTEXT_Dragging");
                yield return CreateItem("UICONTEXT_EmptySolution");
                yield return CreateItem("UICONTEXT_FullScreenMode");
                yield return CreateItem("UICONTEXT_NoSolution");
                yield return CreateItem("UICONTEXT_NotBuildingAndNotDebugging");
                yield return CreateItem("UICONTEXT_SolutionBuilding");
                yield return CreateItem("UICONTEXT_SolutionExists");
                yield return CreateItem("UICONTEXT_SolutionExistsAndNotBuildingAndNotDebugging");
                yield return CreateItem("UICONTEXT_SolutionHasMultipleProjects");
                yield return CreateItem("UICONTEXT_SolutionHasSingleProject");
                yield return CreateItem("UICONTEXT_ToolboxInitialized");
            }
        }

        private IEnumerable<CompletionItem> GetGuidSymbolIds(XmlDocument doc, XPathNavigator navigator)
        {
            var list = new List<CompletionItem>();

            string guid = navigator.GetAttribute("guid", "");

            if (guid == "ImageCatalogGuid")
            {
                PropertyInfo[] monikers = typeof(KnownMonikers).GetProperties(BindingFlags.Static | BindingFlags.Public);
                foreach (PropertyInfo monikerName in monikers)
                {
                    var moniker = (ImageMoniker)monikerName.GetValue(null, null);
                    var icon = new ImageElement(moniker.ToImageId());
                    CompletionItem item = CreateItem(monikerName.Name);

                    var tooltip = new Lazy<object>(() =>
                    {
                        return new CrispImage
                        {
                            Source = moniker.ToBitmap(100),
                            Height = 100,
                            Width = 100,
                        };
                    });

                    item.Properties.AddProperty("tooltip", tooltip);

                    list.Add(item);
                }
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

        private CompletionItem CreateItem(string value)
        {
            return new CompletionItem(value, _source, _icon);
        }
    }
}
