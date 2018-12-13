using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.XPath;
using Microsoft.VisualStudio.Core.Imaging;
using Microsoft.VisualStudio.Imaging;
using Microsoft.VisualStudio.Imaging.Interop;
using Microsoft.VisualStudio.Language.Intellisense.AsyncCompletion;
using Microsoft.VisualStudio.Language.Intellisense.AsyncCompletion.Data;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Adornments;
using VsctCompletion.Completion.Providers;
using Tasks = System.Threading.Tasks;

namespace VsctCompletion.Completion
{
    internal class VsctParser
    {
        private readonly IAsyncCompletionSource _source;
        private static readonly ImageElement _icon = new ImageElement(KnownMonikers.TypePublic.ToImageId());
        private static List<CompletionItem> _monikerItems = new List<CompletionItem>();
        private static bool _isInitializing;

        public VsctParser(IAsyncCompletionSource source)
        {
            _source = source;
            InitializeAsync().ConfigureAwait(false);
        }

        public bool IsInitialized { get; private set; }

        public bool IsAttributeAllowed(string attributeName)
        {
            string[] allowed = new[] { "id", "guid", "package", "href", "editor" };

            return allowed.Contains(attributeName, StringComparer.OrdinalIgnoreCase);
        }

        public bool TryGetCompletionList(SnapshotPoint triggerLocation, string attrName, out IEnumerable<CompletionItem> completions)
        {
            completions = null;

            SnapshotSpan extent = triggerLocation.GetContainingLine().Extent;
            string line = extent.GetText();

            if (TryGetXmlFragment(line, out XPathNavigator navigator))
            {
                completions = GetCompletions(ReadXmlDocument(triggerLocation.Snapshot), navigator, attrName).ToArray();
            }

            return completions != null;
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
                    return new ICompletionProvider[] { new GuidSymbolIdProvider(_monikerItems, _source, _icon) };
                case "editor":
                    return new ICompletionProvider[] { new GuidSymbolProvider(), new EditorProvider() };
                case "href":
                    return new ICompletionProvider[] { new HrefProvider() };
            }

            return Enumerable.Empty<ICompletionProvider>();
        }
        
        public async Tasks.Task InitializeAsync()
        {
            if (_isInitializing)
            {
                return;
            }

            _isInitializing = true;

            await ThreadHelper.JoinableTaskFactory.StartOnIdle(() =>
            {
                PropertyInfo[] monikers = typeof(KnownMonikers).GetProperties(BindingFlags.Static | BindingFlags.Public);
                var list = new List<CompletionItem>();

                foreach (PropertyInfo monikerName in monikers)
                {
                    var moniker = (ImageMoniker)monikerName.GetValue(null, null);
                    var icon = new ImageElement(moniker.ToImageId());
                    CompletionItem item = CreateCompletionItem(monikerName.Name);

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

                _monikerItems = list;
                _isInitializing = false;
                IsInitialized = true;
            });
        }

        private CompletionItem CreateCompletionItem(string value)
        {
            return new CompletionItem(value, _source, _icon);
        }
    }
}
