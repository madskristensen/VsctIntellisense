using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Classification;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.TextManager.Interop;

namespace VsctCompletion.GoToDifinition
{
    public partial class VsctGoToDefinitionCommandHandler : IOleCommandTarget
    {
        private readonly IOleCommandTarget _nextCommandHandler;
        private readonly ITextView _textView;
        private readonly VsctGoToDefinitionCreationListener _provider;

        private IClassifier _classifier;

        public VsctGoToDefinitionCommandHandler(IVsTextView textViewAdapter, ITextView textView, VsctGoToDefinitionCreationListener provider)
        {
            _textView = textView;
            _provider = provider;

            var hresult = textViewAdapter.AddCommandFilter(this, out _nextCommandHandler);
        }

        public int QueryStatus(ref Guid pguidCmdGroup, uint cCmds, OLECMD[] prgCmds, IntPtr pCmdText)
        {
            if (pguidCmdGroup == VSConstants.GUID_VSStandardCommandSet97)
            {
                if (cCmds == 1)
                {
                    switch (prgCmds[0].cmdID)
                    {
                        case (uint)VSConstants.VSStd97CmdID.GotoDefn: // F12
                            prgCmds[0].cmdf = (uint)OLECMDF.OLECMDF_SUPPORTED;
                            prgCmds[0].cmdf |= (uint)OLECMDF.OLECMDF_ENABLED;
                            return VSConstants.S_OK;
                    }
                }
            }

            ThreadHelper.ThrowIfNotOnUIThread();

            return _nextCommandHandler.QueryStatus(ref pguidCmdGroup, cCmds, prgCmds, pCmdText);
        }

        public int Exec(ref Guid pguidCmdGroup, uint nCmdID, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (IsCommandExecutedSuccessful(pguidCmdGroup, nCmdID))
            {
                return VSConstants.S_OK;
            }

            return _nextCommandHandler.Exec(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
        }

        private bool IsCommandExecutedSuccessful(Guid pguidCmdGroup, uint nCmdID)
        {
            if (VsShellUtilities.IsInAutomationFunction(_provider.ServiceProvider))
            {
                return false;
            }

            if (pguidCmdGroup == VSConstants.GUID_VSStandardCommandSet97 && nCmdID == (uint)VSConstants.VSStd97CmdID.GotoDefn)
            {
                return TryGoToDefinition();
            }

            return false;
        }

        public static readonly XNamespace NamespaceVsctCommandTable = "http://schemas.microsoft.com/VisualStudio/2005-10-18/CommandTable";

        public const string nameCommandTable = "CommandTable";

        public static readonly XName NodeNameCommandTable = NamespaceVsctCommandTable + nameCommandTable;

        private bool TryGoToDefinition()
        {
            ITextBuffer buffer = _textView.TextBuffer;

            _classifier = _provider.ClassifierAggregatorService.GetClassifier(buffer);

            ITextSnapshot snapshot = buffer.CurrentSnapshot;

            XElement doc = ReadXmlDocument(snapshot.GetText());

            if (doc == null)
            {
                return false;
            }

            if (doc.Name != NodeNameCommandTable)
            {
                return false;
            }

            SnapshotPoint currentPoint = _textView.Caret.Position.BufferPosition;

            IList<ClassificationSpan> spans = _classifier.GetClassificationSpans(new SnapshotSpan(snapshot, 0, snapshot.Length));

            var firstSpans = spans
                .Where(s => s.Span.Start < currentPoint.Position)
                .OrderByDescending(s => s.Span.Start.Position)
                .ToList();

            ClassificationSpan firstDelimiter = firstSpans.FirstOrDefault(s => s.ClassificationType.IsOfType("XML Attribute Quotes"));

            var lastSpans = spans
                 .Where(s => s.Span.Start >= currentPoint.Position)
                 .OrderBy(s => s.Span.Start.Position)
                 .ToList();

            ClassificationSpan lastDelimiter = lastSpans.FirstOrDefault(s => s.ClassificationType.IsOfType("XML Attribute Quotes"));

            SnapshotSpan? extentTemp = null;

            if (firstDelimiter != null && firstDelimiter.Span.GetText() == "\"\"" && firstDelimiter.Span.Contains(currentPoint))
            {
                extentTemp = new SnapshotSpan(firstDelimiter.Span.Start.Add(1), firstDelimiter.Span.Start.Add(1));
            }
            else if (firstDelimiter != null && lastDelimiter != null && firstDelimiter.Span.GetText() == "\"" && lastDelimiter.Span.GetText() == "\"")
            {
                extentTemp = new SnapshotSpan(firstDelimiter.Span.End, lastDelimiter.Span.Start);
            }

            if (!extentTemp.HasValue)
            {
                return false;
            }

            SnapshotSpan extent = extentTemp.Value;

            var currentValue = extent.GetText();

            XElement currentXmlNode = GetCurrentXmlNode(doc, extent);

            if (currentXmlNode == null)
            {
                return false;
            }

            var containingAttributeSpans = spans
                .Where(s => s.Span.Contains(extent.Start)
                    && s.Span.Contains(extent)
                    && s.ClassificationType.IsOfType("XML Attribute Value"))
                .OrderByDescending(s => s.Span.Start.Position)
                .ToList();

            ClassificationSpan containingAttributeValue = containingAttributeSpans.FirstOrDefault();

            if (containingAttributeValue == null)
            {
                containingAttributeValue = spans
                    .Where(s => s.Span.Contains(extent.Start)
                        && s.Span.Contains(extent)
                        && s.ClassificationType.IsOfType("XML Attribute Quotes")
                        && s.Span.GetText() == "\"\""
                    )
                    .OrderByDescending(s => s.Span.Start.Position)
                    .FirstOrDefault();
            }

            if (containingAttributeValue == null)
            {
                return false;
            }

            ClassificationSpan currentAttr = GetCurrentXmlAttributeName(snapshot, containingAttributeValue, spans);

            if (currentAttr == null)
            {
                return false;
            }

            var currentNodeName = currentXmlNode.Name.LocalName;
            var currentAttributeName = currentAttr.Span.GetText();

            if (TryGoToDefinitionInCommandTable(snapshot, doc, currentXmlNode, currentNodeName, currentAttributeName, currentValue))
            {
                return true;
            }

            return false;
        }

        private ClassificationSpan GetCurrentXmlAttributeName(ITextSnapshot snapshot, ClassificationSpan containingSpan, IList<ClassificationSpan> spans)
        {
            ClassificationSpan currentAttr = spans
                    .Where(s => s.ClassificationType.IsOfType("XML Attribute") && s.Span.Start <= containingSpan.Span.Start)
                    .OrderByDescending(s => s.Span.Start.Position)
                    .FirstOrDefault();

            if (currentAttr != null)
            {
                return currentAttr;
            }

            IList<ClassificationSpan> allSpans = _classifier.GetClassificationSpans(new SnapshotSpan(containingSpan.Span.Snapshot, 0, containingSpan.Span.Snapshot.Length));

            currentAttr = allSpans
                    .Where(s => s.ClassificationType.IsOfType("XML Name"))
                    .OrderByDescending(s => s.Span.Start.Position)
                    .FirstOrDefault();

            if (currentAttr != null)
            {
                return currentAttr;
            }

            return null;
        }

        private XElement GetCurrentXmlNode(XElement doc, SnapshotSpan extent)
        {
            IList<ClassificationSpan> spans = _classifier.GetClassificationSpans(new SnapshotSpan(extent.Snapshot, 0, extent.Snapshot.Length));

            ClassificationSpan elementSpan = spans
                     .Where(s => s.ClassificationType.IsOfType("XML Name") && s.Span.Start <= extent.Start)
                     .OrderByDescending(s => s.Span.Start.Position)
                     .FirstOrDefault();

            if (elementSpan == null)
            {
                return null;
            }

            ITextSnapshotLine line = elementSpan.Span.Start.Subtract(1).GetContainingLine();

            var lineNumber = line.LineNumber + 1;
            var linePosition = elementSpan.Span.Start.Subtract(1).Position - line.Start.Position + 2;

            XElement result = doc.DescendantsAndSelf().FirstOrDefault(e => (e as IXmlLineInfo)?.LineNumber == lineNumber);

            return result;
        }

        private static XElement ReadXmlDocument(string text)
        {
            try
            {
                var doc = XDocument.Parse(text, LoadOptions.SetLineInfo);

                return doc.Root;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private bool TryGoToDefinitionInCommandTable(ITextSnapshot snapshot, XElement doc, XElement currentXmlNode, string currentNodeName, string currentAttributeName, string currentValue)
        {
            XNamespace defaultNamespace = doc.GetDefaultNamespace();

            var namespaceManager = new XmlNamespaceManager(new NameTable());
            namespaceManager.AddNamespace("defaulttablenamespace", defaultNamespace.ToString());

            List<XElement> elements = null;

            if (string.Equals(currentAttributeName, "guid", StringComparison.InvariantCultureIgnoreCase))
            {
                if (!string.IsNullOrEmpty(currentValue))
                {
                    elements = doc.XPathSelectElements("./defaulttablenamespace:Symbols/defaulttablenamespace:GuidSymbol", namespaceManager)
                        .Where(e => e.Attribute("name") != null && string.Equals(e.Attribute("name").Value, currentValue, StringComparison.InvariantCultureIgnoreCase))
                        .ToList();

                    if (elements.Count == 1)
                    {
                        XElement guidXmlElement = elements[0];

                        XAttribute attributeGuidName = guidXmlElement.Attribute("name");

                        if (attributeGuidName != null)
                        {
                            if (TryMoveToElement(snapshot, attributeGuidName))
                            {
                                return true;
                            }
                        }
                        else
                        {
                            if (TryMoveToElement(snapshot, guidXmlElement))
                            {
                                return true;
                            }
                        }
                    }
                }

                elements = doc.XPathSelectElements("./defaulttablenamespace:Symbols", namespaceManager).ToList();

                if (elements.Count == 1)
                {
                    XElement symbolsXmlElement = elements[0];

                    if (TryMoveToElement(snapshot, symbolsXmlElement))
                    {
                        return true;
                    }
                }
            }

            if (string.Equals(currentAttributeName, "id", StringComparison.InvariantCultureIgnoreCase))
            {
                XAttribute guidAttribute = currentXmlNode.Attribute("guid");

                if (guidAttribute != null && !string.IsNullOrEmpty(guidAttribute.Value))
                {
                    elements = doc.XPathSelectElements("./defaulttablenamespace:Symbols/defaulttablenamespace:GuidSymbol", namespaceManager)
                        .Where(e => e.Attribute("name") != null && string.Equals(e.Attribute("name").Value, guidAttribute.Value, StringComparison.InvariantCultureIgnoreCase))
                        .ToList();

                    if (elements.Count == 1)
                    {
                        XElement guidXmlElement = elements[0];

                        if (!string.IsNullOrEmpty(currentValue))
                        {
                            elements = guidXmlElement
                                .Descendants(defaultNamespace + "IDSymbol").Where(e => e.Attribute("name") != null && string.Equals(e.Attribute("name").Value, currentValue, StringComparison.InvariantCultureIgnoreCase))
                                .ToList();

                            if (elements.Count == 1)
                            {
                                XElement idXmlElement = elements[0];

                                XAttribute attributeIdName = idXmlElement.Attribute("name");

                                if (attributeIdName != null)
                                {
                                    if (TryMoveToElement(snapshot, attributeIdName))
                                    {
                                        return true;
                                    }
                                }
                                else
                                {
                                    if (TryMoveToElement(snapshot, idXmlElement))
                                    {
                                        return true;
                                    }
                                }
                            }
                        }

                        XAttribute attributeGuidName = guidXmlElement.Attribute("name");

                        if (attributeGuidName != null)
                        {
                            if (TryMoveToElement(snapshot, attributeGuidName))
                            {
                                return true;
                            }
                        }
                        else
                        {
                            if (TryMoveToElement(snapshot, guidXmlElement))
                            {
                                return true;
                            }
                        }
                    }
                    else
                    {
                        elements = doc.XPathSelectElements("./defaulttablenamespace:Symbols", namespaceManager).ToList();

                        if (elements.Count == 1)
                        {
                            XElement symbolsXmlElement = elements[0];

                            if (TryMoveToElement(snapshot, symbolsXmlElement))
                            {
                                return true;
                            }
                        }
                    }
                }
            }

            return false;
        }

        private bool TryMoveToElement(ITextSnapshot snapshot, IXmlLineInfo xmlElement)
        {
            if (xmlElement == null)
            {
                return false;
            }

            ITextSnapshotLine line = snapshot.GetLineFromLineNumber(xmlElement.LineNumber - 1);

            if (line != null)
            {
                SnapshotPoint point = line.Start;

                _textView.ViewScroller.EnsureSpanVisible(new SnapshotSpan(point, 0), EnsureSpanVisibleOptions.ShowStart | EnsureSpanVisibleOptions.AlwaysCenter);

                point += xmlElement.LinePosition - 1;

                _textView.Selection.Select(new SnapshotSpan(point, 0), false);
                _textView.Selection.IsActive = false;

                _textView.Caret.MoveTo(point);

                _textView.Caret.EnsureVisible();

                return true;
            }

            return false;
        }
    }
}