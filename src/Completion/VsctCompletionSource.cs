using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Language.Intellisense.AsyncCompletion;
using Microsoft.VisualStudio.Language.Intellisense.AsyncCompletion.Data;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Classification;
using Microsoft.VisualStudio.Text.Operations;

namespace VsctCompletion.Completion
{
    class VsctCompletionSource : IAsyncCompletionSource
    {
        private readonly IClassifier _classifier;
        private readonly ITextStructureNavigator _navigator;
        private readonly VsctParser _parser;
        private string _attributeName;

        public VsctCompletionSource(ITextBuffer buffer, IClassifierAggregatorService classifier, ITextStructureNavigatorSelectorService navigator)
        {
            _classifier = classifier.GetClassifier(buffer);
            _navigator = navigator.GetTextStructureNavigator(buffer);
            _parser = new VsctParser(this);
        }

        public async Task<CompletionContext> GetCompletionContextAsync(IAsyncCompletionSession session, CompletionTrigger trigger, SnapshotPoint triggerLocation, SnapshotSpan applicableToSpan, CancellationToken token)
        {
            await EnsureInitializedAsync();

            if (_parser.TryGetCompletionList(triggerLocation, _attributeName, out IEnumerable<CompletionItem> completions))
            {
                return new CompletionContext(completions.ToImmutableArray());
            }

            return CompletionContext.Empty;
        }

        public Task<object> GetDescriptionAsync(IAsyncCompletionSession session, CompletionItem item, CancellationToken token)
        {
            if (item.Properties.TryGetProperty("tooltip", out Lazy<object> tooltip))
            {
                return Task.FromResult<object>(tooltip.Value);
            }

            return Task.FromResult<object>(null);
        }

        public CompletionStartData InitializeCompletion(CompletionTrigger trigger, SnapshotPoint triggerLocation, CancellationToken token)
        {
            IsCompletionSupported(triggerLocation, out _attributeName, out SnapshotSpan span);

            if (!string.IsNullOrEmpty(_attributeName) && _parser.IsAttributeAllowed(_attributeName))
            {
                return new CompletionStartData(CompletionParticipation.ProvidesItems, span);
            }

            return CompletionStartData.DoesNotParticipateInCompletion;
        }

        private bool IsXmlAttributeValue(SnapshotPoint triggerLocation)
        {
            TextExtent extent = _navigator.GetExtentOfWord(triggerLocation - 1);            
            IList<ClassificationSpan> spans = _classifier.GetClassificationSpans(extent.Span);

            return spans.Any(s => s.ClassificationType.IsOfType("XML Attribute Value"));
        }

        private bool IsCompletionSupported(SnapshotPoint triggerLocation, out string attributeName, out SnapshotSpan applicapleTo)
        {
            applicapleTo = new SnapshotSpan(triggerLocation, 0);
            attributeName = null;

            if (!IsXmlAttributeValue(triggerLocation))
            {
                return false;
            }

            applicapleTo = triggerLocation.GetContainingLine().Extent;
            string line = applicapleTo.GetText();

            IList<ClassificationSpan> spans = _classifier.GetClassificationSpans(applicapleTo);
            ClassificationSpan attrValueSpan = spans.FirstOrDefault(s => s.Span.Start <= triggerLocation && s.Span.End >= triggerLocation && s.ClassificationType.IsOfType("XML Attribute Value"));
            int valueSpanIndex = spans.IndexOf(attrValueSpan);

            if (attrValueSpan == null || valueSpanIndex < 3)
            {
                return false;
            }

            applicapleTo = attrValueSpan.Span;
            ClassificationSpan attrNameSpan = spans.ElementAt(valueSpanIndex - 3);

            if (!attrNameSpan.ClassificationType.IsOfType("XML Attribute"))
            {
                return false;
            }

            attributeName = attrNameSpan.Span.GetText();

            return true;
        }

        private async Task EnsureInitializedAsync()
        {
            if (!_parser.IsInitialized)
            {
                await _parser.InitializeAsync();
            }
        }
    }
}