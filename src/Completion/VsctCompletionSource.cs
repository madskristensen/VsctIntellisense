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

        public VsctCompletionSource(ITextBuffer buffer, IClassifierAggregatorService classifier, ITextStructureNavigatorSelectorService navigator)
        {
            _classifier = classifier.GetClassifier(buffer);
            _navigator = navigator.GetTextStructureNavigator(buffer);
            _parser = new VsctParser(this, _classifier);
        }


        public Task<CompletionContext> GetCompletionContextAsync(IAsyncCompletionSession session, CompletionTrigger trigger, SnapshotPoint triggerLocation, SnapshotSpan applicableToSpan, CancellationToken token)
        {
            if (_parser.TryGetCompletionList(triggerLocation, out IEnumerable<CompletionItem> completions))
            {
                var context = new CompletionContext(completions.ToImmutableArray());

                return Task.FromResult(context);
            }

            return Task.FromResult(CompletionContext.Empty);
        }

        public Task<object> GetDescriptionAsync(IAsyncCompletionSession session, CompletionItem item, CancellationToken token)
        {
            if (item.Properties.TryGetProperty("tooltip", out Lazy<object> tooltip))
            {
                return Task.FromResult<object>(tooltip.Value);
            }

            return Task.FromResult<object>(item.DisplayText);
        }

        public CompletionStartData InitializeCompletion(CompletionTrigger trigger, SnapshotPoint triggerLocation, CancellationToken token)
        {
            if (TryFindTokenSpanAtPosition(triggerLocation, token, out SnapshotSpan span))
            {
                return new CompletionStartData(CompletionParticipation.ProvidesItems, span);
            }

            return CompletionStartData.DoesNotParticipateInCompletion;
        }

        private bool TryFindTokenSpanAtPosition(SnapshotPoint triggerLocation, CancellationToken token, out SnapshotSpan span)
        {
            TextExtent extent = _navigator.GetExtentOfWord(triggerLocation - 1);
            span = extent.Span;

            if (token.IsCancellationRequested)
            {
                return false;
            }

            IList<ClassificationSpan> spans = _classifier.GetClassificationSpans(extent.Span);

            return spans.Any(s => s.ClassificationType.IsOfType("XML Attribute Value"));
        }
    }
}