using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using Microsoft.VisualStudio.Core.Imaging;
using Microsoft.VisualStudio.Imaging;
using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Classification;
using VsctCompletion.Completion;

namespace VsctCompletion
{
    internal sealed class IdQuickInfoSource : IAsyncQuickInfoSource
    {
        private static readonly ImageId _icon = KnownMonikers.AbstractCube.ToImageId();

        private readonly ITextBuffer _buffer;
        private readonly IClassifier _classifier;

        public IdQuickInfoSource(ITextBuffer buffer, IClassifier classifier)
        {
            _buffer = buffer;
            _classifier = classifier;
        }

        // This is called on a background thread.
        public async Task<QuickInfoItem> GetQuickInfoItemAsync(IAsyncQuickInfoSession session, CancellationToken cancellationToken)
        {
            SnapshotPoint? triggerPoint = session.GetTriggerPoint(_buffer.CurrentSnapshot);

            if (triggerPoint != null)
            {
                ITextSnapshotLine line = triggerPoint.Value.GetContainingLine();
                IList<ClassificationSpan> spans = _classifier.GetClassificationSpans(line.Extent);
                ClassificationSpan attrValue = spans.FirstOrDefault(s => s.ClassificationType.IsOfType("XML Attribute Value") && s.Span.Contains(triggerPoint.Value.Position));

                if (attrValue != null)
                {
                    ITrackingSpan id = _buffer.CurrentSnapshot.CreateTrackingSpan(attrValue.Span, SpanTrackingMode.EdgeInclusive);

                    var fileName = VsctCompletionSource.GetFileName(attrValue.Span.GetText());

                    // Image exist
                    if (!string.IsNullOrEmpty(fileName))
                    {
                        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

                        var img = new Image
                        {
                            Source = new BitmapImage(new Uri(fileName)),
                            MaxHeight = 500
                        };

                        return new QuickInfoItem(id, img);
                    }
                }
            }

            return null;
        }

        public void Dispose()
        {
            // This provider does not perform any cleanup.
        }
    }
}