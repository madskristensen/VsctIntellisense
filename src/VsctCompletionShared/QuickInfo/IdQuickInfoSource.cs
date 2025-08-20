using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Classification;
using VsctCompletion.Completion;

namespace VsctCompletion
{
    internal sealed class IdQuickInfoSource(ITextBuffer buffer, IClassifier classifier) : IAsyncQuickInfoSource
    {
        // Cache for loaded images to avoid reloading
        private static readonly ConcurrentDictionary<string, BitmapImage> _imageCache = new();

        // This is called on a background thread.
        public async Task<QuickInfoItem> GetQuickInfoItemAsync(IAsyncQuickInfoSession session, CancellationToken cancellationToken)
        {
            SnapshotPoint? triggerPoint = session.GetTriggerPoint(buffer.CurrentSnapshot);

            if (triggerPoint != null)
            {
                ITextSnapshotLine line = triggerPoint.Value.GetContainingLine();
                // Use a for loop for better performance
                IList<ClassificationSpan> spans = classifier.GetClassificationSpans(line.Extent);
                ClassificationSpan attrValue = null;
                for (var i = 0; i < spans.Count; i++)
                {
                    ClassificationSpan s = spans[i];
                    if (s.ClassificationType.IsOfType("XML Attribute Value") && s.Span.Contains(triggerPoint.Value.Position))
                    {
                        attrValue = s;
                        break;
                    }
                }

                if (attrValue != null)
                {
                    ITrackingSpan id = buffer.CurrentSnapshot.CreateTrackingSpan(attrValue.Span, SpanTrackingMode.EdgeInclusive);

                    var fileName = VsctCompletionSource.GetFileName(attrValue.Span.GetText());

                    // Image exist
                    if (!string.IsNullOrEmpty(fileName))
                    {
                        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

                        // Use cache to avoid reloading images
                        var img = new Image
                        {
                            MaxHeight = 500,
                            Source = _imageCache.GetOrAdd(fileName, fn => new BitmapImage(new Uri(fn)))
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