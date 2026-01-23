using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using Microsoft.VisualStudio.Imaging;
using Microsoft.VisualStudio.Imaging.Interop;
using Microsoft.VisualStudio.Language.Intellisense.AsyncCompletion;
using Microsoft.VisualStudio.Language.Intellisense.AsyncCompletion.Data;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Classification;
using Microsoft.VisualStudio.Text.Operations;
using Task = System.Threading.Tasks.Task;

namespace VsctCompletion.Completion
{
    internal class VsctCompletionSource : IAsyncCompletionSource
    {
        private readonly ITextBuffer _buffer;
        private readonly IClassifier _classifier;
        private readonly ITextStructureNavigator _navigator;
        private readonly VsctParser _parser;
        private string _attributeName;

        public VsctCompletionSource(ITextBuffer buffer, IClassifierAggregatorService classifier, ITextStructureNavigatorSelectorService navigator, string file)
        {
            _buffer = buffer;
            _classifier = classifier.GetClassifier(buffer);
            _navigator = navigator.GetTextStructureNavigator(buffer);

            _parser = new VsctParser(this, file);
        }

        public Task<CompletionContext> GetCompletionContextAsync(IAsyncCompletionSession session, CompletionTrigger trigger, SnapshotPoint triggerLocation, SnapshotSpan applicableToSpan, CancellationToken token)
        {
            if (_parser.TryGetCompletionList(triggerLocation, _attributeName, out IEnumerable<CompletionItem> completions))
            {
                return Task.FromResult(new CompletionContext(completions.ToImmutableArray()));
            }

            return Task.FromResult(CompletionContext.Empty);
        }

        public async Task<object> GetDescriptionAsync(IAsyncCompletionSession session, CompletionItem item, CancellationToken token)
        {
            if (item.Properties.TryGetProperty("knownmoniker", out string name))
            {
                try
                {
                    PropertyInfo property = typeof(KnownMonikers).GetProperty(name, BindingFlags.Static | BindingFlags.Public);
                    var moniker = (ImageMoniker)property.GetValue(null, null);

                    await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

                    var image = new CrispImage
                    {
                        Source = await moniker.ToBitmapSourceAsync(100),
                        Height = 100,
                        Width = 100,
                    };

                    return image;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.Write(ex);
                }
            }
            else if (item.Properties.TryGetProperty("IsGlobal", out bool isGlobal) && isGlobal)
            {
                var img = GetFileName(item.DisplayText);

                if (!string.IsNullOrEmpty(img))
                {
                    await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                    
                    return new Image
                    {
                        Source = new BitmapImage(new Uri(img)),
                        MaxHeight = 720 // VS minimum requirements is 1280x720
                    };
                }
            }

            return null;
        }

        public static string GetFileName(string displayName)
        {
            var vsixDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            var menusDir = Path.Combine(vsixDir, "Resources\\Menus");
            
            var fileName = displayName;
            
            while (!string.IsNullOrEmpty(fileName))
            {
                var file = Path.Combine(menusDir, fileName + ".png");
                if (File.Exists(file))
                {
                    return file;
                }

                var sections = fileName.Split(new[] { '.' }, StringSplitOptions.RemoveEmptyEntries);
                fileName = string.Join(".", sections.Take(sections.Length - 1));
            }

            return null;
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
            // Ensure the trigger location is on the correct buffer's snapshot
            ITextSnapshot currentSnapshot = _buffer.CurrentSnapshot;
            SnapshotPoint translatedPoint = triggerLocation.TranslateTo(currentSnapshot, PointTrackingMode.Positive);

            if (translatedPoint.Position == 0)
            {
                return false;
            }

            TextExtent extent = _navigator.GetExtentOfWord(translatedPoint - 1);
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
            var line = applicapleTo.GetText();

            IList<ClassificationSpan> spans = _classifier.GetClassificationSpans(applicapleTo);
            ClassificationSpan attrValueSpan = spans.FirstOrDefault(s => s.Span.Start <= triggerLocation && s.Span.End >= triggerLocation && s.ClassificationType.IsOfType("XML Attribute Value"));
            var valueSpanIndex = spans.IndexOf(attrValueSpan);

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
    }
}