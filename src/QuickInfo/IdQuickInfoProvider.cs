using System.ComponentModel.Composition;
using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Classification;
using Microsoft.VisualStudio.Utilities;
using VsctCompletion.ContentTypes;

namespace VsctCompletion
{
    [Export(typeof(IAsyncQuickInfoSourceProvider))]
    [Name(nameof(IdQuickInfoSourceProvider))]
    [ContentType(VsctContentTypeDefinition.VsctContentType)]
    [Order]
    internal sealed class IdQuickInfoSourceProvider : IAsyncQuickInfoSourceProvider
    {
        [Import]
        private IClassifierAggregatorService ClassifierService { get; set; }

        public IAsyncQuickInfoSource TryCreateQuickInfoSource(ITextBuffer textBuffer)
        {
            IClassifier classifier = ClassifierService.GetClassifier(textBuffer);
            
            // This ensures only one instance per textbuffer is created
            return textBuffer.Properties.GetOrCreateSingletonProperty(() => new IdQuickInfoSource(textBuffer, classifier));
        }
    }
}