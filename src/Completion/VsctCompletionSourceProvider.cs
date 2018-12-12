using System.ComponentModel.Composition;
using Microsoft.VisualStudio.Language.Intellisense.AsyncCompletion;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Classification;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Operations;
using Microsoft.VisualStudio.Utilities;
using VsctCompletion.ContentTypes;

namespace VsctCompletion.Completion
{
    [Export(typeof(IAsyncCompletionSourceProvider))]
    [Name("Chemical element dictionary completion provider")]
    [ContentType(VsctContentTypeDefinition.VsctContentType)]
    internal class VsctCompletionSourceProvider : IAsyncCompletionSourceProvider
    {
        [Import]
        private IClassifierAggregatorService ClassifierService { get; set; }

        [Import]
        private ITextStructureNavigatorSelectorService NavigatorService { get; set; }

        public IAsyncCompletionSource GetOrCreate(ITextView textView)
        {
            ITextBuffer buffer = textView.TextBuffer;
            return textView.Properties.GetOrCreateSingletonProperty(
                     () => new VsctCompletionSource(buffer, ClassifierService, NavigatorService)
                   );
        }
    }
}
