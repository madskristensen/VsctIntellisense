using System;
using System.ComponentModel.Composition;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Language.Intellisense.AsyncCompletion;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Classification;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Operations;
using Microsoft.VisualStudio.TextManager.Interop;
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
            var file = GetFileName(buffer);
            return textView.Properties.GetOrCreateSingletonProperty(
                     () => new VsctCompletionSource(buffer, ClassifierService, NavigatorService, file)
                   );
        }

        public static string GetFileName(ITextBuffer buffer)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!buffer.Properties.TryGetProperty(typeof(IVsTextBuffer), out IVsTextBuffer bufferAdapter))
            {
                return null;
            }

            string ppzsFilename = null;
            var returnCode = -1;

            if (bufferAdapter is IPersistFileFormat persistFileFormat)
            {
                try
                {
                    returnCode = persistFileFormat.GetCurFile(out ppzsFilename, out var pnFormatIndex);
                }
                catch (NotImplementedException)
                {
                    return null;
                }
            }

            if (returnCode != VSConstants.S_OK)
            {
                return null;
            }

            return ppzsFilename;
        }
    }
}
