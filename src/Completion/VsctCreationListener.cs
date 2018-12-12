using System.ComponentModel.Composition;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Language.Intellisense.AsyncCompletion;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.VisualStudio.Utilities;
using VsctCompletion.ContentTypes;

namespace VsctCompletion.Completion
{
    //[Export(typeof(IVsTextViewCreationListener))]
    [ContentType(VsctContentTypeDefinition.VsctContentType)]
    [TextViewRole(PredefinedTextViewRoles.Interactive)]
    internal sealed class VsctTextViewCreationListener : IVsTextViewCreationListener
    {
        [Import]
        IVsEditorAdaptersFactoryService AdaptersFactory { get; set; }

        [Import]
        IAsyncCompletionBroker CompletionBroker { get; set; }

        public void VsTextViewCreated(IVsTextView textViewAdapter)
        {
            IWpfTextView view = AdaptersFactory.GetWpfTextView(textViewAdapter);

            var completion = new VsctCompletionController(view, CompletionBroker);
            textViewAdapter.AddCommandFilter(completion, out IOleCommandTarget completionNext);
            completion.Next = completionNext;
        }
    }
}
