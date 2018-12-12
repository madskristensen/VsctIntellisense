using System;
using System.Threading;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Language.Intellisense.AsyncCompletion;
using Microsoft.VisualStudio.Language.Intellisense.AsyncCompletion.Data;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;

namespace VsctCompletion.Completion
{
    internal sealed class VsctCompletionController : IOleCommandTarget
    {
        public VsctCompletionController(IWpfTextView textView, IAsyncCompletionBroker broker)
        {
            TextView = textView;
            Broker = broker;
        }

        public IWpfTextView TextView { get; private set; }
        public IAsyncCompletionBroker Broker { get; private set; }
        public IOleCommandTarget Next { get; set; }

        public int Exec(ref Guid pguidCmdGroup, uint nCmdID, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

            if (pguidCmdGroup == VSConstants.VSStd2K)
            {
                switch ((VSConstants.VSStd2KCmdID)nCmdID)
                {
                    case VSConstants.VSStd2KCmdID.COMPLETEWORD:
                    case VSConstants.VSStd2KCmdID.SHOWMEMBERLIST:
                        StartSession();
                        return VSConstants.S_OK;
                }
            }

            return Next.Exec(pguidCmdGroup, nCmdID, nCmdID, pvaIn, pvaOut);
        }

        bool StartSession()
        {
            SnapshotPoint caret = TextView.Caret.Position.BufferPosition;
            var trigger = new CompletionTrigger(CompletionTriggerReason.Invoke, TextView.TextSnapshot);
            IAsyncCompletionSession session = Broker.TriggerCompletion(TextView, trigger, caret, CancellationToken.None);

            if (session is IAsyncCompletionSessionOperations sessionInternal)
            {
                sessionInternal.OpenOrUpdate(trigger, TextView.Caret.Position.BufferPosition, CancellationToken.None);
                return true;
            }
            
            return false;
        }

        public int QueryStatus(ref Guid pguidCmdGroup, uint cCmds, OLECMD[] prgCmds, IntPtr pCmdText)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

            if (pguidCmdGroup == VSConstants.VSStd2K)
            {
                switch ((VSConstants.VSStd2KCmdID)prgCmds[0].cmdID)
                {
                    case VSConstants.VSStd2KCmdID.COMPLETEWORD:
                    case VSConstants.VSStd2KCmdID.SHOWMEMBERLIST:
                        prgCmds[0].cmdf = (uint)OLECMDF.OLECMDF_ENABLED | (uint)OLECMDF.OLECMDF_SUPPORTED;
                        return VSConstants.S_OK;
                }
            }

            return Next.QueryStatus(pguidCmdGroup, cCmds, prgCmds, pCmdText);
        }
    }
}