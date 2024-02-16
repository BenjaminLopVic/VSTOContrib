using System;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Word;
using Document = Microsoft.Office.Interop.Word.Document;
using VstoDocument = Microsoft.Office.Tools.Word.Document;
using Factory = Microsoft.Office.Tools.Factory;


namespace VSTOContrib.Word
{
    /// <summary>
    /// 
    /// </summary>
    public class DocumentWrapper
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="DocumentWrapper"/> class.
        /// </summary>
        /// <param name="document">The document.</param>
        public DocumentWrapper(Document document, Factory vstoFactory = null)
        {
            Document = document;
            if (vstoFactory != null && vstoFactory is ApplicationFactory appFactory)
            {
                VstoDocument = appFactory.GetVstoObject(Document);
                VstoDocument.Shutdown += VstoDocument_Shutdown;
                return;
            }
            ((DocumentEvents2_Event)Document).Close += DocumentClose;
        }

        /// <summary>
        /// Occurs when inspector is closed.
        /// </summary>
        public event EventHandler<DocumentClosedEventArgs> Closed;

        /// <summary>
        /// Gets the inspector.
        /// </summary>
        /// <value>The inspector.</value>
        public Document Document { get; private set; }
        public VstoDocument VstoDocument { get; private set; }

        private void DocumentClose()
        {
            ((DocumentEvents2_Event)Document).Close -= DocumentClose;

            Closed?.Invoke(this, new DocumentClosedEventArgs(Document));

            Document = null;
        }

        private void VstoDocument_Shutdown(object sender, EventArgs e)
        {
            VstoDocument.Shutdown -= VstoDocument_Shutdown;

            Closed?.Invoke(this, new DocumentClosedEventArgs(Document));

            Document = null;
        }

    }
}
