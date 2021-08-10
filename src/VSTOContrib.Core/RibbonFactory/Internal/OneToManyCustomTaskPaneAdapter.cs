using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Tools;

namespace VSTOContrib.Core.RibbonFactory.Internal
{
    class OneToManyCustomTaskPaneAdapter : ICustomTaskPaneWrapper
    {
        private readonly CustomTaskPane original;
        private readonly List<CustomTaskPane> customTaskPanes;
        private bool disposed;
        bool hasBeenHidden;

        public bool CheckDispose()
        {
            if (this.disposed)
            {
                return true;
            }

            try
            {
                bool dummy = original.Visible;
            }
            catch (ObjectDisposedException)
            {
                this.Dispose();
            }

            return this.disposed;
        }

        private void DoOnPane(Action<CustomTaskPane> onPane)
        {
            if (disposed) return;

            try
            {
                onPane(original);
            }
            catch (ObjectDisposedException)
            {
                Dispose();
            }
        }

        private T DoOnPane<T>(Func<CustomTaskPane, T> onPane, T defaultValue = default(T))
        {
            if (disposed) return defaultValue;

            try
            {
                return onPane(original);
            }
            catch (ObjectDisposedException)
            {
                Dispose();
                return defaultValue;
            }
        }

        public OneToManyCustomTaskPaneAdapter(CustomTaskPane original, object viewContext)
        {
            ViewContext = viewContext;
            this.original = original;
            customTaskPanes = new List<CustomTaskPane>();
            Add(original);
        }

        public bool ViewRegistered(object view)
        {
            if (disposed) return false;
            return customTaskPanes.Any(c => c.Window == view);
        }

        public void Add(CustomTaskPane customTaskPane)
        {
            if (disposed) return;
            //Sync new task pane's properties up
            customTaskPane.Visible = this.DoOnPane((p) => p.Visible, false);

            var dockPosition = this.DoOnPane((p) => p.DockPosition);
            customTaskPane.DockPosition = dockPosition;

            if (dockPosition != Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionTop &&
                dockPosition != Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionBottom)
            {
                customTaskPane.Width = this.DoOnPane((p) => p.Width);
            }
            if (dockPosition != Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionLeft &&
                dockPosition != Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight)
            {
                customTaskPane.Height = this.DoOnPane((p) => p.Height);
            }

            customTaskPanes.Add(customTaskPane);
            customTaskPane.DockPositionChanged += CustomTaskPaneDockPositionChanged;
            customTaskPane.VisibleChanged += CustomTaskPaneVisibleChanged;
        }

        public void Refresh(object view)
        {

        }

        void CustomTaskPaneVisibleChanged(object sender, EventArgs e)
        {
            if (disposed) return;
            var customTaskPane = (CustomTaskPane)sender;
            Do(c => c.VisibleChanged -= CustomTaskPaneVisibleChanged);

            //Propagate changes, then raise adapter event
            Do(c =>
                   {
                       if (c != customTaskPane)
                           c.Visible = customTaskPane.Visible;
                   });
            var handler = VisibleChanged;
            if (handler != null)
                handler(this, EventArgs.Empty);

            Do(c => c.VisibleChanged += CustomTaskPaneVisibleChanged);
        }

        void CustomTaskPaneDockPositionChanged(object sender, EventArgs e)
        {
            if (disposed) return;
            var customTaskPane = (CustomTaskPane)sender;
            Do(c => c.DockPositionChanged -= CustomTaskPaneDockPositionChanged);

            //Propagate changes, then raise adapter event
            Do(c =>
            {
                if (c != customTaskPane)
                    c.DockPosition = customTaskPane.DockPosition;
            });
            var handler = DockPositionChanged;
            if (handler != null)
                handler(this, EventArgs.Empty);

            Do(c => c.DockPositionChanged += CustomTaskPaneDockPositionChanged);
        }

        private void Do(Action<CustomTaskPane> action)
        {
            if (disposed) return;
            foreach (var customTaskPane in customTaskPanes.ToArray())
            {
                action(customTaskPane);
            }
        }

        public object ViewContext { get; private set; }

        public UserControl Control
        {
            get { return this.DoOnPane((p) => p.Control); }
        }

        public string Title
        {
            get { return this.DoOnPane((p) => p.Title, string.Empty); }
        }

        public object Window
        {
            get { return this.DoOnPane((p) => p.Window); }
        }

        public Microsoft.Office.Core.MsoCTPDockPosition DockPosition
        {
            get { return this.DoOnPane((p) => p.DockPosition); }
            set { Do(c => c.DockPosition = value); }
        }

        public Microsoft.Office.Core.MsoCTPDockPositionRestrict DockPositionRestrict
        {
            get { return this.DoOnPane((p) => p.DockPositionRestrict); }
            set { Do(c => c.DockPositionRestrict = value); }
        }

        public bool Visible
        {
            get { return this.DoOnPane((p) => p.Visible); }
            set { Do(c => c.Visible = value); }
        }

        public event EventHandler VisibleChanged;
        public event EventHandler DockPositionChanged;

        public int Height
        {
            get { return this.DoOnPane((p) => p.Height); }
            set { Do(c => c.Height = value); }
        }

        public int Width
        {
            get { return this.DoOnPane((p) => p.Width); }
            set { Do(c => c.Width = value); }
        }

        public void Dispose()
        {
            if (disposed) return;
            Do(DisposeTaskPane);
            disposed = true;
        }

        void DisposeTaskPane(CustomTaskPane c)
        {
            c.VisibleChanged -= CustomTaskPaneVisibleChanged;
            c.DockPositionChanged -= CustomTaskPaneDockPositionChanged;
            try
            {
                c.Dispose();
            }
            catch (ObjectDisposedException)
            {
                this.disposed = true;
            }

            customTaskPanes.Remove(c);
        }

        public void CleanupView(object view)
        {
            if (disposed) return;
            foreach (var customTaskPane in customTaskPanes.ToArray())
            {
                try
                {
                    var taskPaneWindow = customTaskPane.Window;
                    if (taskPaneWindow != view) continue;
                    DisposeTaskPane(customTaskPane);
                }
                catch (COMException)
                {
                    customTaskPanes.Remove(customTaskPane);
                }

                CleanupView(view);
                break;
            }
        }

        public void HideIfVisible()
        {
            if (disposed || Visible)
            {
                try
                {
                    Visible = false;
                }
                catch
                {
                    // Do nothing.
                }
                finally
                {
                    hasBeenHidden = true;
                }
            }
        }

        public void RestoreIfNeeded()
        {
            if (!hasBeenHidden)
            {
                return;
            }

            try
            {
                Visible = true;
                hasBeenHidden = false;
            }
            catch
            {
                hasBeenHidden = true;
            }
        }
    }
}
