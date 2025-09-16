﻿using System;
using Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using VSTOContrib.Core.Extensions;
using VSTOContrib.Core.RibbonFactory;
using VSTOContrib.Core.RibbonFactory.Interfaces;
using VSTOContrib.Core.Wpf;
using VSTOContrib.Word.RibbonFactory;
using WikipediaWordAddin.WpfControls;
using Document = Microsoft.Office.Interop.Word.Document;

namespace WikipediaWordAddin.OfficeContexts
{
    public class DocumentViewModel : WordRibbonViewModel, IRegisterCustomTaskPane
    {
        readonly WikipediaResultsViewModel wikipediaResultsViewModel;
        bool panelShown, ribbonVisible;
        ICustomTaskPaneWrapper wikipediaResultsTaskPane;
        Microsoft.Office.Tools.Word.Document vstoDocument;

        public DocumentViewModel(WikipediaResultsViewModel wikipediaResultsViewModel)
        {
            this.wikipediaResultsViewModel = wikipediaResultsViewModel;
        }

        public string GetNoteLabelText(IRibbonControl _)
        {
            return "Test Label Text";
        }

        public bool GetNoteLabelVisible(IRibbonControl _)
        {
            return true;
        }

        public bool GetNoteLabelEnabled(IRibbonControl _)
        {
            return false;
        }

        public override void Initialised(Document document)
        {
            PanelShown = false;

            if (document != null)
            {
                vstoDocument= ((ApplicationFactory)VstoFactory).GetVstoObject(document);
                vstoDocument.SelectionChange += VstoDocumentOnSelectionChange;
                RibbonVisible = true;
            }
            else
            {
                RibbonVisible = false;
            }
        }

        void VstoDocumentOnSelectionChange(object sender, SelectionEventArgs e)
        {
            using (var selection = e.Selection.WithComCleanup())
            {
                wikipediaResultsViewModel.Search(selection.Resource.Text);
            }
        }

        public bool RibbonVisible
        {
            get { return ribbonVisible; }
            set
            {
                ribbonVisible = value;
                OnPropertyChanged(() => RibbonVisible);
            }
        }

        public bool PanelShown
        {
            get { return panelShown; }
            set
            {
                if (panelShown == value) return;
                panelShown = value;
                if (wikipediaResultsTaskPane != null) 
                    wikipediaResultsTaskPane.Visible = value;
                OnPropertyChanged(() => PanelShown);
            }
        }

        public void RegisterTaskPanes(Register register)
        {
            wikipediaResultsTaskPane = register(
                () => new WpfPanelHost(System.Windows.Controls.ScrollBarVisibility.Disabled)
                {
                    Child = new WikipediaResultsView //This is a WPF User control
                    {
                        DataContext = wikipediaResultsViewModel //Viewmodel for the user control
                    }
                }, "Wikipedia Results", PanelShown);
            wikipediaResultsTaskPane.VisibleChanged += TaskPaneVisibleChanged;
        }

        public override void Cleanup()
        {
            wikipediaResultsTaskPane.VisibleChanged -= TaskPaneVisibleChanged;
            vstoDocument.SelectionChange -= VstoDocumentOnSelectionChange;
        }

        private void TaskPaneVisibleChanged(object sender, EventArgs e)
        {
            panelShown = wikipediaResultsTaskPane.Visible;
            OnPropertyChanged(() => PanelShown);
        }
    }
}
