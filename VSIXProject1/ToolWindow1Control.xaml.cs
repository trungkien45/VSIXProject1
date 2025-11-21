using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.IO;
using System.Timers;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace VSIXProject1
{
    /// <summary>
    /// Interaction logic for ToolWindow1Control.
    /// </summary>
    public partial class ToolWindow1Control : UserControl
    {
        List<Class1> sourceFiles = new List<Class1>();
        private Timer timer;
        DTE2 dte;
        private Events2 _events;
        private ProjectItemsEvents _projectItemsEvents;

        /// <summary>
        /// Initializes a new instance of the <see cref="ToolWindow1Control"/> class.
        /// </summary>
        public ToolWindow1Control()
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            dte = Package.GetGlobalService(typeof(DTE)) as DTE2;
            _events = (Events2)dte.Events;
            _events.SolutionEvents.Opened += SolutionOpened;
            _events.SolutionEvents.AfterClosing += SolutionEvents_AfterClosing; 
            _events.SolutionEvents.ProjectAdded += ProjectAdded;
            _events.SolutionEvents.ProjectRemoved += ProjectRemoved;
            _events.SolutionEvents.ProjectRenamed += SolutionEvents_ProjectRenamed;
            var _textEvents = _events.get_TextEditorEvents();
            _textEvents.LineChanged += LineChanged;
            _projectItemsEvents = _events.ProjectItemsEvents;
            _projectItemsEvents.ItemAdded += OnItemAdded;
            _projectItemsEvents.ItemRemoved += _projectItemsEvents_ItemRemoved;
            _projectItemsEvents.ItemRenamed += OnItemRenamed;
            this.InitializeComponent();
            this.timer = new Timer();
            timer.Interval = 10000;
            timer.Elapsed += new ElapsedEventHandler(timer_Elapsed);
        }

        //Tested
        private void LineChanged(TextPoint StartPoint, TextPoint EndPoint, int Hint)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            Document doc = dte.ActiveDocument;
            if (doc == null) return;
            if (sourceFiles == null) return;
            // Chỉ bắt file .cs
            if (!doc.FullName.EndsWith(".cs")) return;
                sourceFiles.RemoveAll(s => s.FilePath == doc.FullName);
            var extractor = new SourceFileExtractor(dte);
            TextDocument textDoc = (TextDocument)doc.Object("TextDocument");
            EditPoint editPoint = textDoc.StartPoint.CreateEditPoint();
            string newContent = editPoint.GetText(textDoc.EndPoint);
            var x = extractor.GetNotImplementedLines(filePath: doc.FullName, projectName: doc.ProjectItem.ContainingProject.Name, newContent);

            sourceFiles.AddRange(x);
            dataGrid1.ItemsSource = null;
            dataGrid1.ItemsSource = sourceFiles;
        }
        private void SolutionEvents_ProjectRenamed(Project Project, string OldName)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var x = sourceFiles.FindAll(s => s.Project == Path.GetFileNameWithoutExtension(OldName));
            for (int i = 0; i < x.Count; i++)
            {
                x[i].Project = Project.Name;
            }
            dataGrid1.ItemsSource = null;
            dataGrid1.ItemsSource = sourceFiles;
        }
        private void OnItemRenamed(ProjectItem ProjectItem, string OldName)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var extractor = new SourceFileExtractor(dte);
            var project = ProjectItem.ContainingProject;
            var x = extractor.GetAllSourceFiles(project);
            foreach (var item in x)
            {
                sourceFiles.RemoveAll(s => s.Project == item.Project);
            }
            sourceFiles.RemoveAll(sourceFiles => sourceFiles.Project == project.Name);
            sourceFiles.AddRange(extractor.GetAllSourceFiles(project));
            dataGrid1.ItemsSource = null;
            dataGrid1.ItemsSource = sourceFiles;
        }

        //Tested
        private void SolutionEvents_AfterClosing()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            sourceFiles.Clear();
            dataGrid1.ItemsSource = null;
            dataGrid1.ItemsSource = sourceFiles;
        }

        //Tested
        private void _projectItemsEvents_ItemRemoved(ProjectItem ProjectItem)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var extractor = new SourceFileExtractor(dte);
            var project = ProjectItem.ContainingProject;
            var x = extractor.GetAllSourceFiles(project);
            foreach (var item in x)
            {
                sourceFiles.RemoveAll(s => s.Project == item.Project);
            }
            sourceFiles.RemoveAll(sourceFiles => sourceFiles.Project == project.Name);
            sourceFiles.AddRange(extractor.GetAllSourceFiles(project));
            dataGrid1.ItemsSource = null;
            dataGrid1.ItemsSource = sourceFiles;
        }

        //Tested
        private void OnItemAdded(ProjectItem ProjectItem)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var extractor = new SourceFileExtractor(dte);
            var project = ProjectItem.ContainingProject;
            var x = extractor.GetAllSourceFiles(project);
            foreach (var item in x)
            {
                sourceFiles.RemoveAll(s => s.Project == item.Project);
            }
            sourceFiles.RemoveAll(sourceFiles => sourceFiles.Project == project.Name);
            sourceFiles.AddRange(extractor.GetAllSourceFiles(project));
            dataGrid1.ItemsSource = null;
            dataGrid1.ItemsSource = sourceFiles;
        }
        //Tested
        private void ProjectRemoved(Project Project)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var extractor = new SourceFileExtractor(dte);
            var x = extractor.GetAllSourceFiles(Project);
            foreach (var item in x)
            {
                sourceFiles.RemoveAll(s => s.Project == item.Project);
            }
            sourceFiles.RemoveAll(sourceFiles => sourceFiles.Project == Project.Name);
            dataGrid1.ItemsSource = null;
            dataGrid1.ItemsSource = sourceFiles;
        }
        //Tested
        private void ProjectAdded(Project Project)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var extractor = new SourceFileExtractor(dte);
            sourceFiles.AddRange(extractor.GetAllSourceFiles(Project));
            dataGrid1.ItemsSource = null;
            dataGrid1.ItemsSource = sourceFiles;
        }

        //Tested
        private void SolutionOpened()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var extractor = new SourceFileExtractor(dte);
            sourceFiles = extractor.GetAllSourceFiles();
            dataGrid1.ItemsSource = null;
            dataGrid1.ItemsSource = sourceFiles;
        }

        private void timer_Elapsed(object sender, ElapsedEventArgs e)
        {

            ThreadHelper.JoinableTaskFactory.Run(async delegate
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                IVsSolution solution = (IVsSolution)Package.GetGlobalService(typeof(IVsSolution));
                if (solution != null)
                {
                    solution.GetSolutionInfo(out string solutionDirectory, out string solutionName, out string solutionDirectory2);
                    var solutionPath = solutionDirectory + System.IO.Path.GetFileNameWithoutExtension(solutionName);

                    if (!string.IsNullOrEmpty(solutionPath) && dte != null)
                    {
                        var extractor = new SourceFileExtractor(dte);
                        var sourceFiles = extractor.GetAllSourceFiles();
                        dataGrid1.ItemsSource = null;
                        dataGrid1.ItemsSource = sourceFiles;
                    }
                }
            });
        }

        private void MyMouseLeftButtonDownHandler(object sender, MouseButtonEventArgs e)
        {
            var point = e.GetPosition(dataGrid1);
            var hitTestResult = VisualTreeHelper.HitTest(dataGrid1, point);

            if (hitTestResult != null)
            {
                var clickedItem = hitTestResult.VisualHit;

                while (clickedItem != null && !(clickedItem is DataGridRow))
                {
                    clickedItem = VisualTreeHelper.GetParent(clickedItem);
                }

                if (clickedItem is DataGridRow row)
                {
                    var dataContext = row.DataContext as Class1; // Lấy DataContext của hàng
                    if (dataContext != null)
                    {
                        var window = dte.ItemOperations.OpenFile(dataContext.FilePath);

                        // Đảm bảo tệp đã được mở
                        if (window != null)
                        {
                            // Di chuyển đến dòng cụ thể
                            var textDocument = (TextDocument)window.Document.Object("TextDocument");
                            textDocument.Selection.MoveToLineAndOffset(dataContext.Line, 1);
                            window.Activate();
                        }
                    }
                }
            }
        }

        private void MyToolWindow_Loaded(object sender, RoutedEventArgs e)
        {
            var extractor = new SourceFileExtractor(dte);
            var sourceFiles = extractor.GetAllSourceFiles();
            dataGrid1.ItemsSource = null;
            dataGrid1.ItemsSource = sourceFiles;
        }
    }
}