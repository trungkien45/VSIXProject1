using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
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
        private Timer timer;
        DTE2 dte;
        /// <summary>
        /// Initializes a new instance of the <see cref="ToolWindow1Control"/> class.
        /// </summary>
        public ToolWindow1Control()
        {
            dte = Package.GetGlobalService(typeof(DTE)) as DTE2;
            this.InitializeComponent();
            this.timer = new Timer();
            timer.Interval = 10000;
            timer.Elapsed += new ElapsedEventHandler(timer_Elapsed);
        }

        private void timer_Elapsed(object sender, ElapsedEventArgs e)
        {

            IVsSolution solution = (IVsSolution)Package.GetGlobalService(typeof(IVsSolution));
            if (solution != null)
            {
                solution.GetSolutionInfo(out string solutionDirectory, out string solutionName, out string solutionDirectory2);
                var solutionPath = solutionDirectory + System.IO.Path.GetFileNameWithoutExtension(solutionName);

                if (!string.IsNullOrEmpty(solutionPath) && dte != null)
                {
                    var extractor = new SourceFileExtractor(dte);
                    var sourceFiles = extractor.GetAllSourceFiles();
                    this.Dispatcher.Invoke(() =>
                    {
                        dataGrid1.ItemsSource = sourceFiles;
                    });
                }
            }
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

            timer.Start();
        }
    }
}