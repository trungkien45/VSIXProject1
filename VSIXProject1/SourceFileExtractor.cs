using EnvDTE;
using EnvDTE80;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.VisualStudio.Shell;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace VSIXProject1
{

    public class SourceFileExtractor
    {
        private readonly DTE2 _dte;

        public SourceFileExtractor(DTE2 dte)
        {
            _dte = dte;
        }

        // Lấy danh sách các file và dự án dưới dạng Class1
        public List<Class1> GetAllSourceFiles()
        {
            List<Class1> files = new List<Class1>();

            Solution solution = _dte.Solution;
            if (solution == null || solution.Projects.Count == 0)
                return files;

            foreach (Project project in solution.Projects)
            {
                files.AddRange(GetFilesInProject(project));
            }

            return files;
        }

        internal IEnumerable<Class1> GetAllSourceFiles(Project project)
        {
            return GetFilesInProject(project);
        }

        // Lấy file từ dự án và tạo các đối tượng Class1
        private IEnumerable<Class1> GetFilesInProject(Project project)
        {
            List<Class1> files = new List<Class1>();
            if (project.ProjectItems == null)
                return files;
            foreach (ProjectItem item in project.ProjectItems)
            {
                files.AddRange(GetFilesInProjectItem(item, project.Name));
            }

            return files;
        }

        // Lấy file từ ProjectItem và tạo đối tượng Class1
        private IEnumerable<Class1> GetFilesInProjectItem(ProjectItem item, string projectName)
        {
            List<Class1> files = new List<Class1>();

            if (item.Kind == EnvDTE.Constants.vsProjectItemKindPhysicalFile)
            {
                string filePath = item.FileNames[1];
                if (File.Exists(filePath))
                {
                    files.AddRange(GetNotImplementedLines(filePath, projectName));
                }
            }
            else if (item.ProjectItems != null)
            {
                foreach (ProjectItem subItem in item.ProjectItems)
                {
                    files.AddRange(GetFilesInProjectItem(subItem, projectName));
                }
            }

            return files;
        }

        // Tìm các dòng chứa "throw new NotImplementedException()" và bỏ qua các comment
        public IEnumerable<Class1> GetNotImplementedLines(string filePath, string projectName)
        {
            List<Class1> notImplementedFiles = new List<Class1>();

            // Đọc tệp theo từng dòng và tìm "throw new NotImplementedException()"
            string[] lines = File.ReadAllLines(filePath);
            for (int i = 0; i < lines.Length; i++)
            {
                string trimmedLine = lines[i].Trim();

                // Kiểm tra nếu dòng chứa throw new NotImplementedException() và không phải comment
                if (IsValidCodeLine(trimmedLine, "throw new NotImplementedException()"))
                {
                    notImplementedFiles.Add(new Class1
                    {
                        Project = projectName,
                        FilePath = filePath,
                        File = Path.GetFileName(filePath),
                        Line = i + 1 // Lưu số dòng (bắt đầu từ 1)
                    });
                }
            }

            return notImplementedFiles;
        }

        // Hàm để xác định dòng hợp lệ, không phải comment
        private bool IsValidCodeLine(string line, string codeSnippet)
        {
            // Kiểm tra xem có chứa codeSnippet và không phải là comment
            if (line.Contains(codeSnippet))
            {
                // Loại bỏ comment "//" hoặc "/*...*/"
                // Nếu dòng chứa "//" trước codeSnippet, nó là comment
                int commentIndex = line.IndexOf("//");
                if (commentIndex != -1 && commentIndex < line.IndexOf(codeSnippet))
                {
                    return false;
                }

                // Kiểm tra với comment kiểu /* */
                if (Regex.IsMatch(line, @"/\*.*\*/"))
                {
                    return false;
                }

                return true;
            }
            return false;
        }

        internal List<Class1> GetNotImplementedLines(string filePath, string projectName, string newContent)
        {
            List<Class1> notImplementedFiles = new List<Class1>();

            // Đọc tệp theo từng dòng và tìm "throw new NotImplementedException()"
            string[] lines = newContent.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None);
            for (int i = 0; i < lines.Length; i++)
            {
                string trimmedLine = lines[i].Trim();

                // Kiểm tra nếu dòng chứa throw new NotImplementedException() và không phải comment
                if (IsValidCodeLine(trimmedLine, "throw new NotImplementedException()"))
                {
                    notImplementedFiles.Add(new Class1
                    {
                        Project = projectName,
                        FilePath = filePath,
                        File = Path.GetFileName(filePath),
                        Line = i + 1 // Lưu số dòng (bắt đầu từ 1)
                    });
                }
            }

            return notImplementedFiles;
        }
    }
}