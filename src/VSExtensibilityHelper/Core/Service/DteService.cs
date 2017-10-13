using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using VSLangProj140;

namespace VSExtensibilityHelper.Core.Service
{
    public class DteService
    {
        #region Fields
               
        private static System.IServiceProvider _globalServiceProvider;

        #endregion Fields

        #region Constructors

        static DteService()
        {         
            DteService.DTE = ServiceLocator.GetInstance<DTE2>();
            DteService._globalServiceProvider = ServiceLocator.GetInstance<IServiceProvider>();            
        }

        #endregion Constructors

        #region Properties

        public static Project ActiveProject
        {
            get
            {
                Project result = null;
                object[] array = DteService.DTE.ActiveSolutionProjects as object[];
                if (array != null && array.Length > 0)
                {
                    result = (array[0] as Project);
                }
                return result;
            }
        }

        public static Solution ActiveSolution
        {
            get
            {
                return DteService.DTE.Solution;
            }
        }

        public static DTE2 DTE
        {
            get;
            private set;
        }

        public static SelectedItem SelectedItem
        {
            get
            {
                return DteService.DTE.SelectedItems.Count > 0 ? DteService.DTE.SelectedItems.Item(1) : null;
            }
        }

        public static Project SelectedProject
        {
            get
            {
                return SelectedItem != null ? SelectedItem.Project : null;
            }
        }

        #endregion Properties

        #region Methods

        public static void CreateNewFile(string fileType, string title, string fileContents)
        {
            Document document = DteService.DTE.ItemOperations.NewFile(fileType, title, "{00000000-0000-0000-0000-000000000000}").Document;
            TextSelection textSelection = document.Selection as TextSelection;
            textSelection.SelectAll();
            textSelection.Text = "";
            textSelection.Insert(fileContents, 1);
            textSelection.StartOfDocument(false);
        }

        public static void CreateNewTextFile(string title, string fileContents)
        {
            DteService.CreateNewFile("General\\Text File", title, fileContents);
        }

        public static ProjectItem FindItemByName(ProjectItems collection, string name, bool recursive)
        {
            if (collection != null)
            {
                foreach (ProjectItem projectItem in collection)
                {
                    if (projectItem.Name.Equals(name, StringComparison.InvariantCultureIgnoreCase))
                    {
                        ProjectItem result = projectItem;
                        return result;
                    }
                    if (recursive)
                    {
                        ProjectItem projectItem2 = DteService.FindItemByName(projectItem.ProjectItems, name, recursive);
                        if (projectItem2 != null)
                        {
                            ProjectItem result = projectItem2;
                            return result;
                        }
                    }
                }
            }
            return null;
        }

        public static string GetDefaultRootNameSpace(Project project)
        {
            return (string)project.Properties.Item("RootNameSpace").Value;
        }

        public static string GetLanguageFileExtension(ProjectItem projectItem)
        {
            string language;
            if ((language = projectItem.ContainingProject.CodeModel.Language) != null)
            {
                if (language == "{B5E9BD34-6D3E-4B5D-925E-8A43B79820B4}")
                {
                    return ".cs";
                }
                if (language == "{B5E9BD33-6D3E-4B5D-925E-8A43B79820B4}")
                {
                    return ".vb";
                }
            }
            throw new NotSupportedException();
        }

        public static string GetOutputAssembly(EnvDTE.Project vsProject)
        {
            string fullPath = vsProject.Properties.Item("FullPath").Value.ToString();
            string outputPath = vsProject.ConfigurationManager.ActiveConfiguration.Properties.Item("OutputPath").Value.ToString();
            string outputDir = Path.Combine(fullPath, outputPath);
            string outputFileName = vsProject.Properties.Item("OutputFileName").Value.ToString();
            string assemblyPath = Path.Combine(outputDir, outputFileName);
            return assemblyPath;
        }

        public static Project GetProjectByName(string projectName, bool isUniqueName = true)
        {
            Project result = null;
            foreach (Project project in DteService.DTE.Solution.Projects)
            {
                if ((isUniqueName && project.UniqueName == projectName) || (!isUniqueName && project.Name == projectName))
                {
                    result = project;
                    break;
                }
            }
            return result;
        }

        public static IEnumerable<string> GetProjectFiles(Project project)
        {
            var projectItems = GetProjectItems(project.ProjectItems);
            return projectItems.Select(a => a.Properties.Item("FullPath").Value.ToString());
        }

        public static IEnumerable<ProjectItem> GetProjectItems(ProjectItems projectItems)
        {
            List<ProjectItem> returnValue = new List<ProjectItem>();

            if (null == projectItems)
            {
                return returnValue.ToArray();
            }

            if (projectItems != null)
            {
                for (var i = 1; i <= projectItems.Count; i++)
                {
                    var subItems = projectItems.Item(i);
                    returnValue.Add(subItems);
                    returnValue.AddRange(GetProjectItems(subItems.ProjectItems));
                }
            }

            return returnValue;
        }

        public static IEnumerable<Project> GetProjects(Solution solution)
        {
            List<Project> projects = new List<Project>();

            var enumerator = solution.Projects.GetEnumerator();
            while (enumerator.MoveNext())
            {
                var project = enumerator.Current as Project;
                projects.Add(project);
            }

            return projects;
        }

        public static IEnumerable<string> GetReferences(EnvDTE.Project project)
        {
            var vsproject = project.Object as VSProject3;

            foreach (VSLangProj.Reference reference in vsproject.References)
            {
                if (reference.SourceProject == null)
                {
                    yield return reference.Path;
                }
            }
        }

        public static bool IsProjectFilesSaved(Project project)
        {
            var projectItems = GetProjectItems(project.ProjectItems);
            return projectItems.All(a => a.Saved);
        }

        public static void SetStatus(string message)
        {
            DteService.DTE.StatusBar.Text = message;
        }

        public static bool TryGetSetting<T>(string settingName, string category, string page, out T value)
        {
            bool result = false;
            value = default(T);

            if (DteService.DTE != null)
            {
                Properties properties = DteService.DTE.get_Properties(category, page);
                if (properties != null)
                {
                    value = (T)((object)properties.Item(settingName).Value);
                    result = true;
                }
            }
            return result;
        }

        #endregion Methods
    }
}