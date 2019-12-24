using Common;
using CxViewerAction.Entities;
using CxViewerAction.Entities.WebServiceEntity;
using CxViewerAction.Helpers;
using CxViewerAction.Helpers.DrawingHelper;
using CxViewerAction.MenuLogic;
using CxViewerAction.Services;
using CxViewerAction.Views.DockedView;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.IO;
using CxViewerAction.Views;
using CxViewerAction.QueryDescription;

namespace CxViewerAction
{
    static public class CommonActionsInstance
    {
        private static CommonActions _commonActions;
        public static CommonActions getInstance()
        {
            if (_commonActions == null)
            {
                _commonActions = new CommonActions();
            }
            return _commonActions;
        }
    }

    public class CommonActions
    {
        #region Fields

        private DTE2 _applicationObject = null;
        private ToolWindowPane _scanProgressWin;
        private ToolWindowPane _graphWin;
        private ToolWindowPane _pathWin;
        private ToolWindowPane _reportWin;
        private ToolWindowPane _resultWin;
        private bool wasInit = false;
        private const string vsProjectKindWeb = "{E24C65DC-7377-472b-9ABA-BC803B73C61A}";
        private const string vsProjectKindSolutionFolder = "{66A26720-8FB5-11D2-AA7E-00C04F688DDE}";
        private Dictionary<string, List<string>> fileMapping = new Dictionary<string, List<string>>();

        #endregion

        #region Properties

        public DTE2 ApplicationObject
        {
            set {

                if (_applicationObject == null)
                {
                    _applicationObject = value;
                }
            }
        }

        public ToolWindowPane ScanProgressWin
        {
            set
            {
                if (_scanProgressWin == null)
                    _scanProgressWin = value;
            }
        }

        public ToolWindowPane GraphWin
        {
            set
            {
                if (_graphWin == null)
                    _graphWin = value;
            }
        }

        public ToolWindowPane ResultWin
        {
            set
            {
                if (_resultWin == null)
                    _resultWin = value;
            }
        }

        public ToolWindowPane PathWin
        {
            set
            {
                if (_pathWin == null)
                    _pathWin = value;
            }
        }
        
        public ToolWindowPane ReportWin
        {
            set
            {
                if (_reportWin == null)
                {
                    _reportWin = value;
                    RegisterReportEvents();
                }
            }
        }

        public IPerspectiveView ReportPersepectiveView
        {
            get
            {
                if (_reportWin != null)
                    return (IPerspectiveView)_reportWin.Window;

                return null;
            }
        }

        public IScanView ScanProgressView
        {
            get
            {
                if (_scanProgressWin != null)
                    return (IScanView)_scanProgressWin.Window;

                return null;
            }
        }

        #endregion

        #region API

        public void BuildFileMapping()
        {
            try
            {

                fileMapping.Clear();

                IList<EnvDTE.Project> projects = GetSolutionProjects();
                foreach (EnvDTE.Project project in projects)
                {
                    if (project.ProjectItems != null)
                    {
                        foreach (ProjectItem projectItem in project.ProjectItems)
                        {
                            BuildFileMapping(projectItem, fileMapping);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Create().Error(ex.ToString());
            }
        }

        /// <summary>
        /// Execute system command, like "Save All", "Close" etc
        /// </summary>
        /// <param name="commandName"></param>
        public void ExecuteSystemCommand(string commandName, string args)
        {
            try
            {
                _applicationObject.ExecuteCommand(commandName, args);
            }
            catch (Exception ex)
            {
                Logger.Create().Error(ex.ToString());

                TopMostMessageBox.Show(string.Format("Can't execute {0} command", commandName));
            }
        }

        /// <summary>
        /// Get project path for current selected project in solution explorer
        /// </summary>
        /// <returns></returns>
        public Entities.Project GetSelectedProject()
        {
            string projectName, projectPath;
            Array projects = (Array)_applicationObject.ActiveSolutionProjects;

            List<string> folderPathList = new List<string>();
            List<string> filePathList = new List<string>(); ;
            //Context menu are displayed on project item in solution explorer
            if (_applicationObject.SelectedItems != null)
            {
                foreach (SelectedItem selectedItem in _applicationObject.SelectedItems)
                {
                    if (selectedItem.ProjectItem != null)
                    {
                        if (selectedItem.ProjectItem.Kind == EnvDTEConstants.vsProjectItemKindPhysicalFolder) // folder
                        {
                            folderPathList.Add(selectedItem.ProjectItem.Properties.Item("FullPath").Value.ToString());
                        }
                        else if (selectedItem.ProjectItem.Kind == EnvDTEConstants.vsProjectItemKindPhysicalFile) // item
                        {
                            filePathList.Add(selectedItem.ProjectItem.Properties.Item("FullPath").Value.ToString());
                        }
                    }
                }

            }
            string projectFullPath = string.Empty;
            try
            {
                if (projects.Length == 0)
                {
                    //Context menu are displayed on solution item in solution explorer

                    Solution solution = _applicationObject.Solution;
                    if (String.IsNullOrEmpty(solution.FileName))
                    {
                        return null;
                    }
                    FileInfo fileInfo = new FileInfo(solution.FileName);

                    Entities.Project outputProject = new Entities.Project(fileInfo.Name, fileInfo.DirectoryName, filePathList, folderPathList);

                    AddProjectToSolution(outputProject, solution.Projects);


                    return outputProject;
                }
                else
                {

                    EnvDTE.Project project = ((EnvDTE.Project)projects.GetValue(0));
                    projectFullPath = project.FullName;

                    // For versions earlier than 2013 we have a bug where project.FullName returns http://localhost:XXXX
                    // The following line returns the project full path for all project kinds.
                    // for versions prior to 2013 for web projects, project.FullName return the project loaction with '/' in the end. Our algorithm is based on that behaviour.
                    // in order to maintain this behaviour, we always trim '//' '\' from fullPAth, and append "//". This way the rest of the code would execute as usual.                    
                    if (project.Kind == vsProjectKindWeb) //if project is web
                    {
                        string webProjectPath = project.Properties.Item("FullPath").Value as string;
                        webProjectPath = webProjectPath.TrimEnd(new[] { '\\', '/' });
                        projectFullPath = webProjectPath + "//";
                    }

                    FileInfo fileInfo = new FileInfo(projectFullPath);

                    projectName = Path.GetFileName(project.Name.TrimEnd(new[] { '\\', '/' }));
                    projectPath = fileInfo.Directory.FullName;

                    return new Entities.Project(projectName, projectPath, filePathList, folderPathList);
                }
            }
            catch (ArgumentException ae)
            {
                Logger.Create().Error(ae.ToString());
                if (false == string.IsNullOrEmpty(projectFullPath))
                {
                    Logger.Create().Error("projectFullPath = " + projectFullPath);
                }
            }
            catch (Exception ex)
            {
                Logger.Create().Error(ex.ToString());
            }
            return null;
        }

        /// <summary>
        /// Get problem file from entire project or solution
        /// </summary>
        /// <param name="project"></param>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <param name="length"></param>
        /// <returns></returns>
        public bool ShowFile(string relativeFileName, int row, int column, int length)
        {
            string fileName = Path.GetFileName(relativeFileName);
            try
            {
                if (fileMapping.ContainsKey(fileName))
                {
                    List<string> solutionFiles = fileMapping[fileName];
                    string[] pathParts = relativeFileName.Split(new[] { '\\', '/' });
                    int depth = pathParts.Length - 2;
                    string pathTail = pathParts[depth + 1];
                    while (solutionFiles.Count > 1 && depth >= 0)
                    {
                        pathTail = Path.Combine(pathParts[depth], pathTail);
                        List<string> candidates = new List<string>();
                        foreach (string solutionFile in solutionFiles)
                        {
                            if (solutionFile.EndsWith(pathTail))
                                candidates.Add(solutionFile);
                        }
                        if (candidates.Count == 0)
                            break;

                        solutionFiles = candidates;
                        depth--;
                    }

                    if (solutionFiles.Count > 0)
                    {
                        if (ShowProblemFile(solutionFiles[0], row, column, length))
                            return true;
                    }

                }

            }
            catch (Exception ex)
            {
                Logger.Create().Error(ex.ToString());
            }
            System.Windows.Forms.MessageBox.Show(string.Format("File {0} not found", relativeFileName), "Error", System.Windows.Forms.MessageBoxButtons.OK);
            return false;
        }

        public void reportWinObject_SelectedNodeChanged(CxViewerAction.Entities.WebServiceEntity.TreeNodeData obj)
        {
            ShowProblemFile(obj);
        }

        public void NavigateToQueryDescription(object sender, EventArgs e)
        {
            try
            {
                QueryDescriptionEventArg nodeData = (QueryDescriptionEventArg)e;
                CxRESTApiPortalConfiguration rESTApiPortalConfiguration = new CxRESTApiPortalConfiguration();
                rESTApiPortalConfiguration.InitPortalBaseUrl();
                string urlToDescription = new QueryDescriptionUrlBuilder().Build(nodeData.QueryId, nodeData.QueryName, nodeData.QueryVersionCode);
                
                _applicationObject.ItemOperations.Navigate(urlToDescription, vsNavigateOptions.vsNavigateOptionsDefault);
            }
            catch (Exception ex)
            {
                Logger.Create().Error(ex.ToString());
                TopMostMessageBox.Show(ex.Message);
            }
        }

        public void OpenQueryDescription(string url)
        {
            try
            {
                _applicationObject.ItemOperations.Navigate(url, vsNavigateOptions.vsNavigateOptionsDefault);
            }
            catch (Exception ex)
            {
                Logger.Create().Error(ex.ToString());

                TopMostMessageBox.Show(ex.Message);
            }
        }

        public void reportWinObject_SelectedScanChanged(long scanId)
        {
            try
            {

                CommonData.SelectedScanId = scanId;


                ShowResultLogic showResultLogic = new ShowResultLogic();

                showResultLogic.Act();

                #region Remarks
                //Commands2 commands = (Commands2)_applicationObject.Commands;
                //EnvDTE.Command prevCommand;

                //prevCommand = commands.Item("CxViewerAction.Connect.ShowResults", 1);

                //object customin = null, customout = null;
                //commands.Raise(prevCommand.Guid, prevCommand.ID, ref customin, ref customout);
                #endregion
            }
            catch (Exception ex)
            {
                Logger.Create().Error(ex.ToString());
            }

        }

        public void ShowScanProgressView()
        {
            showView(_scanProgressWin);
        }

        public void CloseScanProgressView()
        {
            closeView(_scanProgressWin);
        }

        public void ClearScanProgressView()
        {
            var dockView = (IScanView)_scanProgressWin.Window;
            dockView.Clear();
        }

        public void ShowGraphView()
        {
            showView(_graphWin);
        }

        public void CloseGraphView()
        {
            closeView(_graphWin);
        }

        public void ShowResultsView()
        {
            showView(_resultWin);
        }

        public void CloseResultsView()
        {
            closeView(_resultWin);
        }

        public void ShowPathView()
        {
            showView(_pathWin);
        }

        public void ClosePathView()
        {
            closeView(_pathWin);
        }

        public void ShowReportView()
        {
            showView(_reportWin);
        }

        public void CloseReportView()
        {
            closeView(_reportWin);
        }

        public void ReportDoPrevResults()
        {
            IPerspectiveView rep = _reportWin.Window as IPerspectiveView;
            PerspectiveHelper.DoPrevResult();
            if (rep == null || rep.Report == null || rep.Report.Tree.Count == 0)
            {
                TopMostMessageBox.Show("There are no vulnerabilities to show");
            }
        }

        public void UpdateScanProgress(ScanStatusBar data)
        {
            if (data == null) return;

            if (data.ClearBeforeUpdateProgress)
            {
                _applicationObject.StatusBar.Clear();
            }

            _applicationObject.StatusBar.Progress(data.InProgress,
                        data.Label, data.Completed, data.Total);
        }

        #endregion

        #region Private methods

        private void RegisterReportEvents()
        {
            var view = _reportWin.Window as IPerspectiveView;
            if (view != null)
            {
                view.SelectedNodeChanged -= reportWinObject_SelectedNodeChanged;
                view.SelectedReportItemChanged -= NavigateToQueryDescription;
                view.SelectedScanChanged -= reportWinObject_SelectedScanChanged;

                view.SelectedNodeChanged += reportWinObject_SelectedNodeChanged;
                view.SelectedReportItemChanged += NavigateToQueryDescription;
                view.SelectedScanChanged += reportWinObject_SelectedScanChanged;
            }
        }

        private void BuildFileMapping(ProjectItem projectItem, Dictionary<string, List<string>> mapping)
        {
            try
            {

                AddFilesToMappingTable(projectItem, mapping);

                if (projectItem.ProjectItems != null && projectItem.ProjectItems.Count > 0)
                {
                    foreach (ProjectItem projectSubItem in projectItem.ProjectItems)
                    {
                        BuildFileMapping(projectSubItem, mapping);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Create().Error(ex.ToString());
            }
        }

        private IList<EnvDTE.Project> GetSolutionProjects()
        {
            List<EnvDTE.Project> list = new List<EnvDTE.Project>();

            try
            {

                Projects projects = _applicationObject.Solution.Projects;

                foreach (EnvDTE.Project project in projects)
                {
                    if (project == null)
                    {
                        continue;
                    }
                    if (project.Kind == ProjectKinds.vsProjectKindSolutionFolder)
                    {
                        list.AddRange(GetSolutionFolderProjects(project));
                    }
                    else
                    {
                        list.Add(project);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Create().Error(ex.ToString());
            }

            return list;
        }

        private IEnumerable<EnvDTE.Project> GetSolutionFolderProjects(EnvDTE.Project solutionFolder)
        {

            List<EnvDTE.Project> list = new List<EnvDTE.Project>();

            try
            {

                foreach (ProjectItem projectItem in solutionFolder.ProjectItems)
                {
                    EnvDTE.Project subProject = projectItem.SubProject;
                    if (subProject == null)
                    {
                        continue;
                    }
                    if (subProject.Kind == ProjectKinds.vsProjectKindSolutionFolder)
                    {
                        list.AddRange(GetSolutionFolderProjects(subProject));
                    }
                    else
                    {
                        list.Add(subProject);
                    }
                }

            }
            catch (Exception ex)
            {
                Logger.Create().Error(ex.ToString());
            }

            return list;
        }

        private void AddFilesToMappingTable(ProjectItem projectItem, Dictionary<string, List<string>> mapping)
        {
            try
            {

                for (short j = 1; j < projectItem.FileCount + 1; j++)
                {
                    string file = projectItem.get_FileNames(j);
                    string fileName = Path.GetFileName(file);

                    if (string.IsNullOrEmpty(fileName))
                        continue;

                    List<string> filePaths;
                    if (mapping.ContainsKey(fileName))
                        filePaths = mapping[fileName];
                    else
                    {
                        filePaths = new List<string>();
                        mapping.Add(fileName, filePaths);
                    }
                    filePaths.Add(file);
                }
            }
            catch (Exception ex)
            {
                Logger.Create().Error(ex.ToString());
            }
        }

        private void AddProjectToSolution(CxViewerAction.Entities.Project outputProject, Projects projects)
        {
            foreach (EnvDTE.Project solutionProject in projects)
            {
                try
                {
                    if (!string.IsNullOrEmpty(solutionProject.FullName))
                    {
                        string projectFullPath = solutionProject.FullName;

                        // For version 2013 we have a bug where project.FullName returns http://localhost:XXXX
                        // The following line returns the project full path for all project kinds.
                        // for versions prior to 2013 for web projects, project.FullName return the project loaction with '/' in the end. Our algorithm is based on that behaviour.
                        // in order to maintain this behaviour, we always trim '//' '\' from fullPAth, and append "//". This way the rest of the code would execute as usual.                    
                        if (solutionProject.Kind == vsProjectKindWeb) //if project is web
                        {
                            string webProjectPath = solutionProject.Properties.Item("FullPath").Value as string;
                            webProjectPath = webProjectPath.TrimEnd(new[] { '\\', '/' });
                            projectFullPath = webProjectPath + "//";
                            FileInfo fileInfo = new FileInfo(projectFullPath);
                            projectFullPath = fileInfo.FullName;
                        }

                        outputProject.ProjectPaths.Add(new Entities.Project(solutionProject.Name, new FileInfo(projectFullPath).DirectoryName));

                    }
                    else // can be virtual folder
                    {
                        AddProjectToSolution(outputProject, solutionProject.ProjectItems);
                    }
                }
                catch (Exception ex)
                {
                    Logger.Create().Error(ex.ToString());
                }
            }
        }

        private void AddProjectToSolution(CxViewerAction.Entities.Project outputProject, ProjectItems projectItems)
        {
            foreach (EnvDTE.ProjectItem solutionProject in projectItems)
            {
                try
                {
                    if (solutionProject.SubProject == null)
                    {
                        continue;
                    }
                    if (!string.IsNullOrEmpty(solutionProject.SubProject.FullName))
                    {
                        outputProject.ProjectPaths.Add(new Entities.Project(solutionProject.SubProject.Name, new FileInfo(solutionProject.SubProject.FullName).DirectoryName));
                    }
                    else // can be virtual folder
                    {
                        if (solutionProject.SubProject.ProjectItems == null)
                        {
                            continue;
                        }
                        AddProjectToSolution(outputProject, solutionProject.SubProject.ProjectItems);
                    }
                }
                catch (Exception ex)
                {
                    Logger.Create().Error(ex.ToString());
                }
            }
        }

        /// <summary>
        /// Show selected project file
        /// </summary>
        /// <param name="file"></param>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <param name="length"></param>
        /// <returns></returns>
        private bool ShowProblemFile(string file, int row, int column, int length)
        {
            FileInfo fileInfo = new FileInfo(file);
        
            if (fileInfo.Exists)
            {
                try
                {
                    _applicationObject.ItemOperations.OpenFile(file, EnvDTEConstants.vsViewKindCode);
                    _applicationObject.ActiveDocument.Activate();

                    TextSelection selection = (TextSelection)_applicationObject.ActiveDocument.Selection;
                    try
                    {
                        selection.MoveToLineAndOffset(row, column, false);
                        selection.CharRight(true, length);
                    }
                    catch (ArgumentException ex)
                    {
                        if (IsJavaScriptFile(fileInfo))
                        {
                            string errMsg = "“This plugin does not support showing results in a compressed min.js file. \n" + 
                                            "To view the full results, please navigate to the Checkmarx results viewer.";
                            TopMostMessageBox.Show(errMsg);

                            return true;
                        }
                    }
                    return true;
                }
                catch (Exception ex)
                {
                    Logger.Create().Error(ex.ToString());

                    TopMostMessageBox.Show(ex.Message);
                }
            }

            return false;
        }

        public bool IsJavaScriptFile(FileInfo fileInfo)
        {
            bool isJSFile = (fileInfo.FullName != null || fileInfo.Name != null) &&
          fileInfo.Extension.Equals(".js", StringComparison.OrdinalIgnoreCase);
          
            return isJSFile;
        }


        private void ShowProblemFile(CxViewerAction.Entities.WebServiceEntity.TreeNodeData treeNode)
        {

            #region [Bind graph view]

            try
            {
                PerspectiveGraphCtrl viewGraph = null;
                if (_graphWin != null)
                {
                    viewGraph = _graphWin.Window as PerspectiveGraphCtrl;
                    if (viewGraph != null)
                    {
                        viewGraph.ClearGraphView();
                        viewGraph.Graph = null;
                        viewGraph.SelectedPath = null;
                        viewGraph.Graph = new Graph(treeNode);
                        viewGraph.MsGalViewer.Refresh();
                        viewGraph.MsGalViewer.ResumeLayout();
                        viewGraph.MsGalViewer.Update();
                        viewGraph.BindData();
                        viewGraph.PathItemClick = GraphClick;
                    }

                    showView(_graphWin);
                }

                #endregion

                #region [Bind result view]
                if (_resultWin != null)
                {
                    PerspectiveResultCtrl viewResult = _resultWin.Window as PerspectiveResultCtrl;

                    viewResult.SelectedNode = treeNode;
                    //if (!_resultWin.Visible || viewResult.IsActive)
                    //{
                    if (!wasInit)
                    {
                        viewResult.SelectedRowChanged += new EventHandler(viewResult_SelectedRowChanged);
                        viewResult.Refresh += new EventHandler(viewResult_Refresh);
                        wasInit = true;
                    }
                    // _resultWin.Visible = true;
                    viewResult.IsActive = false;
                    viewResult.SelectRow();
                    //}

                    showView(_resultWin);

                }
            }
            catch (Exception ex)
            {

                if (ex is System.Net.WebException)
                {
                    Logger.Create().Error(ex.ToString());
                    TopMostMessageBox.Show(ex.Message, "Error");
                }
                else
                {
                    Logger.Create().Error(ex.ToString());
                    TopMostMessageBox.Show("General error occured, please check the log", "Error");
                }
            }
            #endregion
        }

        private void viewResult_Refresh(object sender, EventArgs e)
        {
            TreeNodeData nodeData = (TreeNodeData)e;
            ShowProblemFile(nodeData);
        }

        private void viewResult_SelectedRowChanged(object sender, EventArgs e)
        {
            try
            {
                ResultData data = (ResultData)e;
                CxViewerAction.CxVSWebService.CxWSResultPath resultPath = PerspectiveHelper.GetResultPath(data.ScanId, data.Result.PathId);

                PerspectiveGraphCtrl viewGraph = null;
                if (_graphWin != null)
                {
                    viewGraph = _graphWin.Window as PerspectiveGraphCtrl;
                    if (viewGraph != null)
                    {
                        CxViewerAction.BaseInterfaces.IGraphPath path = viewGraph.FindPath(resultPath);
                        viewGraph.SelectEdgeGraphByPath(path.DirectFlow[0], path.DirectFlow[1], path);
                        viewGraph.BindData();
                        viewGraph.PathItemClick = GraphClick;
                    }
                }

                #region [Bind path view]
                if (_pathWin != null)
                {
                    IPerspectivePathView viewPath = _pathWin.Window as IPerspectivePathView;
                    CxViewerAction.Entities.WebServiceEntity.ReportQueryItemResult path = new CxViewerAction.Entities.WebServiceEntity.ReportQueryItemResult()
                    {
                        Column = resultPath.Nodes[0].Column,
                        FileName = resultPath.Nodes[0].FileName,
                        Line = resultPath.Nodes[0].Line,
                        NodeId = resultPath.Nodes[0].PathNodeId,
                        PathId = resultPath.PathId,
                        Query = data.NodeData.QueryResult
                    };
                    path.Paths = GraphHelper.ConvertNodesToPathes(resultPath.Nodes, data.NodeData.QueryResult, path);
                    viewPath.PathButtonClickHandler = PathButtonClick;

                    viewPath.QueryItemResult = path;

                    viewPath.BindData(resultPath.Nodes[0].PathNodeId);

                    showView(_pathWin);
                }
                #endregion

                ShowFile(resultPath.Nodes[0].FileName, resultPath.Nodes[0].Line, resultPath.Nodes[0].Column, resultPath.Nodes[0].Length);

            }
            catch (Exception ex)
            {
                if (ex is System.Net.WebException)
                {
                    Logger.Create().Error(ex.ToString());
                    TopMostMessageBox.Show(ex.Message, "Error");
                }
                else
                {
                    Logger.Create().Error(ex.ToString());
                    TopMostMessageBox.Show("General error occured, please check the log", "Error");
                }
            }
        }

        private void GraphClick(object sender, EventArgs e)
        {
            ReportQueryItemPathResult graphItem = ((ReportQueryItemPathResult)sender);
            PerspectiveGraphCtrl viewGraph = null;
            if (_graphWin != null)
            {
                viewGraph = _graphWin.Window as PerspectiveGraphCtrl;
                if (viewGraph != null)
                {
                    viewGraph.SelectedPath = viewGraph.FindPath(graphItem.QueryItem);
                    DrawingHelper.SelectedPathItemUniqueID = graphItem.UniqueID;
                    DrawingHelper.isEdgeSelected = false;
                    if (viewGraph.MsGalViewer != null)
                    {
                        viewGraph.MsGalViewer.Refresh();
                        viewGraph.MsGalViewer.ResumeLayout();
                        viewGraph.MsGalViewer.Update();
                    }
                    viewGraph.BindData();
                }
            }

            #region [Bind path view]
            if (_pathWin != null)
            {
                IPerspectivePathView viewPath = _pathWin.Window as IPerspectivePathView;
                viewPath.PathButtonClickHandler = PathButtonClick;

                viewPath.QueryItemResult = graphItem.QueryItem;

                viewPath.BindData(graphItem.NodeId);

                showView(_pathWin);
            }
            #endregion

            PerspectiveResultCtrl viewResult = _resultWin.Window as PerspectiveResultCtrl;
            viewResult.MarkRowAsSelected(graphItem.QueryItem.PathId);

            ShowFile(graphItem.FileName, graphItem.Line, graphItem.Column, graphItem.Length);

        }

        private void PathButtonClick(object sender, EventArgs e)
        {
            ReportQueryItemPathResult reportQueryItemPathResult = ((ColorButton.ColorButton)sender).Tag as ReportQueryItemPathResult;
            PerspectiveGraphCtrl viewGraph = null;
            if (_graphWin != null)
            {
                viewGraph = _graphWin.Window as PerspectiveGraphCtrl;
                if (viewGraph != null)
                {
                    {
                        viewGraph.SelectedPath = viewGraph.FindPath(reportQueryItemPathResult.QueryItem);
                        DrawingHelper.SelectedPathItemUniqueID = reportQueryItemPathResult.UniqueID;
                        DrawingHelper.isEdgeSelected = false;
                    }

                    viewGraph.BindData();

                    if (viewGraph.MsGalViewer != null)
                    {
                        viewGraph.MsGalViewer.Refresh();
                        viewGraph.MsGalViewer.ResumeLayout();
                        viewGraph.MsGalViewer.Update();
                    }
                }
            }

            ShowFile(reportQueryItemPathResult.FileName, reportQueryItemPathResult.Line, reportQueryItemPathResult.Column, reportQueryItemPathResult.Length);

        }

        private void showView(ToolWindowPane window)
        {
            if ((null == window) || (null == window.Frame))
            {
                throw new NotSupportedException("Cannot create tool window");
            }

            IVsWindowFrame windowFrame = (IVsWindowFrame)window.Frame;
            Microsoft.VisualStudio.ErrorHandler.ThrowOnFailure(windowFrame.Show());
        }

        private void closeView(ToolWindowPane window)
        {
            if ((null == window) || (null == window.Frame))
            {
                throw new NotSupportedException("Cannot create tool window");
            }

            IVsWindowFrame windowFrame = (IVsWindowFrame)window.Frame;
            Microsoft.VisualStudio.ErrorHandler.ThrowOnFailure(windowFrame.CloseFrame((uint)__FRAMECLOSE.FRAMECLOSE_PromptSave));
        } 
        
        #endregion
    }
}
