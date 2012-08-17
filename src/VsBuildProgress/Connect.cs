using System;
using Extensibility;
using EnvDTE;
using EnvDTE80;
using Microsoft.WindowsAPICodePack.Taskbar;
using System.Drawing;

namespace VsBuildProgress
{
	public class Connect : IDTExtensibility2
	{
        private DTE2 _applicationObject;
        private BuildEvents _buildEvents;
        private int _numberOfProjectsBuilt;
        private int _numberOfProjects;

        private bool _buildErrorDetected;

		public void OnConnection(object application, ext_ConnectMode connectMode, object addInInstance, ref Array custom)
		{
			_applicationObject = (DTE2)application;

            _buildEvents = _applicationObject.Events.BuildEvents;
            _buildEvents.OnBuildBegin += OnBuildBegin;
            _buildEvents.OnBuildProjConfigDone += OnBuildProjConfigDone;
            _buildEvents.OnBuildDone += OnBuildDone;
		}

		private void OnBuildBegin(vsBuildScope scope, vsBuildAction action)
        {
            InitialiseTaskBar();

		    _buildErrorDetected = false;
		    _numberOfProjectsBuilt = 0;
		    _numberOfProjects = _applicationObject.Solution.SolutionBuild.BuildDependencies.Count;

		    UpdateProgressValueAndState(false);
        }

        private void OnBuildProjConfigDone(string project, string projectConfig, string platform, string solutionConfig, bool success)
        {
            UpdateProgressValueAndState(!success);
        }

        private void OnBuildDone(vsBuildScope scope, vsBuildAction action)
        {
            TaskbarManager.Instance.SetProgressState(_buildErrorDetected
                                                         ? TaskbarProgressBarState.Error
                                                         : TaskbarProgressBarState.Normal);

            System.Threading.Thread.Sleep(100);
        }

	    private static void InitialiseTaskBar()
        {
            TaskbarManager.Instance.SetOverlayIcon(null, string.Empty);
            TaskbarManager.Instance.SetProgressState(TaskbarProgressBarState.NoProgress);
        }

        private void UpdateProgressValueAndState(bool errorThrown)
        {
            if (_numberOfProjectsBuilt < 0)
                _numberOfProjectsBuilt = 0;

            if (_numberOfProjectsBuilt > _numberOfProjects)
                _numberOfProjectsBuilt = _numberOfProjects;

            TaskbarManager.Instance.SetProgressValue(_numberOfProjectsBuilt, _numberOfProjects - 1);

            if (errorThrown)
            {
                _buildErrorDetected = true;
                TaskbarManager.Instance.SetProgressState(TaskbarProgressBarState.Error);
            }

            _numberOfProjectsBuilt++;
        }

        #region Unimplemented methods of IDTExtensibility2

        public void OnAddInsUpdate(ref Array custom)
        {
            
        }

        public void OnBeginShutdown(ref Array custom)
        {
            
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            
        }
        
        public void OnStartupComplete(ref Array custom)
        {
            
        }

        #endregion
    }
}