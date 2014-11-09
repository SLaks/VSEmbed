using System;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell.Interop;

namespace VSEmbed.Services {
	// This file contains stub implementations of services that we don't actually implement.

	class StubVsActivityLog : IVsActivityLog {
		public int LogEntry([ComAliasName("Microsoft.VisualStudio.Shell.Interop.ACTIVITYLOG_ENTRYTYPE")]uint actType, [ComAliasName("Microsoft.VisualStudio.OLE.Interop.LPCOLESTR")]string pszSource, [ComAliasName("Microsoft.VisualStudio.OLE.Interop.LPCOLESTR")]string pszDescription) {
			return 0;
		}

		public int LogEntryGuid([ComAliasName("Microsoft.VisualStudio.Shell.Interop.ACTIVITYLOG_ENTRYTYPE")]uint actType, [ComAliasName("Microsoft.VisualStudio.OLE.Interop.LPCOLESTR")]string pszSource, [ComAliasName("Microsoft.VisualStudio.OLE.Interop.LPCOLESTR")]string pszDescription, Guid guid) {
			return 0;
		}

		public int LogEntryGuidHr([ComAliasName("Microsoft.VisualStudio.Shell.Interop.ACTIVITYLOG_ENTRYTYPE")]uint actType, [ComAliasName("Microsoft.VisualStudio.OLE.Interop.LPCOLESTR")]string pszSource, [ComAliasName("Microsoft.VisualStudio.OLE.Interop.LPCOLESTR")]string pszDescription, Guid guid, int hr) {
			return 0;
		}

		public int LogEntryGuidHrPath([ComAliasName("Microsoft.VisualStudio.Shell.Interop.ACTIVITYLOG_ENTRYTYPE")]uint actType, [ComAliasName("Microsoft.VisualStudio.OLE.Interop.LPCOLESTR")]string pszSource, [ComAliasName("Microsoft.VisualStudio.OLE.Interop.LPCOLESTR")]string pszDescription, Guid guid, int hr, [ComAliasName("Microsoft.VisualStudio.OLE.Interop.LPCOLESTR")]string pszPath) {
			return 0;
		}

		public int LogEntryGuidPath([ComAliasName("Microsoft.VisualStudio.Shell.Interop.ACTIVITYLOG_ENTRYTYPE")]uint actType, [ComAliasName("Microsoft.VisualStudio.OLE.Interop.LPCOLESTR")]string pszSource, [ComAliasName("Microsoft.VisualStudio.OLE.Interop.LPCOLESTR")]string pszDescription, Guid guid, [ComAliasName("Microsoft.VisualStudio.OLE.Interop.LPCOLESTR")]string pszPath) {
			return 0;
		}

		public int LogEntryHr([ComAliasName("Microsoft.VisualStudio.Shell.Interop.ACTIVITYLOG_ENTRYTYPE")]uint actType, [ComAliasName("Microsoft.VisualStudio.OLE.Interop.LPCOLESTR")]string pszSource, [ComAliasName("Microsoft.VisualStudio.OLE.Interop.LPCOLESTR")]string pszDescription, int hr) {
			return 0;
		}

		public int LogEntryHrPath([ComAliasName("Microsoft.VisualStudio.Shell.Interop.ACTIVITYLOG_ENTRYTYPE")]uint actType, [ComAliasName("Microsoft.VisualStudio.OLE.Interop.LPCOLESTR")]string pszSource, [ComAliasName("Microsoft.VisualStudio.OLE.Interop.LPCOLESTR")]string pszDescription, int hr, [ComAliasName("Microsoft.VisualStudio.OLE.Interop.LPCOLESTR")]string pszPath) {
			return 0;
		}

		public int LogEntryPath([ComAliasName("Microsoft.VisualStudio.Shell.Interop.ACTIVITYLOG_ENTRYTYPE")]uint actType, [ComAliasName("Microsoft.VisualStudio.OLE.Interop.LPCOLESTR")]string pszSource, [ComAliasName("Microsoft.VisualStudio.OLE.Interop.LPCOLESTR")]string pszDescription, [ComAliasName("Microsoft.VisualStudio.OLE.Interop.LPCOLESTR")]string pszPath) {
			return 0;
		}
	}

	class StubVsFontAndColorCacheManager : IVsFontAndColorCacheManager {
		public int CheckCache(ref Guid rguidCategory, out int pfHasData) {
			pfHasData = 0;
			return 0;
		}

		public int CheckCacheable(ref Guid rguidCategory, out int pfCacheable) {
			pfCacheable = 0;
			return 0;
		}

		public int ClearAllCaches() {
			return 0;
		}

		public int ClearCache(ref Guid rguidCategory) {
			return 0;
		}

		public int RefreshCache(ref Guid rguidCategory) {
			return 0;
		}
	}

	class StubVsMonitorSelection : IVsMonitorSelection {
		public int AdviseSelectionEvents(IVsSelectionEvents pSink, out uint pdwCookie) {
			pdwCookie = 0;
			return 0;
		}

		public int GetCmdUIContextCookie(ref Guid rguidCmdUI, out uint pdwCmdUICookie) {
			pdwCmdUICookie = 0;
			return 0;
		}

		public int GetCurrentElementValue(uint elementid, out object pvarValue) {
			pvarValue = null;
			return 0;
		}

		public int GetCurrentSelection(out IntPtr ppHier, out uint pitemid, out IVsMultiItemSelect ppMIS, out IntPtr ppSC) {
			pitemid = 0;
			ppHier = IntPtr.Zero;
			ppMIS = null;
			ppSC = IntPtr.Zero;
			return 0;
		}

		public int IsCmdUIContextActive(uint dwCmdUICookie, out int pfActive) {
			pfActive = 0;
			return 0;
		}

		public int SetCmdUIContext(uint dwCmdUICookie, int fActive) {
			return 0;
		}

		public int UnadviseSelectionEvents(uint dwCookie) {
			return 0;
		}
	}

	class StubVsRunningDocumentTable : IVsRunningDocumentTable {
		public int AdviseRunningDocTableEvents(IVsRunningDocTableEvents pSink, out uint pdwCookie) {
			pdwCookie = 42;
			return 0;
		}

		public int FindAndLockDocument(uint dwRDTLockType, string pszMkDocument, out IVsHierarchy ppHier, out uint pitemid, out IntPtr ppunkDocData, out uint pdwCookie) {
			ppHier = null;
			pitemid = 0;
			ppunkDocData = IntPtr.Zero;
			pdwCookie = 0;
			return 0;
		}

		public int GetDocumentInfo(uint docCookie, out uint pgrfRDTFlags, out uint pdwReadLocks, out uint pdwEditLocks, out string pbstrMkDocument, out IVsHierarchy ppHier, out uint pitemid, out IntPtr ppunkDocData) {
			throw new NotImplementedException();
		}

		public int GetRunningDocumentsEnum(out IEnumRunningDocuments ppenum) {
			throw new NotImplementedException();
		}

		public int LockDocument(uint grfRDTLockType, uint dwCookie) {
			throw new NotImplementedException();
		}

		public int ModifyDocumentFlags(uint docCookie, uint grfFlags, int fSet) {
			throw new NotImplementedException();
		}

		public int NotifyDocumentChanged(uint dwCookie, uint grfDocChanged) {
			throw new NotImplementedException();
		}

		public int NotifyOnAfterSave(uint dwCookie) {
			throw new NotImplementedException();
		}

		public int NotifyOnBeforeSave(uint dwCookie) {
			throw new NotImplementedException();
		}

		public int RegisterAndLockDocument(uint grfRDTLockType, string pszMkDocument, IVsHierarchy pHier, uint itemid, IntPtr punkDocData, out uint pdwCookie) {
			throw new NotImplementedException();
		}

		public int RegisterDocumentLockHolder(uint grfRDLH, uint dwCookie, IVsDocumentLockHolder pLockHolder, out uint pdwLHCookie) {
			throw new NotImplementedException();
		}

		public int RenameDocument(string pszMkDocumentOld, string pszMkDocumentNew, IntPtr pHier, uint itemidNew) {
			throw new NotImplementedException();
		}

		public int SaveDocuments(uint grfSaveOpts, IVsHierarchy pHier, uint itemid, uint docCookie) {
			throw new NotImplementedException();
		}

		public int UnadviseRunningDocTableEvents(uint dwCookie) {
			return 0;
		}

		public int UnlockDocument(uint grfRDTLockType, uint dwCookie) {
			throw new NotImplementedException();
		}

		public int UnregisterDocumentLockHolder(uint dwLHCookie) {
			throw new NotImplementedException();
		}
	}

	class StubVsSolultion : IVsSolution {
		public int AddVirtualProject(IVsHierarchy pHierarchy, uint grfAddVPFlags) {
			throw new NotImplementedException();
		}

		public int AddVirtualProjectEx(IVsHierarchy pHierarchy, uint grfAddVPFlags, ref Guid rguidProjectID) {
			throw new NotImplementedException();
		}

		public int AdviseSolutionEvents(IVsSolutionEvents pSink, out uint pdwCookie) {
			pdwCookie = 42;
			return 0;
		}

		public int CanCreateNewProjectAtLocation(int fCreateNewSolution, string pszFullProjectFilePath, out int pfCanCreate) {
			throw new NotImplementedException();
		}

		public int CloseSolutionElement(uint grfCloseOpts, IVsHierarchy pHier, uint docCookie) {
			throw new NotImplementedException();
		}

		public int CreateNewProjectViaDlg(string pszExpand, string pszSelect, uint dwReserved) {
			throw new NotImplementedException();
		}

		public int CreateProject(ref Guid rguidProjectType, string lpszMoniker, string lpszLocation, string lpszName, uint grfCreateFlags, ref Guid iidProject, out IntPtr ppProject) {
			throw new NotImplementedException();
		}

		public int CreateSolution(string lpszLocation, string lpszName, uint grfCreateFlags) {
			throw new NotImplementedException();
		}

		public int GenerateNextDefaultProjectName(string pszBaseName, string pszLocation, out string pbstrProjectName) {
			throw new NotImplementedException();
		}

		public int GenerateUniqueProjectName(string lpszRoot, out string pbstrProjectName) {
			throw new NotImplementedException();
		}

		public int GetGuidOfProject(IVsHierarchy pHierarchy, out Guid pguidProjectID) {
			throw new NotImplementedException();
		}

		public int GetItemInfoOfProjref(string pszProjref, int propid, out object pvar) {
			throw new NotImplementedException();
		}

		public int GetItemOfProjref(string pszProjref, out IVsHierarchy ppHierarchy, out uint pitemid, out string pbstrUpdatedProjref, VSUPDATEPROJREFREASON[] puprUpdateReason) {
			throw new NotImplementedException();
		}

		public int GetProjectEnum(uint grfEnumFlags, ref Guid rguidEnumOnlyThisType, out IEnumHierarchies ppenum) {
			throw new NotImplementedException();
		}

		public int GetProjectFactory(uint dwReserved, Guid[] pguidProjectType, string pszMkProject, out IVsProjectFactory ppProjectFactory) {
			throw new NotImplementedException();
		}

		public int GetProjectFilesInSolution(uint grfGetOpts, uint cProjects, string[] rgbstrProjectNames, out uint pcProjectsFetched) {
			throw new NotImplementedException();
		}

		public int GetProjectInfoOfProjref(string pszProjref, int propid, out object pvar) {
			throw new NotImplementedException();
		}

		public int GetProjectOfGuid(ref Guid rguidProjectID, out IVsHierarchy ppHierarchy) {
			throw new NotImplementedException();
		}

		public int GetProjectOfProjref(string pszProjref, out IVsHierarchy ppHierarchy, out string pbstrUpdatedProjref, VSUPDATEPROJREFREASON[] puprUpdateReason) {
			throw new NotImplementedException();
		}

		public int GetProjectOfUniqueName(string pszUniqueName, out IVsHierarchy ppHierarchy) {
			throw new NotImplementedException();
		}

		public int GetProjectTypeGuid(uint dwReserved, string pszMkProject, out Guid pguidProjectType) {
			throw new NotImplementedException();
		}

		public int GetProjrefOfItem(IVsHierarchy pHierarchy, uint itemid, out string pbstrProjref) {
			throw new NotImplementedException();
		}

		public int GetProjrefOfProject(IVsHierarchy pHierarchy, out string pbstrProjref) {
			throw new NotImplementedException();
		}

		public int GetProperty(int propid, out object pvar) {
			throw new NotImplementedException();
		}

		public int GetSolutionInfo(out string pbstrSolutionDirectory, out string pbstrSolutionFile, out string pbstrUserOptsFile) {
			throw new NotImplementedException();
		}

		public int GetUniqueNameOfProject(IVsHierarchy pHierarchy, out string pbstrUniqueName) {
			throw new NotImplementedException();
		}

		public int GetVirtualProjectFlags(IVsHierarchy pHierarchy, out uint pgrfAddVPFlags) {
			throw new NotImplementedException();
		}

		public int OnAfterRenameProject(IVsProject pProject, string pszMkOldName, string pszMkNewName, uint dwReserved) {
			throw new NotImplementedException();
		}

		public int OpenSolutionFile(uint grfOpenOpts, string pszFilename) {
			throw new NotImplementedException();
		}

		public int OpenSolutionViaDlg(string pszStartDirectory, int fDefaultToAllProjectsFilter) {
			throw new NotImplementedException();
		}

		public int QueryEditSolutionFile(out uint pdwEditResult) {
			throw new NotImplementedException();
		}

		public int QueryRenameProject(IVsProject pProject, string pszMkOldName, string pszMkNewName, uint dwReserved, out int pfRenameCanContinue) {
			throw new NotImplementedException();
		}

		public int RemoveVirtualProject(IVsHierarchy pHierarchy, uint grfRemoveVPFlags) {
			throw new NotImplementedException();
		}

		public int SaveSolutionElement(uint grfSaveOpts, IVsHierarchy pHier, uint docCookie) {
			throw new NotImplementedException();
		}

		public int SetProperty(int propid, object var) {
			throw new NotImplementedException();
		}

		public int UnadviseSolutionEvents(uint dwCookie) {
			return 0;
		}
	}

	class StubVsUIShellOpenDocument : IVsUIShellOpenDocument4 {
		[return: ComAliasName("VsShell.VSDOCINPROJECT")]
		public int IsDocumentInAProject2(string pszMkDocument, bool fSupportExternalItems, out IVsUIHierarchy ppUIH, out uint pitemid, out Microsoft.VisualStudio.OLE.Interop.IServiceProvider ppSP) {
			ppUIH = null;
			pitemid = 0;
			ppSP = null;
			return 0;
		}

		public IVsWindowFrame OpenDocumentViaProject2(string pszMkDocument, ref Guid rguidLogicalView, bool fSupportExternalItems, out Microsoft.VisualStudio.OLE.Interop.IServiceProvider ppSP, out IVsUIHierarchy ppHier, out uint pitemid) {
			throw new NotImplementedException();
		}
	}
}
