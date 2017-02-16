 
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using ExportPartData;
using System.Windows.Forms;
using System.IO;
using EPDM.Interop.epdm;

namespace AddinConvertTo.Classes
{
    public class Batches
    {
        #region Variables Batch
        static IEdmBatchGet batchGetter;
        static IEdmFile5 aFile;
        static IEdmFolder5 aFolder;
        static IEdmPos5 aPos;
        static IEdmSelectionList6 fileList = null;
        static IEdmBatchUnlock batchUnlocker;
        static IEdmBatchAdd poAdder;
        static EdmSelItem[] ppoSelection;
        static string fileNameErr = "";

        static public bool statusChank { get; set; }

        #endregion
        public static List<FilesData.TaskParam> BatchGetVariable(IEdmVault7 vault, List<FilesData.TaskParam> listType)
        {
            var Fulllist = new List<FilesData.TaskParam>();
            foreach (var fileVar in listType)
            {
                var filePath = fileVar.FolderPath + "\\" + fileVar.ConvertFile;
                var rev = CheckRev(vault, fileVar.TaskType, filePath);
                Fulllist.Add(new FilesData.TaskParam
                {
                    CurrentVersion = fileVar.CurrentVersion,
                    FileName = fileVar.FileName,
                    FolderPath = fileVar.FolderPath,
                    FullFilePath = fileVar.FullFilePath,
                    FolderID = fileVar.FolderID,
                    IdPDM = fileVar.IdPDM,
                    TaskType = fileVar.TaskType,
                    ConvertFile = fileVar.ConvertFile,
                    Revision = rev,
                    ExistEdrawingAndPDF = fileVar.CurrentVersion.ToString() == rev,
                    ExistCutList = true,
                    //ExistCutList = ExportXmlSql.ExistXml(fileVar.FileName.ToUpper().Replace(".SLDPRT", ""), fileVar.CurrentVersion),
                    ExistDXF = true
                    //ExistDXF = GetConfigNamePdm(vault, fileVar.IdPDM, fileVar.CurrentVersion)
                });
            }
            return Fulllist;
        }
        static string CheckRev(IEdmVault7 vault, int TaskType, string filePath)
        {
            var variable = "";
            try
            {
                var aFolder = default(IEdmFolder5);
                if (TaskType == 1)
                {
                    aFile = vault.GetFileFromPath(filePath, out aFolder);
                }
                if (TaskType == 2)
                {
                    aFile = vault.GetFileFromPath(filePath, out aFolder);
                }

                object oVarRevision;

                if (aFile == null)
                {
                    variable = "0";
                }
                else
                {
                    var pEnumVar = (IEdmEnumeratorVariable8)aFile.GetEnumeratorVariable();
                    pEnumVar.GetVar("Revision", "", out oVarRevision);
                    if (oVarRevision == null)
                    {
                        variable = "0";
                    }
                    else
                    {
                        variable = oVarRevision.ToString();
                    }
                }
                statusChank = true;
            }
            catch (COMException ex)
            {
                Logger.ToLog("BatchGet HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.ToString(), 10001);
                statusChank = false;
            }
            return variable;
        }
        public static bool GetConfigNamePdm(IEdmVault7 vault, int IdPDM, int CurrentVersion)
        {
            var result = true;
            //Get configurations
            var fileEdm = (IEdmFile5)vault.GetObject(EdmObjectType.EdmObject_File, IdPDM);
            EdmStrLst5 cfgList = default(EdmStrLst5);
            cfgList = fileEdm.GetConfigurations();
            string cfgName = null;
            IEdmPos5 pos = default(IEdmPos5);
            pos = cfgList.GetHeadPosition();
            while (!pos.IsNull)
            {
                cfgName = cfgList.GetNext(pos);
                if (cfgName != "@")
                {
                    Exception ex;
                    var existDxf = Dxf.ExistDxf(IdPDM, CurrentVersion, cfgName,out ex);
                    //Logger.ToLog("Проверка на DXF; \n IdPDM: " + IdPDM + "\n Version: " + CurrentVersion + "\n ConfigName: " + cfgName + "\n Dxf.ExistDxf: " + existDxf, 10003);
                    if (!existDxf)
                    {
                        return result = false;
                    }
                }
            }
            return result;
        }
        public static void ClearLocalCache(IEdmVault7 vault, List<FilesData.TaskParam> listType)
        {
            try
            {
                var ClearLocalCache = (IEdmClearLocalCache3)vault.CreateUtility(EdmUtility.EdmUtil_ClearLocalCache);
                ClearLocalCache.IgnoreToolboxFiles = true;

                //Declare and create the IEdmBatchListing object
                var BatchListing = (IEdmBatchListing)vault.CreateUtility(EdmUtility.EdmUtil_BatchList);

                foreach (var item in listType)
                {
                    ClearLocalCache.AddFileByPath(item.FullFilePath);
                    //((IEdmBatchListing2)BatchListing).AddFileCfg(KvPair.Key, DateTime.Now, (Convert.ToInt32(KvPair.Value)), "@", Convert.ToInt32(EdmListFileFlags.EdmList_Nothing));
                }

                //Clear the local cache of the reference files
                ClearLocalCache.CommitClear();
            }
            catch (COMException ex)
            {
                Logger.ToLog("ERROR ClearLocalCache файл: " + fileNameErr + " HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.ToString(), 10001);
            }
        }

        public static void BatchGet(IEdmVault7 vault, List<FilesData.TaskParam> listType)
        {
            try
            {
                batchGetter = (IEdmBatchGet)vault.CreateUtility(EdmUtility.EdmUtil_BatchGet);
                foreach (var taskVar in listType)
                {
                    //aFile = vault.GetFileFromPath(taskVar.FullFilePath, out ppoRetParentFolder);
                    //aFile =(IEdmFile5) vault.GetObject(EdmObjectType.EdmObject_File, taskVar.IdPDM);
                    //aPos = aFile.GetFirstFolderPosition();
                    //aFolder = aFile.GetNextFolder(aPos);
                    batchGetter.AddSelectionEx((EdmVault5)vault, taskVar.IdPDM, taskVar.FolderID, taskVar.CurrentVersion);
                }
                if ((batchGetter != null))
                {
                    //batchGetter.CreateTree(0, (int)EdmGetCmdFlags.Egcf_SkipExisting + (int)EdmGetCmdFlags.Egcf_SkipUnlockedWritable + (int)EdmGetCmdFlags.Egcf_RefreshFileListing);
                    batchGetter.CreateTree(0, (int)EdmGetCmdFlags.Egcf_SkipExisting + (int)EdmGetCmdFlags.Egcf_SkipUnlockedWritable);
                    batchGetter.GetFiles(0, null);
                }
                statusChank = true;
            }
            catch (COMException ex)
            {
                Logger.ToLog("ERROR BatchGet HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.ToString(), 10001);
                statusChank = false;
            }
        }
        public static void BatchDelete(IEdmVault7 vault, List<FilesData.TaskParam> listType)
        {
            EdmBatchDelErrInfo[] ppoDelErrors = null;
            try
            {
                var batchDeleter = (IEdmBatchDelete3)vault.CreateUtility(EdmUtility.EdmUtil_BatchDelete);

                foreach (var fileVar in listType)
                {
                    if (!fileVar.ExistEdrawingAndPDF)
                    {
                        var filePath = fileVar.FolderPath + "\\" + fileVar.ConvertFile;
                        // Add selected file to the batch
                        batchDeleter.AddFileByPath(filePath);
                    }
                }
                batchDeleter.ComputePermissions(true, null);
                var retVal = batchDeleter.CommitDelete(0, null);
                if ((!retVal))
                {
                    batchDeleter.GetCommitErrors(ppoDelErrors);
                }
                statusChank = true;
            }
            catch (COMException ex)
            {
                Logger.ToLog("ERROR BatchDelete - " + ppoDelErrors + " - HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.ToString(), 10001);
                statusChank = false;
            }
        }
        public static void BatchAddFiles(IEdmVault7 vault, List<FilesData.TaskParam> listType)
        {
            string msg = "";
            try
            {
                var result = default(bool);
                poAdder = (IEdmBatchAdd)vault.CreateUtility(EdmUtility.EdmUtil_BatchAdd);
                foreach (var fileName in listType)
                {
                    fileNameErr = fileName.FolderPath + "\\" + fileName.FileName;
                    poAdder.AddFileFromPathToPath(fileName.FolderPath + "\\" + fileName.ConvertFile, fileName.FolderPath, 0, "", 0);
                }
                var edmFileInfo = new EdmFileInfo[listType.Count];

                result = Convert.ToBoolean(poAdder.CommitAdd(0, edmFileInfo, 0));
                var idx = edmFileInfo.GetLowerBound(0);
                while (idx <= edmFileInfo.GetUpperBound(0))
                {
                    string row = null;
                    row = "(" + edmFileInfo[idx].mbsPath + ") arg = " + Convert.ToString(edmFileInfo[idx].mlArg);

                    if (edmFileInfo[idx].mhResult == 0)
                    {
                        row = row + " status = OK " + edmFileInfo[idx].mbsPath;
                    }
                    else
                    {
                        string oErrName = "";
                        string oErrDesc = "";

                        vault.GetErrorString(edmFileInfo[idx].mhResult, out oErrName, out oErrDesc);
                        row = row + " status = " + oErrName;
                    }

                    idx = idx + 1;
                    msg = msg + "\n" + row;
                }
                statusChank = true;
            }
            catch (COMException ex)
            {
                Logger.ToLog("ERROR BatchAddFiles " + msg + ", file: " + fileNameErr + " HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.ToString(), 10001);
                statusChank = false;
            }
        }
        public static void BatchSetVariable(IEdmVault7 vault, List<FilesData.TaskParam> listType)
        {
            try
            {
                foreach (var fileVar in listType)
                {
                    var filePath = fileVar.FolderPath + "\\" + fileVar.ConvertFile;
                    fileNameErr = filePath;
                    IEdmFolder5 folder;
                    aFile = vault.GetFileFromPath(filePath, out folder);
                    //var edmFile = (IEdmFile5)vault.GetObject(EdmObjectType.EdmObject_File, fileVar.IdPDM);

                    var pEnumVar = (IEdmEnumeratorVariable8)aFile.GetEnumeratorVariable(); ;
                    pEnumVar.SetVar("Revision", "", fileVar.CurrentVersion);
                }
                statusChank = true;
            }
            catch (COMException ex)
            {
                statusChank = false;
                Logger.ToLog("ERROR BatchSetVariable файл: " + fileNameErr + ", " + ex.ToString(), 10001);
            }
        }
        public static void BatchUnLock(IEdmVault7 vault, List<FilesData.TaskParam> listType)
        {
            try
            {
                ppoSelection = new EdmSelItem[listType.Count];
                batchUnlocker = (IEdmBatchUnlock)vault.CreateUtility(EdmUtility.EdmUtil_BatchUnlock);
                var i = 0;
                foreach (var fileName in listType)
                {
                    var filePath = fileName.FolderPath + "\\" + fileName.ConvertFile;
                    fileNameErr = filePath;
                    IEdmFolder5 folder = default(IEdmFolder5);
                    aFile = vault.GetFileFromPath(filePath, out folder);
                    aPos = aFile.GetFirstFolderPosition();
                    aFolder = aFile.GetNextFolder(aPos);

                    ppoSelection[i] = new EdmSelItem();
                    ppoSelection[i].mlDocID = aFile.ID;
                    ppoSelection[i].mlProjID = aFolder.ID;

                    i = i + 1;
                }
                // Add selections to the batch of files to check in
                batchUnlocker.AddSelection((EdmVault5)vault, ppoSelection);
                if ((batchUnlocker != null))
                {
                    batchUnlocker.CreateTree(0, (int)EdmUnlockBuildTreeFlags.Eubtf_ShowCloseAfterCheckinOption + (int)EdmUnlockBuildTreeFlags.Eubtf_MayUnlock);
                    fileList = (IEdmSelectionList6)batchUnlocker.GetFileList((int)EdmUnlockFileListFlag.Euflf_GetUnlocked + (int)EdmUnlockFileListFlag.Euflf_GetUndoLocked + (int)EdmUnlockFileListFlag.Euflf_GetUnprocessed);
                    batchUnlocker.UnlockFiles(0, null);
                }
                statusChank = true;
            }
            catch (COMException ex)
            {
                Logger.ToLog("ERROR BatchUnLock файл: '" + fileNameErr + "', " + ex.StackTrace + " " + ex.ToString(), 10001);
                statusChank = false;
            }
        }
    }
}