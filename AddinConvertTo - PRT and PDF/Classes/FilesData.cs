
using EPDM.Interop.epdm;
using PdmAsmBomToXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

namespace AddinConvertTo.Classes
{
    public class FilesData
    {
        public class TaskParam
        {
            public string FileName { get; set; }
            public string FolderPath { get; set; }
            public int FolderID { get; set; }
            public int IdPDM { get; set; }
            public int CurrentVersion { get; set; }
            public string FullFilePath { get; set; }
            public int TaskType { get; set; }
            public string ConvertFile { get; set; }
            public string Revision { get; set; }
            public bool ExistDXF { get; set; }
            public bool ExistCutList { get; set; }
            public bool ExistEdrawingAndPDF { get; set; }
        }

        static public List<TaskParam> FilesPdm(IEdmVault5 vault, ref EdmCmd poCmd, ref Array ppoData)
        {
            var list = new List<TaskParam>();
            try
            {
                foreach (EdmCmdData edmCmdData in ppoData)
                {
                    var fileId = edmCmdData.mlObjectID1;
                    var parentFolderId = edmCmdData.mlObjectID2; // for Task
                    //var parentFolderId = edmCmdData.mlObjectID3; // for Addin
                    var fileEdm = (IEdmFile5)vault.GetObject(EdmObjectType.EdmObject_File, fileId);
                    var folder = (IEdmFolder5)vault.GetObject(EdmObjectType.EdmObject_Folder, parentFolderId);
                    var taskParam = new TaskParam();
                    taskParam.IdPDM = fileEdm.ID;
                    taskParam.CurrentVersion = fileEdm.CurrentVersion;
                    taskParam.FileName = fileEdm.Name;
                    taskParam.FolderPath = folder.LocalPath;
                    taskParam.FolderID = folder.ID;
                    taskParam.FullFilePath = folder.LocalPath + "\\" + fileEdm.Name;
                    list.Add(taskParam);
                }
            }
            catch (COMException ex)
            {
                Logger.ToLog("OnTaskSetup HRESULT = 0x" + ex.ErrorCode.ToString("X") + ex.ToString(), 10001);
            }
            catch (Exception ex)
            {
                Logger.ToLog(ex.ToString() + "; " + ex.StackTrace, 1000);
            }
            return list;
        }

        static public List<TaskParam> SolidWorksFiles(IEdmVault5 vault, ref EdmCmd poCmd, ref Array ppoData)
        {
            var list = new List<TaskParam>();

            foreach (var item in FilesPdm(vault, ref poCmd, ref ppoData))
            {
                var taskParam = new TaskParam();
                taskParam.IdPDM = item.IdPDM;
                var extension = Path.GetExtension(item.FileName);
                switch (extension.ToLower())
                {
                    case ".sldasm":
                        List<Exception> exOut = null;
                        AsmBomToXml.ВыгрузкаСборкиВXml(item.FullFilePath, out exOut);
                        if (exOut != null)
                        {
                            foreach (var itemEx in exOut)
                            {
                                Logger.ToLog("Error ВыгрузкаСборкиВXml: " + itemEx.Message, 10003);
                            }
                        }
                        break;

                    case ".sldprt":
                        //var existSheetMetal = new SheetMetal(item.FullFilePath); // Проверка на листовой метал DocumentManagement
                        //var resultMetal = existSheetMetal.ExistSheetMetalDM();
                        //if(resultMetal)
                        //{
                        taskParam.CurrentVersion = item.CurrentVersion;
                        taskParam.FileName = item.FileName;
                        taskParam.FolderPath = item.FolderPath;
                        taskParam.FolderID = item.FolderID;
                        taskParam.FullFilePath = item.FullFilePath;
                        taskParam.ConvertFile = item.FileName.ToUpper().Replace(".SLDPRT", ".EPRT");
                        taskParam.TaskType = 1;
                        list.Add(taskParam);
                        //}
                        break;

                    case ".slddrw":
                        taskParam.CurrentVersion = item.CurrentVersion;
                        taskParam.FileName = item.FileName;
                        taskParam.FolderPath = item.FolderPath;
                        taskParam.FolderID = item.FolderID;
                        taskParam.FullFilePath = item.FullFilePath;
                        taskParam.ConvertFile = item.FileName.ToUpper().Replace(".SLDDRW", ".PDF");
                        taskParam.TaskType = 2;
                        list.Add(taskParam);
                        break;
                }
            }
            return list;
        }

        static public List<TaskParam> ExcFiles(IEdmVault5 vault, ref EdmCmd poCmd, ref Array ppoData)
        {
            var list = new List<TaskParam>();
            foreach (var item in FilesPdm(vault, ref poCmd, ref ppoData))
            {
                var taskParam = new TaskParam();
                taskParam.IdPDM = item.IdPDM;
                var extension = Path.GetExtension(item.FileName);
                switch (extension.ToLower())
                {
                    case ".xlsx":
                        taskParam.FullFilePath = item.FullFilePath;
                        list.Add(taskParam);
                        break;

                    case ".xls":
                        taskParam.FullFilePath = item.FullFilePath;
                        list.Add(taskParam);
                        break;
                }
            }
            return list;
        }
    }
}