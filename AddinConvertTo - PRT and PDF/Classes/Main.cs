using System;
using System.Linq; 
using System.Windows.Forms;
using System.Diagnostics;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Runtime.InteropServices;
using AddinConvertTo.Classes;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using AddinConvertTo;
using ExportPartData;
using System.Collections;
using SolidWorks.Interop.swdocumentmgr;
using EPDM.Interop.epdm;

namespace NewAddinToPDM
{
    public class ClassHome : IEdmAddIn5
    {
        #region Variables
            SldWorks swApp;
            Process[] processes;
            IEdmTaskInstance edmTaskInstance;
            int _currentVer, errors, warnings, countItems;
            bool statusTask = true, statusChunks; // Статус на ошибки
            int checkNumberOfChunk; // Номер пачки          
            Exception ex;
            List<FilesData.TaskParam> filesPdm { get; set; }
            List<FilesData.TaskParam> BatchGetVariable { get; set; }
        #endregion
        #region EdmAddIn
            public void GetAddInInfo(ref EdmAddInInfo poInfo, IEdmVault5 poVault, IEdmCmdMgr5 poCmdMgr)
            {
                try
                {
                    #region Addin version
                        //// Addin version
                        //poInfo.mbsAddInName = "C# Task Add-In";
                        //poInfo.mbsCompany = "Vents";
                        //poInfo.mbsDescription = "PDF & EPRT";
                        //poInfo.mlAddInVersion = 1;
                        //poInfo.mlRequiredVersionMajor = 6;
                        //poInfo.mlRequiredVersionMinor = 4;

                        //poCmdMgr.AddCmd(1, "Convert", (int)EdmMenuFlags.EdmMenu_Nothing, "", "", 0, 0);
                    #endregion
                    #region Task version
                        // Task version
                        const int ver = 8;
                        poInfo.mbsAddInName = "Make .eprt and .pdf files Task Add-In ver." + ver;
                        poInfo.mbsCompany = "Vents";
                        poInfo.mbsDescription = "Создание и сохранение файлов .pdf и .eprt для листовых деталей";
                        poInfo.mlAddInVersion = ver;
                        _currentVer = poInfo.mlAddInVersion;

                        //Minimum SolidWorks Enterprise PDM version
                        //needed for C# Task Add-Ins is 10.0
                        poInfo.mlRequiredVersionMajor = 10;
                        poInfo.mlRequiredVersionMinor = 0;

                        //Register this add-in as a task add-in
                        poCmdMgr.AddHook(EdmCmdType.EdmCmd_TaskRun);
                        //Register this add-in to be called when
                        //selected as a task in the Administration tool
                        poCmdMgr.AddHook(EdmCmdType.EdmCmd_TaskSetup);
                        poCmdMgr.AddHook(EdmCmdType.EdmCmd_TaskSetupButton);
                        poCmdMgr.AddHook(EdmCmdType.EdmCmd_TaskDetails);
                    #endregion
                }
                catch (COMException ex)
                {
                    Logger.ToLog("GetAddInInfo HRESULT = 0x" + ex.ErrorCode.ToString("X") + "; " + ex.ToString(), 10001);
                }
                catch (Exception ex)
                {
                    Logger.ToLog(ex.ToString() + "; \n" + ex.StackTrace, 1000);
                }
            }
            public void OnCmd(ref EdmCmd poCmd, ref Array ppoData)
            {
                try
                {
                    #region Task Addin version
                    //PauseToAttachProcess(poCmd.meCmdType.ToString());
                    switch (poCmd.meCmdType)
                    {
                        case EdmCmdType.EdmCmd_TaskRun:
                            OnTaskRun(ref poCmd, ref ppoData);
                            break;
                        case EdmCmdType.EdmCmd_TaskSetup:
                            OnTaskSetup(ref poCmd, ref ppoData);
                            break;
                        case EdmCmdType.EdmCmd_TaskDetails:
                            break;
                        case EdmCmdType.EdmCmd_TaskSetupButton:
                            break;
                    }
                    #endregion
                    #region Addin version
                    //vault = poCmd.mpoVault as IEdmVault7;
                    //switch (poCmd.meCmdType)
                    //        {
                    //            case EdmCmdType.EdmCmd_Menu:
                    //                OnTaskRun(ref poCmd, ref ppoData);
                    //                break;
                    //        }
                    #endregion
                }
                catch (COMException ex)
                {
                    Logger.ToLog("OnCmd HRESULT = 0x" + ex.ErrorCode.ToString("X") + "; " + ex.ToString(), 10001);

                }
                catch (Exception ex)
                {
                    Logger.ToLog("OnCmd = " + ex.ToString() + "; " + ex.StackTrace, 1000);
                }
            }
        #endregion
        // 1000 код ошибки для Exception
        // 10001 код ошибки для ComException
        // 1001 код ошибки для SwModel
        // 1002 код на старте задачи
        #region Task Addin
            static void OnTaskSetup(ref EdmCmd poCmd, ref Array ppoData)
            {
                try
                {
                    //Get the property interface used to
                    //access the framework
                    var edmTaskProperties = (IEdmTaskProperties)poCmd.mpoExtra;
                    //Set the property flag that says you want a
                    //menu item for the user to launch the task
                    //and a flag to support scheduling
                    edmTaskProperties.TaskFlags = (int)EdmTaskFlag.EdmTask_SupportsInitForm + (int)EdmTaskFlag.EdmTask_SupportsDetails + (int)EdmTaskFlag.EdmTask_SupportsChangeState;
                    //edmTaskProperties.TaskFlags = (int)EdmTaskFlag.EdmTask_SupportsChangeState + (int)EdmTaskFlag.EdmTask_SupportsInitExec;
                    //Set up the menu commands to launch this task
                    var edmTaskMenuCmds = new EdmTaskMenuCmd[1];
                    edmTaskMenuCmds[0].mbsMenuString = "Выгрузить"; //"Выгрузка PDF и Edrawing для листовых деталей"
                    edmTaskMenuCmds[0].mbsStatusBarHelp = "Выгрузить"; //"Выгрузка PDF и Edrawing для листовых деталей"
                    edmTaskMenuCmds[0].mlCmdID = 1;
                    edmTaskMenuCmds[0].mlEdmMenuFlags = (int)EdmMenuFlags.EdmMenu_Nothing;
                    edmTaskProperties.SetMenuCmds(edmTaskMenuCmds);
                }
                catch (COMException ex)
                {
                    Logger.ToLog("OnTaskSetup HRESULT = 0x" + ex.ErrorCode.ToString("X") + "; " + ex.ToString(), 10001);
                }
                catch (Exception ex)
                {
                    Logger.ToLog("OnTaskSetup Error" + ex.ToString() + ppoData, 1000);
                }
            }
            void OnTaskRun(ref EdmCmd poCmd, ref Array ppoData)
            {
                List<KeyValuePair<int, bool>> listChunksEx = new List<KeyValuePair<int, bool>>();
                edmTaskInstance = (IEdmTaskInstance)poCmd.mpoExtra;
                edmTaskInstance.SetStatus(EdmTaskStatus.EdmTaskStat_Running);
                try
                {
                    var vault = poCmd.mpoVault as IEdmVault7;
                    countItems = 1;

                    #region Выгрузка XML из Excell спецификации
                    //var filesExcell = FilesData.ExcFiles(vault, ref poCmd, ref ppoData);
                    //if (filesExcell.Count() > 0)
                    //{ 
                    //    Batches.BatchGet(vault, filesExcell);
                    //    Exception exOut = null;
                    //    foreach (var item in filesExcell.Where(x => Path.GetExtension(x.FullFilePath).Contains("xl")))
                    //    {
                    //        Logger.ToLog($"EXCEL: { item.FullFilePath}", 10003);
                    //        SpecFromExcel.Spec.ExportToXml(item.FullFilePath, out exOut);
                    //    }
                    //    if (exOut != null)
                    //    {
                    //        Logger.ToLog(exOut.Message + "; " + exOut.StackTrace, 1000);
                    //        statusTask = false;
                    //    }
                    //}
                    #endregion

                    filesPdm = FilesData.SolidWorksFiles(vault, ref poCmd, ref ppoData);

                    BatchGetVariable = Batches.BatchGetVariable(vault, filesPdm);

                    //ExistEdrawingAndPDF false если версии не совпадают
                    //var sortBatchGetVariable = BatchGetVariable.Where(x => (x.TaskType == 1 && (!x.ExistCutList || !x.ExistDXF || !x.ExistEdrawingAndPDF)) ||
                    //   (x.TaskType == 2 && !x.ExistEdrawingAndPDF)).OrderBy(y => y.TaskType).ToList();
                    var sortBatchGetVariable = BatchGetVariable.Where(x => !x.ExistEdrawingAndPDF).OrderBy(y => y.TaskType).ToList();

                    //var frmTest = new FormTest() { listForm = BatchGetVariable };
                    //frmTest.ShowDialog();
                    //return;

                    #region Log
                    Logger.ToLog("", 123);
                    Logger.ToLog("", 123);
                    Logger.ToLog(new string('-', 500), 123);
                    Logger.ToLog("", 123);
                    Logger.ToLog("", 123);
                    Logger.ToLog($"Выгрузка сборки в XML", 10003);
                    Logger.ToLog($"Задача '{edmTaskInstance.TaskName}' ID: {edmTaskInstance.ID}, для {sortBatchGetVariable.Count} элемента(ов), OnTaskRun ver.{_currentVer}", 1002);
                    Logger.ToLog("Получаем перечень файлов для конвертации: ", 10003);
                    //Logger.ToLog(new string('-', 50), 10003);
                    //foreach (var item in sortBatchGetVariable)
                    //{
                    //    var extension = Path.GetExtension(item.FileName);
                    //    switch (extension.ToLower())
                    //    {
                    //        case ".sldprt":
                    //            Logger.ToLog(countItems++ + " Файл: " + item.FileName + ", ID: " + item.IdPDM + ", CurVer: " + item.CurrentVersion +
                    //                ", Совпадают ли версии = " + item.ExistEdrawingAndPDF + ", Есть ли CutList = " + item.ExistCutList + ", Есть ли DXF = " + item.ExistDXF +
                    //                //", Листовой метал: " + item.ExistSheetMetal +
                    //                ", TaskType = " + item.TaskType, 10003);
                    //            break;
                    //        case ".slddrw":
                    //            Logger.ToLog(countItems++ + " Файл: " + item.FileName + ", ID: " + item.IdPDM + ", CurVer: " + item.CurrentVersion +
                    //                ", Совпадают ли версии = " + item.ExistEdrawingAndPDF + ", TaskType = " + item.TaskType, 10003);
                    //            break;
                    //    }
                    //}
                    //foreach (var item in filesExcell)
                    //{
                    //    Logger.ToLog(countItems++ + " Файл: " + item.FullFilePath + ", ID: " + item.IdPDM.ToString(), 10003);
                    //}
                    //Logger.ToLog(new string('-', 50), 10003);
                    Logger.ToLog("Файлы получены.", 10003);
                    #endregion
                    
                    edmTaskInstance.SetProgressRange(sortBatchGetVariable.Count, 0, "Запуск задачи.");

                    Logger.ToLog("Очистка локального кэша", 10003);
                    Batches.ClearLocalCache(vault, sortBatchGetVariable);
                    //return;
                    Logger.ToLog("Выполняется batch get", 10003);
                    Batches.BatchGet(vault, sortBatchGetVariable);
                    Logger.ToLog("Выполняется batch delete", 10003);
                    Batches.BatchDelete(vault, sortBatchGetVariable);
                    Logger.ToLog("Запускаем SldWorks и выполняем конвертацию", 10003);
                    // Run SldWorks
                    if (sortBatchGetVariable.Count != 0)
                    {
                        processes = Process.GetProcessesByName("SLDWORKS");
                        foreach (var process in processes)
                        {
                            process.Kill();
                        }
                        swApp = new SldWorks() { Visible = true };

                        // Разбиваем лист на группы
                        var nChunks = 10;
                        var totalLength = sortBatchGetVariable.Count();
                        var chunkLength = (int)Math.Ceiling(totalLength / (double)nChunks);
                        var partsToList = Enumerable.Range(0, chunkLength).Select(i => sortBatchGetVariable.Skip(i * 10).Take(10).ToList()).ToList();
                        Logger.ToLog("Пачек: " + partsToList.Count(), 123);
                        checkNumberOfChunk = 1;
                        
                        Exception ex = null;
                        for (var i = 0; i < partsToList.Count; i++)
                        {
                            Chunks(vault, partsToList[i], out ex);
                            listChunksEx.Add(new KeyValuePair<int, bool>(i, ex == null ? true : false));
                            if (ex != null)
                            { Logger.ToLog(ex.ToString() + "; " + ex.StackTrace, 1000); }
                        }
                    }
                }
                catch (COMException ex)
                {
                    statusTask = false;
                    this.ex = ex;
                    Logger.ToLog("OnTaskRun HRESULT = 0x" + ex.ErrorCode.ToString("X") + "; " + ex.StackTrace + "; \n" + ex.StackTrace, 10001);
                }
                catch (Exception ex)
                {
                    statusTask = false;
                    this.ex = ex;
                    Logger.ToLog(ex.ToString() + "; " + ex.StackTrace, 1000);
                }
                finally
                {
                    Logger.ToLog("~~~ FINALLY STATUS: " + statusTask + ", Пачек с ошибками: " + listChunksEx.Where(x => x.Value == false).Count(), 123);
                    if (listChunksEx.Where(x => x.Value == false).Count() == 0 & statusTask == true)
                    {
                        edmTaskInstance.SetProgressPos(BatchGetVariable.Count, "Выполнено, ID = " + edmTaskInstance.ID);
                        edmTaskInstance.SetStatus(EdmTaskStatus.EdmTaskStat_DoneOK, GetHashCode());
                    }
                    else
                    {
                        edmTaskInstance.SetProgressPos(BatchGetVariable.Count, "Ошибка, ID = " + edmTaskInstance.ID);
                        edmTaskInstance.SetStatus(EdmTaskStatus.EdmTaskStat_DoneFailed, GetHashCode());
                    }
                }
            }
        #endregion
       
        #region METHODS
            #region CHUNKS
            // Разбиваем массив файлов на пачки
            public void Chunks(IEdmVault7 vault, List<FilesData.TaskParam> chunks, out Exception exChunk)
            {
                exChunk = null;
                try
                {
                    Logger.ToLog("", 123);
                    Logger.ToLog("ПАЧКА №: " + checkNumberOfChunk++, 123);
                    Logger.ToLog("Запускаем SldWorks и выполняем конвертацию", 10003);
                    var ForBatchAdd = Convert(swApp, chunks);
                    Logger.ToLog($"ForBatchAdd.Count {ForBatchAdd.Count}", 123);
                    if (ForBatchAdd.Count > 0)
                    {
                        Logger.ToLog("Регистрация", 123);
                        Batches.BatchAddFiles(vault, ForBatchAdd);
                        Batches.BatchSetVariable(vault, ForBatchAdd);
                        Batches.BatchUnLock(vault, ForBatchAdd);
                    }
                    statusChunks = Batches.statusChank;
                }
                catch (Exception ex)
                {
                    exChunk = ex;
                    statusChunks = false;
                    statusTask = false;
                    Logger.ToLog("Convert: " + exChunk.Message + "; " + exChunk.StackTrace, 10001);
                }
                finally
                {
                    Logger.ToLog("Статус пачки: " + statusChunks, 123);
                }
            }
            #endregion
            public List<FilesData.TaskParam> Convert(SldWorks swApp, List<FilesData.TaskParam> OpenDocList)
            {
                var ListForBatchAdd = new List<FilesData.TaskParam>();
                var fileNameErr = "";
                foreach (var item in OpenDocList)
                {
                    fileNameErr = item.FullFilePath;
                    var status = true;
                    edmTaskInstance.SetProgressPos(countItems++, item.FullFilePath);
                    #region SLDPRT
                    if (item.TaskType == 1) // Только для листового металла
                    {
                        var swModel = swApp.OpenDoc6(item.FullFilePath, (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", 0, 0);
                        //Logger.ToLog(item.FullFilePath, errors); // Лог ошибки OpenDoc6
                        if (!IsSheetMetalPart(swModel)) { status = false; } // Проверка на листовой метал
                        // Запуск конвертации
                        #region Запуск конвертации
                        if (status) // Если листовой метал
                        {
                            Logger.ToLog($"~~~~~~ SLDPRT TaskType: {item.TaskType}, FullFilePath: {item.FullFilePath}, FileName: { item.FileName}", 10003);
                            Exception exOut = null;
                            List<Exception> exOutList = null;
                            List<Dxf.DxfFile> listDxf = new List<Dxf.DxfFile>();
                            List<ExportXmlSql.DataToExport> listCutList = new List<ExportXmlSql.DataToExport>();
                            var list = new List<Bends.SolidWorksFixPattern.PartBendInfo>();
                            var confArray = (object[])swModel.GetConfigurationNames();
                            var StatusFixbend = false;
                            #region ExistCutList
                            if (!item.ExistCutList)
                            {
                                //foreach (var confName in confArray) // Проходимся по всем конфигурация для Fix, CutList
                                //{
                                //    Configuration swConf = swModel.GetConfigurationByName(confName.ToString());
                                //    if (swConf.IsDerived()) continue;
                                //    swModel.ShowConfiguration2(confName.ToString());
                                //    if (!item.ExistCutList) // Проверка на CutList: false - XML отсутствует, true - XML есть
                                //    {
                                //        Bends.Fix(swApp, out list, false);
                                //        StatusFixbend = true;
                                //        List<ExportXmlSql.DataToExport> listCutListConf;
                                //        ExportXmlSql.GetCurrentConfigPartData(swApp, item.CurrentVersion, item.IdPDM, false, false, out listCutListConf, out exOut);
                                //        listCutList.AddRange(listCutListConf);
                                //    }
                                //    if (exOut != null)
                                //    {
                                //        Logger.ToLog(exOut.Message + "; " + exOut.StackTrace, 10001);
                                //        statusChunks = false;
                                //    }
                                //}
                                ////CutList To Sql
                                //if (listCutList != null)
                                //{
                                //    ExportXmlSql.ExportDataToXmlSql(item.FileName.ToUpper().Replace(".SLDPRT", ""), listCutList, out exOut);
                                //    if(exOut != null)
                                //    {
                                //        Logger.ToLog(exOut.Message + "; " + exOut.StackTrace, 10001);
                                //        statusChunks = false;
                                //    }
                                //}
                            }
                            #endregion
                            #region ExistDXF
                            if (!item.ExistDXF)
                            {
                                //foreach (var confName in confArray) // Проходимся по всем конфигурация для Fix, DXF
                                //{
                                //    //Configuration swConf = swModel.GetConfigurationByName(confName.ToString());
                                //    //if (swConf.IsDerived()) continue;
                                //    if (!item.ExistDXF) // Проверка на Dxf
                                //    {
                                //        if (!Dxf.ExistDxf(item.IdPDM, item.CurrentVersion, confName.ToString()))
                                //        {
                                //            swModel.ShowConfiguration2(confName.ToString());
                                //            if (!StatusFixbend)
                                //            {
                                //                Bends.Fix(swApp, out list, false);
                                //            }
                                //            if (Dxf.Save(swApp, out exOut, item.IdPDM, item.CurrentVersion, out listDxf, false, false, confName.ToString()))
                                //            {
                                //                if (exOut != null)
                                //                {
                                //                    Logger.ToLog(exOut.Message + "; " + exOut.StackTrace, 10001);
                                //                }
                                //                Dxf.AddToSql(listDxf, true, out exOutList);
                                //                if (exOutList != null)
                                //                {
                                //                    foreach (var itemEx in exOutList)
                                //                    {
                                //                        Logger.ToLog($"DXF AddToSql err: {itemEx.Message}", 10001);
                                //                    }
                                //                    statusChunks = false;
                                //                }
                                //            }
                                //        }
                                //    }
                                //}
                            }
                            #endregion
                            #region ExistEdrawingAndPDF
                            if (!item.ExistEdrawingAndPDF)
                            {
                                try
                                {
                                    status = ConvertToErpt(swApp, swModel, item.FullFilePath) ? true : false;
                                    if (status)
                                    {
                                        ListForBatchAdd.Add(new FilesData.TaskParam
                                        {
                                            CurrentVersion = item.CurrentVersion,
                                            FileName = item.FileName,
                                            FolderPath = item.FolderPath,
                                            FullFilePath = item.FullFilePath,
                                            FolderID = item.FolderID,
                                            TaskType = item.TaskType,
                                            ConvertFile = item.ConvertFile,
                                            Revision = item.Revision
                                        });
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Logger.ToLog(ex.ToString() + "; " + ex.StackTrace, 10001);
                                    statusChunks = false;
                                }
                            }
                            #endregion
                           
                            Logger.ToLog($"CloseDoc: {item.FileName}", 123);
                        }
                        #endregion
                        swApp.CloseDoc(item.FullFilePath);
                    }
                    #endregion
                    #region SLDDRW
                    if (item.TaskType == 2)
                    {
                        try
                        {
                            Logger.ToLog($"~~~~~~ SLDDRW TaskType: {item.TaskType}, FullFilePath: {item.FullFilePath}, FileName: {item.FileName}", 10003);
                            var swModel = swApp.OpenDoc6(item.FullFilePath, (int)swDocumentTypes_e.swDocDRAWING, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", 0, 0);
                            //Logger.ToLog(item.FullFilePath, errors);
                            status = ConvertToPdf(swApp, swModel, item.FullFilePath) ? true : false; 

                            if (status)
                            {
                                ListForBatchAdd.Add(new FilesData.TaskParam
                                {
                                    CurrentVersion = item.CurrentVersion,
                                    FileName = item.FileName,
                                    FolderPath = item.FolderPath,
                                    FullFilePath = item.FullFilePath,
                                    FolderID = item.FolderID,
                                    TaskType = item.TaskType,
                                    ConvertFile = item.ConvertFile,
                                    Revision = item.Revision
                                });
                            }
                            swApp.CloseDoc(item.FullFilePath);
                        }
                        catch (Exception ex)
                        {
                            Logger.ToLog(ex.ToString() + ", \n" + ex.StackTrace, 123);
                            statusChunks = false;
                        }
                    }
                    #endregion
                }
                return ListForBatchAdd;
            }
            #region CONVERT TO PDF
                static public DispatchWrapper[] ObjectArrayToDispatchWrapperArray(object[] Objects)
                {
                    var ArraySize = 0;
                    ArraySize = Objects.GetUpperBound(0);
                    var d = new DispatchWrapper[ArraySize + 1];
                    var ArrayIndex = 0;
                    for (ArrayIndex = 0; ArrayIndex <= ArraySize; ArrayIndex++)
                    {
                        d[ArrayIndex] = new DispatchWrapper(Objects[ArrayIndex]);
                    }
                    return d;
                }
                bool ConvertToPdf(SldWorks swApp, ModelDoc2 swModel, string filePath)
                {
                    var result = false;
                    try
                    {
                        var swDraw = (DrawingDoc)swModel;
                        //var swExtension = swModel.Extension;
                        var obj = (string[])swDraw.GetSheetNames();
                        var count = obj.Length;
                        var objs = new object[count];
                        for (var i = 0; i <= count - 1; i++)
                        {
                            swDraw.ActivateSheet((obj[i]));
                            var swSheet = (Sheet)swDraw.GetCurrentSheet();
                            objs[i] = swSheet;
                        }
                        var swExportPDFData = (ExportPdfData)swApp.GetExportFileData((int)swExportDataFileType_e.swExportPdfData);
                        var dispWrapArr = (DispatchWrapper[])ObjectArrayToDispatchWrapperArray((objs));
                        swExportPDFData.SetSheets((int)swExportDataSheetsToExport_e.swExportData_ExportSpecifiedSheets, (dispWrapArr));

                        var fileParthPdf = filePath.ToUpper().Replace("SLDDRW", "PDF");
                        swModel.Extension.SaveAs(fileParthPdf, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent + (int)swSaveAsOptions_e.swSaveAsOptions_UpdateInactiveViews, swExportPDFData, 0, 0);
                        //Logger.ToLog("SAVE TO PDF + " + fileParthPdf, 10003);
                        //Logger.ToLog(filePath, errors);
                        result = true;
                    }
                    catch (Exception ex)
                    {
                        Logger.ToLog("CONVERT TO PDF + " + ex.ToString() + "; " + ex.StackTrace, 1000);
                        statusChunks = false;
                    }
                    return result;
                }
            #endregion
            #region CONVERT TO EDRAWING
                static bool IsSheetMetalPart(ModelDoc2 swModel)
                {
                    var isSheet = false;
                    try
                    {
                        var swPart = (PartDoc)swModel;
                        if (swPart != null)
                        {
                            var vBodies = swPart.GetBodies2((int)swBodyType_e.swSolidBody, false);
                            foreach (Body2 vBody in vBodies)
                            {
                                var isSheetMetal = vBody.IsSheetMetal();
                                if (!isSheetMetal) continue;
                                isSheet = true;
                            }
                        }
                        //Logger.ToLog($"IsSheetMetalPart: {isSheet}", 123);
                    }
                    catch (Exception)
                    {
                        isSheet = false;
                    }
                    return isSheet;
                }
                bool ConvertToErpt(SldWorks swApp, ModelDoc2 swModel, string filePath)
                {
                    var result = false;
                    try
                    {
                        swApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swEdrawingsSaveAsSelectionOption, (int)swEdrawingSaveAsOption_e.swEdrawingSaveAll);
                        swApp.SetUserPreferenceToggle(((int)(swUserPreferenceToggle_e.swEDrawingsOkayToMeasure)), true);
                        var fileParthEprt = filePath.ToUpper().Replace("SLDPRT", "EPRT");
                        swModel.Extension.SaveAs(fileParthEprt, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, 0, 0);
                        //Logger.ToLog(filePath, errors);
                        result = true;
                    }
                    catch (Exception ex)
                    {
                        Logger.ToLog("ERROR CONVERT TO ERPT + " + ex.ToString() + "; " + ex.StackTrace, 1000);
                        statusChunks = false;
                    }
                    return result;
                }

        public void OnCmd(ref EdmCmd poCmd, ref EdmCmdData[] ppoData)
        {
            throw new NotImplementedException();
        }
        #endregion
        #endregion
    }
}