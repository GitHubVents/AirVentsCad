using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using FixBends;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using Environment = System.Environment;
using View = SolidWorks.Interop.sldworks.View;
using EPDM.Interop.epdm;


// TODO Проверка языка в Cutlist - с каким языком запустился SolidWorks


namespace MakeDxfUpdatePartData
{
    public partial class MakeDxfExportPartDataClass
    {
        static class LoggerMine
        {
            private const string ConnectionString = "Data Source=192.168.14.11;Initial Catalog=SWPlusDB;User ID=sa;Password=PDMadmin";

            private const string ClassName = "MakeDxfExportPartDataClass";
            
            public static void Error(string message, string код, string функция)
            {
                #region
                //using (var streamWriter = new StreamWriter("C:\\log.txt", true))
                //{
                //    streamWriter.WriteLine("{0,-20}  {2,-7} {3,-20} {1}", DateTime.Now + ":", message, "Error", ClassName);
                //}
                #endregion
                WriteToBase(message, "Error", код, ClassName, функция);
            }

            public static void Info(string message, string код, string функция)
            {
                #region
                //using (var streamWriter = new StreamWriter("C:\\log.txt", true))
                //{
                //    streamWriter.WriteLine("{0,-20}  {2,-7} {3,-20} {1}", DateTime.Now + ":", message, "Info", ClassName);
                //}
                #endregion
                WriteToBase(message, "Info", код, ClassName, функция);
            }

            static void WriteToBase(string описание, string тип, string код, string модуль, string функция)
            {
                using (var con = new SqlConnection(ConnectionString))
                {
                    try
                    {
                        con.Open();
                        var sqlCommand = new SqlCommand("AddErrorLog", con) { CommandType = CommandType.StoredProcedure };

                        var sqlParameter = sqlCommand.Parameters;

                        sqlParameter.AddWithValue("@UserName", Environment.UserName + " (" + System.Net.Dns.GetHostName() + ")");
                        sqlParameter.AddWithValue("@ErrorModule", модуль);
                        sqlParameter.AddWithValue("@ErrorMessage", описание);
                        sqlParameter.AddWithValue("@ErrorCode", код);
                        sqlParameter.AddWithValue("@ErrorTime", DateTime.Now);
                        sqlParameter.AddWithValue("@ErrorState", тип);
                        sqlParameter.AddWithValue("@ErrorFunction", функция);

                        #region
                        //var returnParameter = sqlCommand.Parameters.Add("@ProjectNumber", SqlDbType.Int);
                        //returnParameter.Direction = ParameterDirection.ReturnValue;
                        #endregion

                        sqlCommand.ExecuteNonQuery();

                        #region
                        //var result = Convert.ToInt32(returnParameter.Value);

                        //switch (result)
                        //{
                        //    case 0:
                        //        MessageBox.Show("Подбор №" + Номерподбора.Text + " уже существует!");
                        //        break;
                        //}
                        #endregion

                    }
                    catch (Exception exception)
                    {
                        Error(exception.ToString(), Convert.ToString(exception.GetHashCode()), "WriteToBase");
                    }
                    finally
                    {
                        con.Close();
                    }
                }
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="MakeDxfExportPartDataClass"/> class.
        /// </summary>
        public MakeDxfExportPartDataClass()
        {
            _sPathToSaveDxf = TestingCode ? @"C:\Temp\" : @"\\srvkb\DXF\";
            _xmlPath = TestingCode ? @"C:\Temp\" : @"\\srvkb\SolidWorks Admin\XML\";

            _шаблонЧертежаРазверткиВнеХранилища = @"\\srvkb\SolidWorks Admin\Templates\flattpattern.drwdot";
            // _шаблонЧертежаРазвертки = "\\Библиотека проектирования\\Templates\\flattpattern.drwdot";
            // _папкаШаблонов = "\\Библиотека проектирования\\Templates\\";
            _папкаШаблонов = @"\\srvkb\SolidWorks Admin\Templates\";
            _connectionString = "Data Source=srvkb;Initial Catalog=SWPlusDB;Persist Security Info=True;User ID=sa;Password=PDMadmin;MultipleActiveResultSets=True";
        }

        #region ModelCode

        const bool TestingCode = false;

        private const bool ШаблоныВХранилище = true;
        private readonly string _sPathToSaveDxf;
        private readonly string _xmlPath;
        private int _currentVersion;
        private string _eDrwFileName;
        private readonly string _шаблонЧертежаРазвертки;
        private readonly string _шаблонЧертежаРазверткиВнеХранилища;
        private readonly string _папкаШаблонов;
        private readonly string _connectionString;
        //
        /// <summary>
        /// Gets or sets the name of the PDM base. For example: "Vents-PDM"
        /// </summary>
        public string PdmBaseName { get; set; }

        /// <summary>
        /// Creates the flatt pattern update cutlist and edrawing.
        /// </summary>
        /// <param name="filePath">The file path.</param>
        /// <param name="eDrwFileName">Name of the e DRW file.</param>
        /// <param name="isErrors">if set to <c>true</c> [is errors].</param>
        /// <param name="makeDxf">if set to <c>true</c> [make DXF].</param>
        /// <param name="makeEprt">if set to <c>true</c> [make eprt].</param>
        /// <param name="swVisible">if set to <c>true</c> [sw visible].</param>
        /// <param name="relativePath"></param>
        public void CreateFlattPatternUpdateCutlistAndEdrawing(string filePath, out string eDrwFileName, out bool isErrors, bool makeDxf, bool makeEprt, bool swVisible, bool relativePath)
        {
            
            isErrors = false;
            eDrwFileName = "";
            var idPdm = 0;

            #region Сбор информации по детали и сохранение разверток
            
            SldWorks swApp = null;
            try
            {
                LoggerMine.Info("Запущен метод для обработки детали по пути " + filePath, "", "CreateFlattPatternUpdateCutlistAndEdrawing");

                var vault1 = new EdmVault5();
                vault1.LoginAuto(PdmBaseName, 0);

                if (relativePath)
                {
                    //filePath = filePath.Replace("D:", "E:"); // new FileInfo( vault1.RootFolder + "\\" + filePath ).FullName;
                    //MessageBox.Show(filePath);
                }

                try
                {
                    IEdmFolder5 oFolder;
                    var edmFile5 = vault1.GetFileFromPath(filePath, out oFolder);
                    edmFile5.GetFileCopy(0, 0, oFolder.ID, (int)EdmGetFlag.EdmGet_Simple);
                    idPdm = edmFile5.ID;
                    _currentVersion = edmFile5.CurrentVersion;
                }
                catch (Exception exception)
                {
                  LoggerMine.Error(
                      $"Ошибка при получении значения последней версии файла {Path.GetFileName(filePath)}", exception.ToString(), "CreateFlattPatternUpdateCutlistAndEdrawing");
                }

                try
                {
                    swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
                }
                catch (Exception)
                {
                    swApp = new SldWorks { Visible = swVisible };
                }
                if (swApp == null)
                {
                    isErrors = true;
                    LoggerMine.Info("isErrors = true на 93-й строке ", "", "");
                    return; 
                }

                var thisProc = Process.GetProcessesByName("SLDWORKS")[0];
                thisProc.PriorityClass = ProcessPriorityClass.RealTime;

                try
                {
                    swApp.Visible = swVisible;
                }
                catch (Exception exception)
                {
                    LoggerMine.Error($"Ошибка при попытке погашния {Path.GetFileName(filePath)}", exception.ToString(), "CreateFlattPatternUpdateCutlistAndEdrawing");
                }

                IModelDoc2 swModel;

                #region To Delete

                //try
                //{
                //IEdmFolder5 oFolder;
                //var edmFile5 = vault1.GetFileFromPath(filePath, out oFolder);
                //edmFile5.GetFileCopy(0, 0, oFolder.ID, (int)EdmGetFlag.EdmGet_Simple);
                //_currentVersion = edmFile5.CurrentVersion;
                //    //swApp.SetUserPreferenceStringValue(((int)(swUserPreferenceStringValue_e.swFileLocationsDocumentTemplates)), vault1.RootFolderPath + _папкаШаблонов);
                //    swApp.SetUserPreferenceStringValue(((int)(swUserPreferenceStringValue_e.swFileLocationsDocumentTemplates)), _папкаШаблонов);
                //}
                //catch (Exception exception)
                //{
                //    LoggerMine.Error(String.Format("Ошибка: {0} Строка: {1}", exception.ToString(), exception.StackTrace), exception.Source, "CreateFlattPatternUpdateCutlistAndEdrawing");
                //}

                #endregion

                try
                {
                    swModel = swApp.OpenDoc6(filePath, (int)swDocumentTypes_e.swDocPART,
                                (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", 0, 0);
                  
                }
                catch (Exception exception)
                {
                    LoggerMine.Error(string.Format("Ошибка при обработке детали {2}: {0} Строка: {1}", exception.ToString(), exception.StackTrace,
                        Path.GetFileName(filePath)), exception.Source, "CreateFlattPatternUpdateCutlistAndEdrawing");
                    isErrors = true;
                    LoggerMine.Info("isErrors = true на 127-й строке ", "", "");
                    return;
                }

                try
                {
                    if (!IsSheetMetalPart((IPartDoc)swModel))
                    {
                        swApp.CloseDoc(Path.GetFileName(swModel.GetPathName()));
                        if (!makeDxf) return;
                        swApp.ExitApp();
                        swApp = null;
                        return;
                    }
                }
                catch (Exception exception)
                {
                    LoggerMine.Error(string.Format("Ошибка2 при обработке детали {2}: {0} Строка: {1}", exception.ToString(), exception.StackTrace, 
                        Path.GetFileName(filePath)), exception.Source, "CreateFlattPatternUpdateCutlistAndEdrawing");
                    isErrors = true;
                    LoggerMine.Info("isErrors = true на 145-й строке ", "", "");
                }

                Configuration activeconfiguration;
                string[] swModelConfNames;

                try
                {
                    activeconfiguration = (Configuration)swModel.GetActiveConfiguration();
                    swModelConfNames = (string[])swModel.GetConfigurationNames();
                }
                catch (Exception exception)
                {
                    LoggerMine.Error(exception.ToString(),"", "");
                    isErrors = true;
                    LoggerMine.Info("isErrors = true на 160-й строке ", "", "");
                    return;
                }
                
                try
                {
                    foreach (var name in from name in swModelConfNames
                                         let config = (Configuration)swModel.GetConfigurationByName(name)
                                         where config.IsDerived()
                                         select name)
                    {
                        try
                        {
                            swModel.DeleteConfiguration(name);
                        }
                        catch (Exception exception)
                        {
                            LoggerMine.Error(string.Format("Ошибка при удалении конфигурации '{2}' в модели '{3}': {0} Строка: {1}", exception.ToString(), exception.StackTrace, name, swModel.GetTitle()),
                                exception.Source, "CreateFlattPatternUpdateCutlistAndEdrawing");
                        }
                    }
                }
                catch (Exception exception)
                {
                    LoggerMine.Error(string.Format("Ошибка при удалении конфигураций в модели '{2}': {0} Строка: {1}", exception.ToString(), exception.StackTrace, swModel.GetTitle()),
                                 exception.Source, "CreateFlattPatternUpdateCutlistAndEdrawing");
                    LoggerMine.Info("isErrors = true на 186-й строке ", "", "");
                    isErrors = true;
                }

                ModelDocExtension swModelDocExt;
                string[] swModelConfNames2;

                try
                {
                    swModelDocExt = swModel.Extension;
                    swModelConfNames2 = (string[])swModel.GetConfigurationNames();
                }
                catch (Exception exception)
                {
                    LoggerMine.Error(exception.ToString(), "", "");
                    isErrors = true;
                    LoggerMine.Info("isErrors = true на 202-й строке ", "", "");
                    return;
                }
                
                // Проход по всем родительским конфигурациям

                var dataList = new List<DataToExport>();
                
                try
                {
                    foreach (var configName in from name in swModelConfNames2
                                               let config = (Configuration)swModel.GetConfigurationByName(name)
                                               where !config.IsDerived()
                                               select name)
                    {

                        swModel.ShowConfiguration2(configName);
                        swModel.EditRebuild3();

                     //   MessageBox.Show(" Попытка вставить размеры ");
                        //new Bends().Fix(swApp);


                        //GabaritsForPaintingCamera(swModel, configName);
                        //MessageBox.Show(" Размеры добавлены? ");

                        var confiData = new DataToExport
                        {
                            Config = configName,
                            FileName = filePath.Substring(filePath.LastIndexOf('\\') + 1),
                            IdPdm = idPdm
                        };
                        
                        //FileInfo template = null;

                        //try
                        //{
                        // //   template = new FileInfo(_шаблонЧертежаРазверткиВнеХранилища);
                        //   // Thread.Sleep(1000);
                        //}
                        //catch (Exception exception)
                        //{
                        //    LoggerMine.Error("Проблемы с получением шаблона чертежа", exception.StackTrace, "CreateFlattPatternUpdateCutlistAndEdrawing");
                        //  //  template = new FileInfo(_шаблонЧертежаРазверткиВнеХранилища);
                        // //   Thread.Sleep(1000);

                        //}
                        //finally
                        //{
                        //    if (template == null)
                        //    {
                        //        template = new FileInfo(_шаблонЧертежаРазверткиВнеХранилища);
                        //    }
                        //}

                        //if (!template.Exists)
                        //{
                        //    LoggerMine.Error("Не удалось найти шаблон чертежа по пути \n" + template.FullName + "\nПроверте подключение к " + template.Directory, "", "CreateFlattPatternUpdateCutlistAndEdrawing");
                        //    isErrors = true;
                        //    LoggerMine.Info("isErrors = true на 229-й строке ", "", "");
                        //}

                        if (swApp != null)
                        {
                            DrawingDoc swDraw = null;
                            //if (makeDxf)
                            //{
                            //     swDraw = (DrawingDoc)
                            //        swApp.INewDocument2(template.FullName, (int)swDwgPaperSizes_e.swDwgPaperA0size, 0.841, 0.594);
                            //    swDraw.CreateFlatPatternViewFromModelView3(swModel.GetPathName(), configName, 0.841 / 2, 0.594 / 2, 0, true, false);
                            //    ((IModelDoc2)swDraw).ForceRebuild3(true);
                            //}

                            #region Разгибание всех сгибов

                            try
                            {
                                swModel.EditRebuild3();
                                //var swPart = (IPartDoc)swModel;


                                new Bends().Fix(swApp);


                                //Feature swFeature = swPart.FirstFeature();
                                //const string strSearch = "FlatPattern";
                                //while (swFeature != null)
                                //{
                                //    var nameTypeFeature = swFeature.GetTypeName2();

                                //    if (nameTypeFeature == strSearch)
                                //    {
                                //        swFeature.Select(true);
                                //        swPart.EditUnsuppress();

                                //        Feature swSubFeature = swFeature.GetFirstSubFeature();
                                //        while (swSubFeature != null)
                                //        {
                                //            var nameTypeSubFeature = swSubFeature.GetTypeName2();

                                //            if (nameTypeSubFeature == "UiBend")
                                //            {
                                //                swFeature.Select(true);
                                //                swPart.EditUnsuppress();
                                //                swModel.EditRebuild3();

                                //                try
                                //                {
                                //                    swSubFeature.SetSuppression2(
                                //                        (int)swFeatureSuppressionAction_e.swUnSuppressDependent,
                                //                        (int)swInConfigurationOpts_e.swConfigPropertySuppressFeatures,
                                //                        swModelConfNames2);

                                //                    //swSubFeature.SetSuppression2(
                                //                    //    (int)swFeatureSuppressionAction_e.swUnSuppressFeature,
                                //                    //    (int)swInConfigurationOpts_e.swAllConfiguration,
                                //                    //    swModelConfNames2);
                                //                }
                                //                catch (Exception)
                                //                {

                                //                }
                                //            }
                                //            swSubFeature = swSubFeature.GetNextSubFeature();
                                //        }
                                //    }
                                //    swFeature = swFeature.GetNextFeature();
                                //}

                                //swModel.EditRebuild3();


                            }
                            catch (Exception exception)
                            {
                                MessageBox.Show(exception.ToString());
                            }

                            #endregion

                            swModel.ForceRebuild3(false);

                            var swCustProp = swModelDocExt.CustomPropertyManager[configName];
                            string valOut;

                            string materialId;
                            swCustProp.Get4("MaterialID", true, out valOut, out materialId);
                            if (string.IsNullOrEmpty(materialId))
                            {
                                confiData.MaterialId = null;
                            }
                            else
                            {
                                confiData.MaterialId = int.Parse(materialId);
                            }

                            string paintX;
                            swCustProp.Get4("Длина", true, out valOut, out paintX);
                            if (string.IsNullOrEmpty(paintX))
                            {
                                confiData.PaintX = null;
                            }
                            else
                            {
                                confiData.PaintX = double.Parse(paintX);
                            }

                            string paintY;
                            swCustProp.Get4("Ширина", true, out valOut, out paintY);
                            if (string.IsNullOrEmpty(paintY))
                            {
                                confiData.PaintY = null;
                            }
                            else
                            {
                                confiData.PaintY = double.Parse(paintY);
                            }

                            string paintZ;
                            swCustProp.Get4("Высота", true, out valOut, out paintZ);
                            if (string.IsNullOrEmpty(paintZ))
                            {
                                confiData.PaintZ = null;
                            }
                            else
                            {
                                confiData.PaintZ = double.Parse(paintZ);
                            }

                            string codMaterial;
                            swCustProp.Get4("Код материала", true, out valOut, out codMaterial);
                            confiData.КодМатериала = codMaterial;

                            string материал;
                            swCustProp.Get4("Материал", true, out valOut, out материал);
                            confiData.Материал = материал;

                            string обозначение;
                            swCustProp.Get4("Обозначение", true, out valOut, out обозначение);
                            confiData.Обозначение = обозначение;

                            var swCustPropForDescription = swModelDocExt.CustomPropertyManager[""];
                            string наименование;
                            swCustPropForDescription.Get4("Наименование", true, out valOut, out наименование);
                            confiData.Наименование = наименование;

                            var thikness = GetFromCutlist(swModel, "Толщина листового металла");

                            if (makeDxf)
                            {
                                var errors = 0;
                                var warnings = 0;
                                var newDxf = (IModelDoc2)swDraw;

                                newDxf.Extension.SaveAs(
                                    _sPathToSaveDxf + Path.GetFileNameWithoutExtension(swModel.GetPathName()) + "-" + configName + "-" + thikness + ".dxf",  
                                    (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                                    (int)swSaveAsOptions_e.swSaveAsOptions_UpdateInactiveViews, null, ref errors, ref warnings);

                                swApp.CloseDoc(Path.GetFileName(newDxf.GetPathName()));
                            }
                            

                            //UpdateCustomPropertyListFromCutList
                            const string длинаГраничнойРамкиName = "Длина граничной рамки";
                            const string ширинаГраничнойРамкиName = "Ширина граничной рамки";
                            const string толщинаЛистовогоМеталлаNAme = "Толщина листового металла";
                            const string сгибыName = "Сгибы";
                            const string площадьПокрытияName = "Площадь покрытия";

                            Feature swFeat2 = swModel.FirstFeature();
                            while (swFeat2 != null)
                            {
                                if (swFeat2.GetTypeName2() == "SolidBodyFolder")
                                {
                                    new Bends().Fix(swApp);
                                    BodyFolder swBodyFolder = swFeat2.GetSpecificFeature2();
                                    swFeat2.Select2(false, -1);
                                    swBodyFolder.SetAutomaticCutList(true);
                                    swBodyFolder.UpdateCutList();

                                    Feature swSubFeat = swFeat2.GetFirstSubFeature();
                                    while (swSubFeat != null)
                                    {
                                        if (swSubFeat.GetTypeName2() == "CutListFolder")
                                        {
                                            BodyFolder bodyFolder = swSubFeat.GetSpecificFeature2();
                                            swSubFeat.Select2(false, -1);
                                            bodyFolder.SetAutomaticCutList(true);
                                            bodyFolder.UpdateCutList();
                                            var swCustPrpMgr = swSubFeat.CustomPropertyManager;
                                            swCustPrpMgr.Add("Площадь поверхности", "Текст",
                                                "\"SW-SurfaceArea@@@Элемент списка вырезов1@" + Path.GetFileName(swModel.GetPathName()) + "\"");

                                            string длинаГраничнойРамки;
                                            swCustPrpMgr.Get4(длинаГраничнойРамкиName, true, out valOut,
                                                out длинаГраничнойРамки);
                                            swCustProp.Set(длинаГраничнойРамкиName, длинаГраничнойРамки);
                                            confiData.ДлинаГраничнойРамки = длинаГраничнойРамки;

                                            string ширинаГраничнойРамки;
                                            swCustPrpMgr.Get4(ширинаГраничнойРамкиName, true, out valOut,
                                                out ширинаГраничнойРамки);
                                            swCustProp.Set(ширинаГраничнойРамкиName, ширинаГраничнойРамки);
                                            confiData.ШиринаГраничнойРамки = ширинаГраничнойРамки;

                                            string толщинаЛистовогоМеталла;
                                            swCustPrpMgr.Get4(толщинаЛистовогоМеталлаNAme, true, out valOut,
                                                out толщинаЛистовогоМеталла);
                                            swCustProp.Set(толщинаЛистовогоМеталлаNAme, толщинаЛистовогоМеталла);
                                            confiData.ТолщинаЛистовогоМеталла = толщинаЛистовогоМеталла;

                                            string сгибы;
                                            swCustPrpMgr.Get4(сгибыName, true, out valOut, out сгибы);
                                            swCustProp.Set(сгибыName, сгибы);
                                            confiData.Сгибы = сгибы;

                                            //string площадьПоверхности;
                                            //swCustPrpMgr.Get4("Площадь поверхности", true, out valOut,  out площадьПоверхности);
                                            //swCustProp.Set(площадьПокрытияName, площадьПоверхности);

                                            var myMassProp = swModel.Extension.CreateMassProperty();
                                            var площадьПоверхности = Convert.ToString( Math.Round(myMassProp.SurfaceArea * 1000)/1000
                                                );
                                            
                                            swCustProp.Set(площадьПокрытияName, площадьПоверхности);

                                            //MessageBox.Show("ПлощадьПокрытия = " + площадьПоверхности);

                                            try
                                            {
                                               // MessageBox.Show("ПлощадьПокрытия = double.Parse(" + площадьПоверхности);
                                                confiData.ПлощадьПокрытия = double.Parse(площадьПоверхности.Replace(".", ","));
                                            }
                                            catch (Exception)
                                            {
                                                //MessageBox.Show(exception.ToString());
                                                //confiData.ПлощадьПокрытия = double.Parse(площадьПоверхности.Replace(",", "."));
                                            }
                                            finally
                                            {
                                               // MessageBox.Show("confiData.ПлощадьПокрытия = " + площадьПоверхности);
                                            }
                                        }
                                        swSubFeat = swSubFeat.GetNextFeature();
                                    }
                                }
                                swFeat2 = swFeat2.GetNextFeature();
                            }
                        }
                        dataList.Add(confiData);
                    }
                }
                catch (Exception exception)
                {
                    //MessageBox.Show(exception.ToString());
                    LoggerMine.Error(exception.ToString(), "Строка 377", "");
                }

                try
                {
                    swModel.ShowConfiguration2(activeconfiguration.Name);

                }
                catch (Exception exception)
                {
                    //MessageBox.Show(exception.ToString());
                    LoggerMine.Error(exception.ToString(), "", "");
                    isErrors = true;
                    LoggerMine.Info("isErrors = true на 392-й строке ", "", "");
                }

                try
                {
                    ExportDataToXmlSql(swModel, dataList);
                }
                catch (Exception exception)
                {
                    //MessageBox.Show(exception.ToString());
                    LoggerMine.Error(exception.ToString(), "", "");
                    isErrors = true;
                    LoggerMine.Info("isErrors = true на 403-й строке ", "", "");
                }
                
            #endregion
                
                #region Сохранение детали в eDrawing

                if (makeEprt)
                {
                    string modelName;
                    try
                    {
                        modelName = Path.GetFileNameWithoutExtension(swModel.GetPathName());
                        _eDrwFileName = Path.GetDirectoryName(swModel.GetPathName()) + "\\" + modelName + ".eprt";
                        eDrwFileName = _eDrwFileName;
                    }
                    catch (Exception exception)
                    {
                        LoggerMine.Error(exception.ToString(), "", "");
                        isErrors = true;
                        LoggerMine.Info("isErrors = true на 423-й строке ", "", "");
                        return;
                    }

                    try
                    {
                        // todo: удаление документов перед новым сохранением. Осуществить поиск по имени
                        var existingDocument = SearchDoc(modelName + ".eprt", SwDocType.SwDocNone);

                        if (existingDocument != "")
                        {
                            LoggerMine.Info($"Файл есть в базе {modelName} и будет удален.. ", "", "CreateFlattPatternUpdateCutlistAndEdrawing");
                            //CheckInOutPdm(new List<FileInfo> { new FileInfo(eDrwFileName) }, false, PdmBaseName);
                            DeleteFileFromPdm(existingDocument, PdmBaseName);
                        }
                        else
                        {
                            File.Delete(_eDrwFileName);
                        }
                    }
                    catch (Exception)
                    {
                        try
                        {
                            File.Delete(_eDrwFileName);
                        }
                        catch (Exception exception)
                        {
                            LoggerMine.Error(exception.StackTrace, "", "File.Delete(_eDrwFileName);");
                            isErrors = true;
                            LoggerMine.Info("isErrors = true на 453-й строке ", "", "");
                        }
                    }

                    #region ToDelete
                    //if (new FileInfo(_eDrwFileName).Exists)
                    //{
                    //    LoggerMine.Info("Файл есть в базе " + swModel.GetTitle());
                    //    //CheckInOutPdm(new List<FileInfo> { new FileInfo(eDrwFileName) }, false, PdmBaseName);
                    //    DeleteFileFromPdm(_eDrwFileName, PdmBaseName);
                    //}
                    #endregion

                    try
                    {
                        swApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swEdrawingsSaveAsSelectionOption,
                        (int)swEdrawingSaveAsOption_e.swEdrawingSaveAll);
                        swApp.SetUserPreferenceToggle(((int)(swUserPreferenceToggle_e.swEDrawingsOkayToMeasure)), true);
                        swModel.Extension.SaveAs(_eDrwFileName, (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                                (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, 0, 0);
                    }
                    catch (Exception exception)
                    {
                        LoggerMine.Error(exception.StackTrace, "", "File.Delete(_eDrwFileName);");
                        isErrors = true;
                        LoggerMine.Info("isErrors = true на 478-й строке ", "", "");
                    }

                    #region To delete

                    //try
                    //{
                    //    // swApp.CloseDoc(Path.GetFileName(swModel.GetPathName()));
                    //    swApp.CloseDoc(modelName + ".sldprt");

                    //    if (makeDxf)
                    //    {
                    //        swApp.ExitApp();
                    //        swApp = null;
                    //    }
                    //    LoggerMine.Info("Обработка файла " + modelName + ".sldprt" + " успешно завершена", "", "CreateFlattPatternUpdateCutlistAndEdrawing");
                    //    isErrors = false;
                    //}
                    //catch (Exception exception)
                    //{
                    //    LoggerMine.Error(exception.StackTrace, "", "File.Delete(_eDrwFileName);");
                    //    isErrors = true;
                    //    LoggerMine.Info("isErrors = true на 497-й строке ", "", "");
                    //}

                    #endregion

                }
                
                #endregion

                try
                {
                    // swApp.CloseDoc(Path.GetFileName(swModel.GetPathName()));
                    var namePrt = swApp.IActiveDoc2.GetTitle().ToLower().Contains(".sldprt")
                        ? swApp.IActiveDoc2.GetTitle()
                        : swApp.IActiveDoc2.GetTitle() + ".sldprt";
                    swApp.CloseDoc(namePrt);

                    if (makeDxf)
                    {
                        swApp.ExitApp();
                        swApp = null;
                    }
                    LoggerMine.Info(
                        "Обработка файла " + swApp.IActiveDoc2.GetTitle() + ".sldprt" + ".sldprt" + " успешно завершена",
                        "", "CreateFlattPatternUpdateCutlistAndEdrawing");
                    isErrors = false;
                }
                catch (Exception exception)
                {
                    LoggerMine.Error(exception.StackTrace, "", "File.Delete(_eDrwFileName);");
                    isErrors = true;
                    LoggerMine.Info("isErrors = true на 497-й строке ", "", "");
                }

                #region
                finally
                {
                    //try
                    //{
                    //    swApp.CloseDoc(swApp.IActiveDoc2.GetTitle() + ".sldprt");
                    //}
                    //catch (Exception)
                    //{
                    //    swApp.CloseDoc(Path.GetFileName(swModel.GetPathName()));
                    //}
                }
                #endregion
                
            }
            catch (Exception exception)
            {
                LoggerMine.Error(
                    $"Общая ошибка метода: {exception.ToString()} Строка: {exception.StackTrace} exception.Source - ", exception.Source, "CreateFlattPatternUpdateCutlistAndEdrawing");
                if (swApp == null) return;
                if (makeDxf)
                {
                    swApp.ExitApp();
                }
                isErrors = true;
                LoggerMine.Info("isErrors = true на 506-й строке ", "", "");
            }
        }

        static void GabaritsForPaintingCamera(IModelDoc2 swmodel)
        {
            try
            {
                const long valueset = 1000;
                const int swDocPart = 1;
                const int swDocAssembly = 2;

                for (var i = 0; i < swmodel.GetConfigurationCount(); i++)
                {
                    i = i + 1;
                    var configname = swmodel.IGetConfigurationNames(ref i);
                    
                    Configuration swConf = swmodel.GetConfigurationByName(configname);
                    if (swConf.IsDerived()) continue;
                    //swmodel.ShowConfiguration2(configname);
                    swmodel.EditRebuild3();

                    switch (swmodel.GetType())
                    {
                        case swDocPart:
                        {
                                // MessageBox.Show("swDocPart");
                                var part = (PartDoc) swmodel;
                                var box = part.GetPartBox(true);

                                swmodel.AddCustomInfo3(configname, "Длина", 30, "");
                                swmodel.AddCustomInfo3(configname, "Ширина", 30, "");
                                swmodel.AddCustomInfo3(configname, "Высота", 30, "");

                                // swmodel.AddCustomInfo3(configname, "Длина", , "");

                                // MessageBox.Show(configname);

                                swmodel.CustomInfo2[configname, "Длина"] =
                                    Convert.ToString(
                                    Math.Round(Convert.ToDecimal((long) (Math.Abs(box[0] - box[3])*valueset)), 0),
                                    CultureInfo.InvariantCulture);
                                swmodel.CustomInfo2[configname, "Ширина"] =
                                    Convert.ToString(
                                    Math.Round(Convert.ToDecimal((long) (Math.Abs(box[1] - box[4])*valueset)), 0),
                                    CultureInfo.InvariantCulture);
                                swmodel.CustomInfo2[configname, "Высота"] =
                                    Convert.ToString(
                                    Math.Round(Convert.ToDecimal((long) (Math.Abs(box[2] - box[5])*valueset)), 0),
                                    CultureInfo.InvariantCulture);
                        }
                            break;

                        case swDocAssembly:
                        {
                                //MessageBox.Show("AssemblyDoc");
                                var swAssy = (AssemblyDoc) swmodel;

                                var boxAss = swAssy.GetBox((int) swBoundingBoxOptions_e.swBoundingBoxIncludeRefPlanes);

                                swmodel.AddCustomInfo3(configname, "Длина", 30, "");
                                swmodel.AddCustomInfo3(configname, "Ширина", 30, "");
                                swmodel.AddCustomInfo3(configname, "Высота", 30, "");

                                swmodel.CustomInfo2[configname, "Длина"] =
                                    Convert.ToString(
                                        Math.Round(Convert.ToDecimal((long) (Math.Abs(boxAss[0] - boxAss[3])*valueset)), 0),
                                        CultureInfo.InvariantCulture);
                                swmodel.CustomInfo2[configname, "Ширина"] =
                                    Convert.ToString(
                                        Math.Round(Convert.ToDecimal((long) (Math.Abs(boxAss[1] - boxAss[4])*valueset)), 0),
                                        CultureInfo.InvariantCulture);
                                swmodel.CustomInfo2[configname, "Высота"] =
                                    Convert.ToString(
                                        Math.Round(Convert.ToDecimal((long) (Math.Abs(boxAss[2] - boxAss[5])*valueset)), 0),
                                        CultureInfo.InvariantCulture);
                        }
                            break;
                    }
                    swmodel.EditRebuild3();
                }
            }
            catch (Exception )
            {
                //MessageBox.Show(exception.ToString());
            }
        }

        static void GabaritsForPaintingCamera(IModelDoc2 swmodel, string configname)
        {
            try
            {
                const long valueset = 1000;

                    swmodel.GetConfigurationByName(configname);
                    swmodel.EditRebuild3();

                var part = (PartDoc)swmodel;
                var box = part.GetPartBox(true);

                swmodel.AddCustomInfo3(configname, "Длина", 30, "");
                swmodel.AddCustomInfo3(configname, "Ширина", 30, "");
                swmodel.AddCustomInfo3(configname, "Высота", 30, "");


               // MessageBox.Show(configname);

                swmodel.CustomInfo2[configname, "Длина"] =
                    Convert.ToString(
                        Math.Round(Convert.ToDecimal((long)(Math.Abs(box[0] - box[3]) * valueset)), 0),
                        CultureInfo.InvariantCulture);
                swmodel.CustomInfo2[configname, "Ширина"] =
                    Convert.ToString(
                        Math.Round(Convert.ToDecimal((long)(Math.Abs(box[1] - box[4]) * valueset)), 0),
                        CultureInfo.InvariantCulture);
                swmodel.CustomInfo2[configname, "Высота"] =
                    Convert.ToString(
                        Math.Round(Convert.ToDecimal((long)(Math.Abs(box[2] - box[5]) * valueset)), 0),
                        CultureInfo.InvariantCulture);

                swmodel.EditRebuild3();
            }
            catch (Exception exception)
            {
                //MessageBox.Show(exception.ToString());
            }
        }


        /// <summary>
        /// Creates the flatt pattern update cutlist and edrawing2.
        /// </summary>
        /// <param name="filePath">The file path.</param>
        /// <param name="eDrwFileName">Name of the e DRW file.</param>
        /// <param name="isErrors">if set to <c>true</c> [is errors].</param>
        /// <param name="makeDxf">if set to <c>true</c> [make DXF].</param>
        /// <param name="makeEprt">if set to <c>true</c> [make eprt].</param>
        /// <param name="swVisible">if set to <c>true</c> [sw visible].</param>
        public void CreateFlattPatternUpdateCutlistAndEdrawing2(string filePath, out string eDrwFileName, out bool isErrors, bool makeDxf, bool makeEprt, bool swVisible)
        {
            isErrors = false;

            eDrwFileName = "";
            var idPdm = 0;

                #region Сбор информации по детали и сохранение разверток

            SldWorks swApp = null;
            try
            {
                LoggerMine.Info("Запущен метод для обработки детали по пути " + filePath, "", "CreateFlattPatternUpdateCutlistAndEdrawing");

                var vault1 = new EdmVault5();
                vault1.LoginAuto(PdmBaseName, 0);

                try
                {
                    IEdmFolder5 oFolder;
                    var edmFile5 = vault1.GetFileFromPath(filePath, out oFolder);
                    edmFile5.GetFileCopy(0, 0, oFolder.ID, (int)EdmGetFlag.EdmGet_Simple);
                    idPdm = edmFile5.ID;
                    _currentVersion = edmFile5.CurrentVersion;
                }
                catch (Exception exception)
                {
                    LoggerMine.Error(
                        $"Ошибка при получении значения последней версии файла {Path.GetFileName(filePath)}", exception.ToString(), "CreateFlattPatternUpdateCutlistAndEdrawing");
                }

                try
                {
                    swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
                }
                catch (Exception)
                {
                    swApp = new SldWorks { Visible = swVisible };
                }
                if (swApp == null)
                {
                    isErrors = true;
                    return;
                }
                try
                {
                    swApp.Visible = swVisible;
                }
                catch (Exception exception)
                {
                    LoggerMine.Error($"Ошибка при попытке погашния {Path.GetFileName(filePath)}", exception.ToString(), "CreateFlattPatternUpdateCutlistAndEdrawing");
                }

                IModelDoc2 swModel;

                try
                {
                    swModel = swApp.OpenDoc6(filePath, (int)swDocumentTypes_e.swDocPART,
                                (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", 0, 0);
                    swModel.Extension.ViewDisplayRealView = false;
                }
                catch (Exception exception)
                {
                    LoggerMine.Error(string.Format("Ошибка при обработке детали {2}: {0} Строка: {1}", exception.ToString(), exception.StackTrace, Path.GetFileName(filePath)), exception.Source, "CreateFlattPatternUpdateCutlistAndEdrawing");
                    isErrors = true;
                    return;
                }

                try
                {
                    if (!IsSheetMetalPart((IPartDoc)swModel))
                    {
                        swApp.CloseDoc(Path.GetFileName(swModel.GetPathName()));
                        if (!makeDxf) return;
                        swApp.ExitApp();
                        swApp = null;
                        return;
                    }
                }
                catch (Exception exception)
                {
                    LoggerMine.Error(string.Format("Ошибка2 при обработке детали {2}: {0} Строка: {1}", exception.ToString(), exception.StackTrace, Path.GetFileName(filePath)), exception.Source, "CreateFlattPatternUpdateCutlistAndEdrawing");
                    isErrors = true;
                }

                Configuration activeconfiguration;
                string[] swModelConfNames;

                try
                {
                    activeconfiguration = (Configuration)swModel.GetActiveConfiguration();
                    swModelConfNames = (string[])swModel.GetConfigurationNames();
                }
                catch (Exception exception)
                {
                    LoggerMine.Error(exception.ToString(), "", "");
                    isErrors = true;
                    return;
                }

                if (!makeEprt)
                {
                    try
                    {
                        swModel.EditRebuild3();
                        var swPart = (PartDoc)swModel;
                        var arrNamesConfig = (string[])swModel.GetConfigurationNames();

                        Feature swFeature = swPart.FirstFeature();
                        const string strSearch = "FlatPattern";

                        while (swFeature != null)
                        {
                            var nameTypeFeature = swFeature.GetTypeName2();

                            if (nameTypeFeature == strSearch)
                            {
                                Feature swSubFeature = swFeature.GetFirstSubFeature();
                                while (swSubFeature != null)
                                {
                                    var nameTypeSubFeature = swSubFeature.GetTypeName2();

                                    if (nameTypeSubFeature == "UiBend")
                                    {
                                        swSubFeature.SetSuppression2(
                                                       (int)swFeatureSuppressionAction_e.swUnSuppressDependent,
                                                       (int)swInConfigurationOpts_e.swAllConfiguration,
                                                       arrNamesConfig);

                                        //swSubFeature.SetSuppression2(
                                        //    (int)swFeatureSuppressionAction_e.swUnSuppressFeature,
                                        //    (int)swInConfigurationOpts_e.swAllConfiguration,
                                        //    arrNamesConfig);
                                    }
                                    swSubFeature = swSubFeature.GetNextSubFeature();
                                }
                            }
                            swFeature = swFeature.GetNextFeature();
                        }
                        swModel.EditRebuild3();
                    }
                    catch (Exception)
                    {
                        throw;
                    }
                }
                
                try
                {
                    foreach (var name in from name in swModelConfNames
                                         let config = (Configuration)swModel.GetConfigurationByName(name)
                                         where config.IsDerived()
                                         select name)
                    {
                        try
                        {
                            swModel.DeleteConfiguration(name);
                        }
                        catch (Exception exception)
                        {
                            LoggerMine.Error(string.Format("Ошибка при удалении конфигурации '{2}' в модели '{3}': {0} Строка: {1}",
                                exception.ToString(), exception.StackTrace, name, swModel.GetTitle()),
                                exception.Source, "CreateFlattPatternUpdateCutlistAndEdrawing");
                        }
                    }
                }
                catch (Exception exception)
                {
                    LoggerMine.Error(string.Format("Ошибка при удалении конфигураций в модели '{2}': {0} Строка: {1}",
                        exception.ToString(), exception.StackTrace, swModel.GetTitle()),
                        exception.Source, "CreateFlattPatternUpdateCutlistAndEdrawing");
                    LoggerMine.Info("isErrors = true на 186-й строке ", "", "");
                    isErrors = true;
                }

                ModelDocExtension swModelDocExt;
                string[] swModelConfNames2;

                try
                {
                    swModelDocExt = swModel.Extension;
                    swModelConfNames2 = (string[])swModel.GetConfigurationNames();
                }
                catch (Exception exception)
                {
                    LoggerMine.Error(exception.ToString(), "", "");
                    isErrors = true;
                    return;
                }

                // Проход по всем родительским конфигурациям

                var dataList = new List<DataToExport>();

                try
                {
                    foreach (var configName in from name in swModelConfNames2
                                               let config = (Configuration)swModel.GetConfigurationByName(name)
                                               where !config.IsDerived()
                                               select name)
                    {
                        swModel.ShowConfiguration2(configName);
                        swModel.EditRebuild3();

                        var confiData = new DataToExport
                        {
                            Config = configName,
                            FileName = filePath.Substring(filePath.LastIndexOf('\\') + 1),
                            IdPdm = idPdm
                        };

                        if (swApp != null)
                        {
                            swModel.ForceRebuild3(false);

                            var swCustProp = swModelDocExt.CustomPropertyManager[configName];
                            string valOut;

                            string materialId;
                            swCustProp.Get4("MaterialID", true, out valOut, out materialId);
                            if (string.IsNullOrEmpty(materialId))
                            {
                                confiData.MaterialId = null;
                            }
                            else
                            {
                                confiData.MaterialId = int.Parse(materialId);
                            }


                            string paintX;
                            swCustProp.Get4("Длина", true, out valOut, out paintX);
                            if (string.IsNullOrEmpty(paintX))
                            {
                                confiData.PaintX = null;
                            }
                            else
                            {
                                confiData.PaintX = double.Parse(paintX);
                            }
                            

                            string paintY;
                            swCustProp.Get4("Ширина", true, out valOut, out paintY);
                            if (string.IsNullOrEmpty(paintY))
                            {
                                confiData.PaintY = null;
                            }
                            else
                            {
                                confiData.PaintY = double.Parse(paintY);
                            }

                            string paintZ;
                            swCustProp.Get4("Высота", true, out valOut, out paintZ);
                            if (string.IsNullOrEmpty(paintZ))
                            {
                                confiData.PaintZ = null;
                            }
                            else
                            {
                                confiData.PaintZ = double.Parse(paintZ);
                            }


                            string codMaterial;
                            swCustProp.Get4("Код материала", true, out valOut, out codMaterial);
                            confiData.КодМатериала = codMaterial;

                            string материал;
                            swCustProp.Get4("Материал", true, out valOut, out материал);
                            confiData.Материал = материал;

                            string обозначение;
                            swCustProp.Get4("Обозначение", true, out valOut, out обозначение);
                            confiData.Обозначение = обозначение;

                            var swCustPropForDescription = swModelDocExt.CustomPropertyManager[""];
                            string наименование;
                            swCustPropForDescription.Get4("Наименование", true, out valOut, out наименование);
                            confiData.Наименование = наименование;

                            var thikness = GetFromCutlist(swModel, "Толщина листового металла");


                            //UpdateCustomPropertyListFromCutList
                            const string длинаГраничнойРамкиName = "Длина граничной рамки";
                            const string ширинаГраничнойРамкиName = "Ширина граничной рамки";
                            const string толщинаЛистовогоМеталлаNAme = "Толщина листового металла";
                            const string сгибыName = "Сгибы";
                            const string площадьПокрытияName = "Площадь покрытия";

                            Feature swFeat2 = swModel.FirstFeature();
                            while (swFeat2 != null)
                            {
                                if (swFeat2.GetTypeName2() == "SolidBodyFolder")
                                {
                                    BodyFolder swBodyFolder = swFeat2.GetSpecificFeature2();
                                    swFeat2.Select2(false, -1);
                                    swBodyFolder.SetAutomaticCutList(true);
                                    swBodyFolder.UpdateCutList();

                                    Feature swSubFeat = swFeat2.GetFirstSubFeature();
                                    while (swSubFeat != null)
                                    {
                                        if (swSubFeat.GetTypeName2() == "CutListFolder")
                                        {
                                            BodyFolder bodyFolder = swSubFeat.GetSpecificFeature2();
                                            swSubFeat.Select2(false, -1);
                                            bodyFolder.SetAutomaticCutList(true);
                                            bodyFolder.UpdateCutList();
                                            var swCustPrpMgr = swSubFeat.CustomPropertyManager;
                                            swCustPrpMgr.Add("Площадь поверхности", "Текст",
                                                "\"SW-SurfaceArea@@@Элемент списка вырезов1@" + Path.GetFileName(swModel.GetPathName()) + "\"");

                                            string длинаГраничнойРамки;
                                            swCustPrpMgr.Get4(длинаГраничнойРамкиName, true, out valOut,
                                                out длинаГраничнойРамки);
                                            swCustProp.Set(длинаГраничнойРамкиName, длинаГраничнойРамки);
                                            confiData.ДлинаГраничнойРамки = длинаГраничнойРамки;

                                            string ширинаГраничнойРамки;
                                            swCustPrpMgr.Get4(ширинаГраничнойРамкиName, true, out valOut,
                                                out ширинаГраничнойРамки);
                                            swCustProp.Set(ширинаГраничнойРамкиName, ширинаГраничнойРамки);
                                            confiData.ШиринаГраничнойРамки = ширинаГраничнойРамки;

                                            string толщинаЛистовогоМеталла;
                                            swCustPrpMgr.Get4(толщинаЛистовогоМеталлаNAme, true, out valOut,
                                                out толщинаЛистовогоМеталла);
                                            swCustProp.Set(толщинаЛистовогоМеталлаNAme, толщинаЛистовогоМеталла);
                                            confiData.ТолщинаЛистовогоМеталла = толщинаЛистовогоМеталла;

                                            string сгибы;
                                            swCustPrpMgr.Get4(сгибыName, true, out valOut, out сгибы);
                                            swCustProp.Set(сгибыName, сгибы);
                                            confiData.Сгибы = сгибы;

                                            string площадьПоверхности;
                                            swCustPrpMgr.Get4("Площадь поверхности", true, out valOut,
                                                out площадьПоверхности);
                                            swCustProp.Set(площадьПокрытияName, площадьПоверхности);
                                            confiData.ПлощадьПокрытия = Convert.ToDouble(площадьПоверхности);
                                        }
                                        swSubFeat = swSubFeat.GetNextFeature();
                                    }
                                }
                                swFeat2 = swFeat2.GetNextFeature();
                            }
                        }
                        dataList.Add(confiData);
                    }
                }
                catch (Exception exception)
                {
                    LoggerMine.Error(exception.ToString(), "Строка 377", "");
                }

                try
                {
                    swModel.ShowConfiguration2(activeconfiguration.Name);

                }
                catch (Exception exception)
                {
                    LoggerMine.Error(exception.ToString(), "", "");
                    isErrors = true;
                    LoggerMine.Info("isErrors = true на 392-й строке ", "", "");
                }

                try
                {
                    ExportDataToXmlSql(swModel, dataList);
                }
                catch (Exception exception)
                {
                    LoggerMine.Error(exception.ToString(), "", "");
                    isErrors = true;
                    LoggerMine.Info("isErrors = true на 403-й строке ", "", "");
                }

            #endregion
                
                #region Сохранение детали в eDrawing

                if (makeEprt)
                {
                    string modelName;
                    try
                    {
                        modelName = Path.GetFileNameWithoutExtension(swModel.GetPathName());
                        _eDrwFileName = Path.GetDirectoryName(swModel.GetPathName()) + "\\" + modelName + ".eprt";
                        eDrwFileName = _eDrwFileName;
                    }
                    catch (Exception exception)
                    {
                        LoggerMine.Error(exception.ToString(), "", "");
                        isErrors = true;
                        LoggerMine.Info("isErrors = true на 423-й строке ", "", "");
                        return;
                    }

                    try
                    {
                        // todo: удаление документов перед новым сохранением. Осуществить поиск по имени
                        var existingDocument = SearchDoc(modelName + ".eprt", SwDocType.SwDocNone);

                        if (existingDocument != "")
                        {
                            LoggerMine.Info($"Файл есть в базе {modelName} и будет удален.. ", "", "CreateFlattPatternUpdateCutlistAndEdrawing");
                            //CheckInOutPdm(new List<FileInfo> { new FileInfo(eDrwFileName) }, false, PdmBaseName);
                            DeleteFileFromPdm(existingDocument, PdmBaseName);
                        }
                        else
                        {
                            File.Delete(_eDrwFileName);
                        }
                    }
                    catch (Exception)
                    {
                        try
                        {
                            File.Delete(_eDrwFileName);
                        }
                        catch (Exception exception)
                        {
                            LoggerMine.Error(exception.StackTrace, "", "File.Delete(_eDrwFileName);");
                            isErrors = true;
                            LoggerMine.Info("isErrors = true на 453-й строке ", "", "");
                        }
                    }

                    #region ToDelete
                    //if (new FileInfo(_eDrwFileName).Exists)
                    //{
                    //    LoggerMine.Info("Файл есть в базе " + swModel.GetTitle());
                    //    //CheckInOutPdm(new List<FileInfo> { new FileInfo(eDrwFileName) }, false, PdmBaseName);
                    //    DeleteFileFromPdm(_eDrwFileName, PdmBaseName);
                    //}
                    #endregion

                    try
                    {
                        swApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swEdrawingsSaveAsSelectionOption,
                        (int)swEdrawingSaveAsOption_e.swEdrawingSaveAll);
                        swApp.SetUserPreferenceToggle(((int)(swUserPreferenceToggle_e.swEDrawingsOkayToMeasure)), true);
                        swModel.Extension.SaveAs(_eDrwFileName, (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                                (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, 0, 0);
                    }
                    catch (Exception exception)
                    {
                        LoggerMine.Error(exception.StackTrace, "", "File.Delete(_eDrwFileName);");
                        isErrors = true;
                        LoggerMine.Info("isErrors = true на 478-й строке ", "", "");
                    }

                    #region CreateFlattPatternUpdateCutlistAndEdrawing

                    //try
                    //{
                    //    // swApp.CloseDoc(Path.GetFileName(swModel.GetPathName()));
                    //    swApp.CloseDoc(modelName + ".sldprt");

                    //    if (makeDxf)
                    //    {
                    //        swApp.ExitApp();
                    //        swApp = null;
                    //    }
                    //    LoggerMine.Info("Обработка файла " + modelName + ".sldprt" + " успешно завершена", "", "CreateFlattPatternUpdateCutlistAndEdrawing");
                    //    isErrors = false;
                    //}
                    //catch (Exception exception)
                    //{
                    //    LoggerMine.Error(exception.StackTrace, "", "File.Delete(_eDrwFileName);");
                    //    isErrors = true;
                    //    LoggerMine.Info("isErrors = true на 497-й строке ", "", "");
                    //}

                    #endregion
                }

                #endregion

                try
                {
                    // swApp.CloseDoc(Path.GetFileName(swModel.GetPathName()));
                    var namePrt = swApp.IActiveDoc2.GetTitle().ToLower().Contains(".sldprt")
                        ? swApp.IActiveDoc2.GetTitle()
                        : swApp.IActiveDoc2.GetTitle() + ".sldprt";
                    swApp.CloseDoc(namePrt);

                    if (makeDxf)
                    {
                        swApp.ExitApp();
                        swApp = null;
                    }
                    LoggerMine.Info(
                        "Обработка файла " + swApp.IActiveDoc2.GetTitle() + ".sldprt" + ".sldprt" + " успешно завершена",
                        "", "CreateFlattPatternUpdateCutlistAndEdrawing");
                    isErrors = false;
                }
                catch (Exception exception)
                {
                    LoggerMine.Error(exception.StackTrace, "", "File.Delete(_eDrwFileName);");
                    isErrors = true;
                    LoggerMine.Info("isErrors = true на 497-й строке ", "", "");
                }
                finally
                {
                    //try
                    //{
                    //    swApp.CloseDoc(swApp.IActiveDoc2.GetTitle() + ".sldprt");
                    //}
                    //catch (Exception)
                    //{
                    //    swApp.CloseDoc(Path.GetFileName(swModel.GetPathName()));
                    //}
                }
            }
            catch (Exception exception)
            {
                LoggerMine.Error(
                    $"Общая ошибка метода: {exception.ToString()} Строка: {exception.StackTrace} exception.Source - ", exception.Source, "CreateFlattPatternUpdateCutlistAndEdrawing");
                if (swApp == null) return;
                if (makeDxf)
                {
                    swApp.ExitApp();
                }
                isErrors = true;
                LoggerMine.Info("isErrors = true на 506-й строке ", "", "");
            }
        }

        /// <summary>
        /// Creates the eprt.
        /// </summary>
        /// <param name="filePath">The file path.</param>
        /// <param name="eDrwFileName">Name of the e DRW file.</param>
        /// <param name="isErrors">if set to <c>true</c> [is errors].</param>
        /// <param name="makeEprt">if set to <c>true</c> [make eprt].</param>
        /// <param name="version">The version.</param>
        public void CreateEprt(string filePath, out string eDrwFileName, out bool isErrors, bool makeEprt,out int version)
        {
            isErrors = false;
            eDrwFileName = "";
            version = 0;

            SldWorks swApp = null;
            try
            {
                LoggerMine.Info("Запущен метод для обработки детали по пути " + filePath, "", "CreateFlattPatternUpdateCutlistAndEdrawing");

                var vault1 = new EdmVault5();
                vault1.LoginAuto(PdmBaseName, 0);
                
                try
                {
                    IEdmFolder5 oFolder;
                    var edmFile5 = vault1.GetFileFromPath(filePath, out oFolder);
                    edmFile5.GetFileCopy(0, 0, oFolder.ID, (int)EdmGetFlag.EdmGet_Simple);
                    _currentVersion = edmFile5.CurrentVersion;
                }
                catch (Exception exception)
                {
                    LoggerMine.Error(
                        $"Ошибка при получении значения последней версии файла {Path.GetFileName(filePath)}", exception.ToString(), "CreateFlattPatternUpdateCutlistAndEdrawing");
                }
                
                try
                {
                    swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
                }
                catch (Exception)
                {
                    swApp = new SldWorks { Visible = true };
                }
                if (swApp == null)
                {
                    isErrors = true;
                    return;
                }

                IModelDoc2 swModel;

                try
                {
                    swModel = swApp.OpenDoc6(filePath, (int)swDocumentTypes_e.swDocPART,
                                (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", 0, 0);
                    swModel.Extension.ViewDisplayRealView = false;
                    swModel.EditRebuild3();
                    swModel.ForceRebuild3(false);
                }
                catch (Exception exception)
                {
                    LoggerMine.Error(string.Format("Ошибка при обработке детали {2}: {0} Строка: {1}", exception.ToString(), exception.StackTrace, Path.GetFileName(filePath)), exception.Source, "CreateFlattPatternUpdateCutlistAndEdrawing");
                    isErrors = true;
                    return;
                }

                try
                {
                    if (!IsSheetMetalPart((IPartDoc)swModel))
                    {
                        swApp.CloseDoc(Path.GetFileName(swModel.GetPathName()));
                        swApp.ExitApp();
                        swApp = null;
                        return;
                    }
                }
                catch (Exception exception)
                {
                    LoggerMine.Error(string.Format("Ошибка2 при обработке детали {2}: {0} Строка: {1}", exception.ToString(), exception.StackTrace, Path.GetFileName(filePath)), exception.Source, "CreateFlattPatternUpdateCutlistAndEdrawing");
                    isErrors = true;
                }

                if (makeEprt)
                {
                    #region Сохранение детали в eDrawing

                    var modelName = Path.GetFileNameWithoutExtension(swModel.GetPathName());
                        _eDrwFileName = Path.GetDirectoryName(swModel.GetPathName()) + "\\" + modelName + ".eprt";
                        eDrwFileName = _eDrwFileName;
                        try
                        {
                            IEdmFolder5 oFolder;
                            var edmFile5 = vault1.GetFileFromPath(_eDrwFileName, out oFolder);
                            if (oFolder == null)
                            {
                            //MessageBox.Show("Версия");
                            goto m1;
                            }
                            edmFile5.GetFileCopy(0, 0, oFolder.ID, (int) EdmGetFlag.EdmGet_Simple);
                            var existVer = 0;
                            var pEnumVar = (IEdmEnumeratorVariable8) edmFile5.GetEnumeratorVariable();
                            try
                            {
                                object currentVersionEPrt;
                                pEnumVar.GetVar("Revision", "", out currentVersionEPrt);
                                int.TryParse(Convert.ToString(currentVersionEPrt), out existVer);
                            }
                            catch (Exception)
                            {
                              //MessageBox.Show(exception.ToString(), "Revision");
                            }
                        if (existVer == _currentVersion)
                        {
                            //MessageBox.Show( Convert.ToString(existVer), Convert.ToString(_currentVersion));
                            goto m2;
                        }
                    }
                    catch (Exception exception)
                    {
                        LoggerMine.Error(exception.ToString(), "", "");
                        isErrors = true;
                        LoggerMine.Info("isErrors = true на 423-й строке ", "", "");
                        //return;
                        goto m2;
                    }

                    #region Delete(_eDrwFileName)

                        try
                        {
                            // todo: удаление документов перед новым сохранением. Осуществить поиск по имени
                            var existingDocument = SearchDoc(modelName + ".eprt", SwDocType.SwDocNone);
                            if (existingDocument != "")
                            {
                                LoggerMine.Info($"Файл есть в базе {modelName} и будет удален.. ", "", "CreateFlattPatternUpdateCutlistAndEdrawing");
                                //CheckInOutPdm(new List<FileInfo> { new FileInfo(eDrwFileName) }, false, PdmBaseName);
                                DeleteFileFromPdm(existingDocument, PdmBaseName);
                            }
                            else
                            {
                                File.Delete(_eDrwFileName);
                            }
                        }
                        catch (Exception)
                        {
                            try
                            {
                                File.Delete(_eDrwFileName);
                            }
                            catch (Exception exception)
                            {
                                LoggerMine.Error(exception.StackTrace, "", "File.Delete(_eDrwFileName);");
                                isErrors = true;
                                LoggerMine.Info("isErrors = true на 453-й строке ", "", "");
                            }
                        }

                    #endregion

                    #region ToDelete
                    //if (new FileInfo(_eDrwFileName).Exists)
                    //{
                    //    LoggerMine.Info("Файл есть в базе " + swModel.GetTitle());
                    //    //CheckInOutPdm(new List<FileInfo> { new FileInfo(eDrwFileName) }, false, PdmBaseName);
                    //    DeleteFileFromPdm(_eDrwFileName, PdmBaseName);
                    //}
                    #endregion
                          m1:
                    try
                    {
                        if (swApp != null)
                        {
                            swApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swEdrawingsSaveAsSelectionOption,
                                (int)swEdrawingSaveAsOption_e.swEdrawingSaveAll);
                            swApp.SetUserPreferenceToggle(((int)(swUserPreferenceToggle_e.swEDrawingsOkayToMeasure)), true);
                        }
                        swModel.Extension.SaveAs(_eDrwFileName, (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                                (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, 0, 0);
                        version = _currentVersion;
                    }
                    catch (Exception exception)
                    {
                        LoggerMine.Error(exception.StackTrace, "", "File.Delete(_eDrwFileName);");
                        isErrors = true;
                    }

                    #endregion
                }
           
                m2:
                try
                {
                    if (swApp != null)
                    {
                        swApp.CloseDoc(swApp.IActiveDoc2.GetTitle() + ".sldprt");
                        swApp.ExitApp();
                        LoggerMine.Info("Обработка файла " + swApp.IActiveDoc2.GetTitle() + ".sldprt" + ".sldprt" + " успешно завершена", "", "CreateFlattPatternUpdateCutlistAndEdrawing");
                    }
                    isErrors = false;
                }
                catch (Exception exception)
                {
                    LoggerMine.Error(exception.StackTrace, "", "File.Delete(_eDrwFileName);");
                    swApp.ExitApp();
                    isErrors = true;
                }

            }
            catch (Exception exception)
            {
                LoggerMine.Error(
                    $"Общая ошибка метода: {exception.ToString()} Строка: {exception.StackTrace} exception.Source - ", exception.Source, "CreateFlattPatternUpdateCutlistAndEdrawing");
                if (swApp == null) return;
                isErrors = true;
            }
        }

        /// <summary>
        /// Registrations in PDM.
        /// </summary>
        /// <param name="newFileName">Name of the eDRW file.</param>
        public bool RegistrationPdm(string newFileName)
        {
            return CheckInOutPdm(new List<FileInfo> { new FileInfo(newFileName) }, true, PdmBaseName);
        }

        /// <summary>
        /// Registrations the PDM version.
        /// </summary>
        /// <param name="newFileName">New name of the file.</param>
        /// <param name="version">The version.</param>
        /// <returns></returns>
        public bool RegistrationPdmVersion(string newFileName, int version)
        {
            try
            {
                var vault1 = new EdmVault5();
                vault1.LoginAuto(PdmBaseName, 0);
                IEdmFolder5 oFolder;
                var wait = 1;
                m1:
                Thread.Sleep(500);
                var edmFile5 = vault1.GetFileFromPath(newFileName, out oFolder);
                if (oFolder == null)
                {
                    wait++;
                    if (wait == 20) goto m2;
                    goto m1;
                }

                m2:
                if (oFolder != null) edmFile5.GetFileCopy(0, 0, oFolder.ID, (int)EdmGetFlag.EdmGet_Simple);
                var pEnumVar = (IEdmEnumeratorVariable8)edmFile5.GetEnumeratorVariable();
                pEnumVar.SetVar("Revision", "", Convert.ToString(version));
                pEnumVar.CloseFile(true);
            }
            catch (Exception)
            {
                //MessageBox.Show(exception.ToString());
            }
            return CheckInOutPdm(new List<FileInfo> { new FileInfo(newFileName) }, true, PdmBaseName);
        }

        #region Additional Methods

        static void DeleteFileFromPdm(string filePath, string pdmBase)
        {
            LoggerMine.Info($"Удаление файла по пути {filePath} базе PDM - {pdmBase}", "", "DeleteFileFromPdm");

            var retryCount = 2;
            var success = false;
            var ex = new Exception();
            while (!success && retryCount > 0)
            {
                try
                {
                    var vault1 = new EdmVault5();
                    IEdmFolder5 oFolder;
                    vault1.LoginAuto(pdmBase, 0);
                   // vault1.Login("kb81", "1", pdmBase);

                    vault1.GetFileFromPath(filePath, out oFolder);

                    var vault2 = (IEdmVault7)vault1;
                    var batchDeleter = (IEdmBatchDelete3)vault2.CreateUtility(EdmUtility.EdmUtil_BatchDelete);
                    batchDeleter.AddFileByPath(filePath);
                    batchDeleter.ComputePermissions(true);
                    batchDeleter.CommitDelete(0);
                    //LoggerMine.Info(string.Format("batchDeleter.CommitDelete - {0}", commitDelete));

                    LoggerMine.Info(string.Format("В базе PDM - {1}, удален файл по пути {0}", filePath, pdmBase), "", "DeleteFileFromPdm");

                    success = true;
                }
                catch (Exception exception)
                {
                    retryCount--;
                    ex = exception;
                    Thread.Sleep(200);
                    if (retryCount == 0)
                    {
                        // throw; //or handle error and break/return
                    }
                    LoggerMine.Error(
                        $"Во время удаления по пути {filePath} возникла ошибка. База - {pdmBase}. Ошибка: {exception.ToString()} Строка: {exception.StackTrace}", exception.Source, "DeleteFileFromPdm");
                }
            }
            if (!success)
            {
                LoggerMine.Error(
                    $"Во время удаления по пути {filePath} возникла ошибка. База - {pdmBase}. Ошибка: {ex.ToString()} Строка: {ex.StackTrace}", ex.Source, "DeleteFileFromPdm");
            }
        }

        static bool CheckInOutPdm(IEnumerable<FileInfo> filesList, bool registration, string pdmBase)
        {
            if (filesList.Count() == 1)
            {
                Thread.Sleep(10000);
            }
            foreach (var file in filesList)
            {
                var retryCount = 2;
                var success = false;
                var ex = new Exception();
                while (!success && retryCount > 0)
                {
                    try
                    {
                        var vault1 = new EdmVault5();
                        IEdmFolder5 oFolder;
                        vault1.LoginAuto(pdmBase, 0);

                        //vault1.Login("kb81","1",pdmBase);
                        var b = 1;
                        m1:
                        var edmFile5 = vault1.GetFileFromPath(file.FullName, out oFolder);
                        if (oFolder == null)
                        {
                            b++;
                            if (b == 10) goto m2;
                            goto m1;
                        }
                        m2:
                        LoggerMine.Info(string.Format("Хранилище - {1}, файл {2} по пути {0}", file.FullName, pdmBase, edmFile5.Name), "", "CheckInOutPdm");
                        // Разрегистрировать
                        if (registration == false)
                        {
                            if (oFolder != null)
                            {
                                edmFile5.GetFileCopy(0, 0, oFolder.ID, (int)EdmGetFlag.EdmGet_Simple);
                                edmFile5.LockFile(oFolder.ID, 0);
                            }
                        }
                        // Зарегистрировать
                        if (registration)
                        {
                            Thread.Sleep(50);
                            if (edmFile5.IsLocked)
                                if (oFolder != null) edmFile5.UnlockFile(oFolder.ID, "");
                            Thread.Sleep(50);
                        }

                        LoggerMine.Info(string.Format("В хранилище - {1}, зарегестрирован документ по пути {0}", file.FullName, pdmBase), "", "CheckInOutPdm");

                        success = true;
                    }
                    catch (Exception exception)
                    {
                        retryCount--;
                        ex = exception;
                        Thread.Sleep(200);
                        if (retryCount == 0)
                        {
                            // throw; //or handle error and break/return
                        }
                        LoggerMine.Error(string.Format("Во время регистрации документа по пути {0} возникла ошибка{3}\nБаза - {1}. {2}", file.FullName, pdmBase, exception.ToString(), exception.StackTrace), exception.TargetSite.Name, "CheckInOutPdm");
                        return false;
                    }
                }
                if (success) continue;
                LoggerMine.Error(
                    $"Во время регистрации документа по пути {file.FullName} возникла ошибка\nБаза - {pdmBase}. {ex.ToString()}", ex.TargetSite.Name, "CheckInOutPdm");
                return false;
            }
            return true;
        }

        static void AddFileToPdm(string path, string pdmBase)
        {
            try
            {
                LoggerMine.Info($"Создание папки по пути {path} для сохранения", "", "AddFileToPdm");
                var vault1 = new EdmVault5();
                if (!vault1.IsLoggedIn)
                {
                    vault1.LoginAuto(pdmBase, 0);
                   // vault1.Login("kb81", "1", pdmBase);
                }

                var vault2 = (IEdmVault7)vault1;
                var fileDirectory = new FileInfo(path).DirectoryName;
                var fileFolder = vault2.GetFolderFromPath(fileDirectory);
                var result = fileFolder.AddFile(fileFolder.ID, "", Path.GetFileName(path));
                LoggerMine.Info(string.Format("Создание файла по пути {0} в папке {2} завершено. {1}", path, result, fileFolder.Name), "", "AddFileToPdm");
            }
            catch (Exception exception)
            {
                LoggerMine.Error(string.Format("Не удалось создать файл по пути {0}. Ошибка: {2} Строка: {1}", path, exception.StackTrace, exception.ToString()), exception.TargetSite.Name, "AddFileToPdm");
            }
        }
        
        static string GetFromCutlist(IModelDoc2 swModel, string property)
        {
            LoggerMine.Info(string.Format("Получение свойства '{1}' из CutList'а для {0}. Имя конфигурации '{2}'", new FileInfo(swModel.GetPathName()).Name, property, swModel.IGetActiveConfiguration().Name), "", "GetFromCutlist");

            var propertyValue = "";

            try
            {
                Feature swFeat2 = swModel.FirstFeature();
                while (swFeat2 != null)
                {
                    if (swFeat2.GetTypeName2() == "SolidBodyFolder")
                    {
                        BodyFolder swBodyFolder = swFeat2.GetSpecificFeature2();
                        swFeat2.Select2(false, -1);
                        swBodyFolder.SetAutomaticCutList(true);
                        swBodyFolder.UpdateCutList();

                        Feature swSubFeat = swFeat2.GetFirstSubFeature();
                        while (swSubFeat != null)
                        {
                            if (swSubFeat.GetTypeName2() == "CutListFolder")
                            {
                                BodyFolder bodyFolder = swSubFeat.GetSpecificFeature2();
                                swSubFeat.Select2(false, -1);
                                bodyFolder.SetAutomaticCutList(true);
                                bodyFolder.UpdateCutList();
                                var swCustPrpMgr = swSubFeat.CustomPropertyManager;
                                //swCustPrpMgr.Add("Площадь поверхности", "Текст", "\"SW-SurfaceArea@@@Элемент списка вырезов1@ВНС-901.81.002.SLDPRT\"");
                                string valOut;
                                swCustPrpMgr.Get4(property, true, out valOut, out propertyValue);
                            }
                            swSubFeat = swSubFeat.GetNextFeature();
                        }
                    }
                    swFeat2 = swFeat2.GetNextFeature();
                }
                LoggerMine.Info("Метод GetFromCutlist() для " + new FileInfo(swModel.GetPathName()).Name + " завершен.", "", "GetFromCutlist");
            }
            catch (Exception exception)
            {
                LoggerMine.Error(string.Format("Во время получение свойства возникла ошибка: '{1}' в строке: {0}. Сообщение: '{2}'", exception.Source, exception.StackTrace, exception.ToString()), "", "GetFromCutlist");
            }
            
           
            return propertyValue;
        }

        static bool IsSheetMetalPart(IPartDoc swPart)
        {
            var mod = (IModelDoc2) swPart;

            LoggerMine.Info("Проверка на листовую деталь " + mod.GetTitle(), "", "IsSheetMetalPart");
            try
            {
                var isSheet = false;

                var vBodies = swPart.GetBodies2((int)swBodyType_e.swSolidBody, false);

                foreach (Body2 vBody in vBodies)
                {
                    try
                    {
                        var isSheetMetal = vBody.IsSheetMetal();
                        if (!isSheetMetal) continue;
                        isSheet = true;
                    }
                    catch
                    {
                        isSheet = false;
                    }
                }

                LoggerMine.Info(
                    $"Проверка детали {mod.GetTitle()} завершена. Она {(isSheet ? "листовая" : "не листовая")}.", "", "IsSheetMetalPart");
                return isSheet;
            }
            catch (Exception)
            {
                LoggerMine.Info("Проверка завершена. Деталь не из листового материала.", "", "IsSheetMetalPart");
                // var swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
                // swApp.ExitApp();
                return false;
            }
        }

        #endregion
        
        #region Data to export

        class DataToExport
        {
            public string Config;
            public string Материал;
            public string Обозначение;
            public double ПлощадьПокрытия;
            public string КодМатериала;
            
            public string ДлинаГраничнойРамки;
            public string ШиринаГраничнойРамки;
            public string Сгибы;
            public string ТолщинаЛистовогоМеталла;
            public string Наименование;

            public int? MaterialId;

            public string FileName;
            
            public double? PaintX;
            public double? PaintY;
            public double? PaintZ;

            public int IdPdm;

            
        }
        
        void ExportDataToXmlSql(IModelDoc2 swModel, IEnumerable<DataToExport> dataToExport)
        {
            if (swModel == null || dataToExport == null)
            {
                // ReSharper disable once ConditionIsAlwaysTrueOrFalse
                // ReSharper disable ConditionIsAlwaysTrueOrFalse
                LoggerMine.Error("Попытка запуска ExportDataToXmlSql() с пустыми параметрами. " + swModel == null ? "swModel = null" : "" + dataToExport == null ? "dataToExport = null" : "", "", "ExportDataToXmlSql");
                // ReSharper restore ConditionIsAlwaysTrueOrFalse
                return;
            }

            LoggerMine.Info("Выгрузка данных в XML файл и SQL базу по детали " + new FileInfo(swModel.GetPathName()).Name, "", "ExportDataToXmlSql");
            
            try
            {
                //var myXml = new System.Xml.XmlTextWriter(@"\\srvkb\SolidWorks Admin\XML\" + swModel.GetTitle() + ".xml", System.Text.Encoding.UTF8);
                //const string xmlPath = @"\\srvkb\SolidWorks Admin\XML\";
                //const string xmlPath = @"C:\Temp\";

                var myXml = new System.Xml.XmlTextWriter(_xmlPath + Path.GetFileNameWithoutExtension(swModel.GetPathName()) + ".xml", System.Text.Encoding.UTF8);

                myXml.WriteStartDocument();
                myXml.Formatting = System.Xml.Formatting.Indented;
                myXml.Indentation = 2;

                // создаем элементы
                myXml.WriteStartElement("xml");
                myXml.WriteStartElement("transactions");
                myXml.WriteStartElement("transaction");

                myXml.WriteStartElement("document");

                foreach (var configData in dataToExport)
                {
                    #region XML

                    // Конфигурация
                    myXml.WriteStartElement("configuration");
                    myXml.WriteAttributeString("name", configData.Config);

                    // Материал
                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "Материал");
                    myXml.WriteAttributeString("value", configData.Материал);
                    myXml.WriteEndElement();

                    // Наименование  -- Из таблицы свойств
                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "Наименование");
                    myXml.WriteAttributeString("value", configData.Наименование);
                    myXml.WriteEndElement();

                    // Обозначение
                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "Обозначение");
                    myXml.WriteAttributeString("value", configData.Обозначение);
                    myXml.WriteEndElement();

                    // Площадь покрытия
                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "Площадь покрытия");
                    myXml.WriteAttributeString("value", Convert.ToString(configData.ПлощадьПокрытия).Replace(",","."));
                    myXml.WriteEndElement();

                    // ERP code
                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "Код_Материала");
                    myXml.WriteAttributeString("value", configData.КодМатериала);
                    myXml.WriteEndElement();

                    // Длина граничной рамки

                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "Длина граничной рамки");
                    myXml.WriteAttributeString("value", configData.ДлинаГраничнойРамки);
                    myXml.WriteEndElement();

                    // Ширина граничной рамки
                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "Ширина граничной рамки");
                    myXml.WriteAttributeString("value", configData.ШиринаГраничнойРамки);
                    myXml.WriteEndElement();

                    // Сгибы
                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "Сгибы");
                    myXml.WriteAttributeString("value", configData.Сгибы);
                    myXml.WriteEndElement();

                    // Толщина листового металла
                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "Толщина листового металла");
                    myXml.WriteAttributeString("value", configData.ТолщинаЛистовогоМеталла);
                    myXml.WriteEndElement();

                    // Версия последняя
                    myXml.WriteStartElement("attribute");
                    myXml.WriteAttributeString("name", "Версия");
                    myXml.WriteAttributeString("value", Convert.ToString(_currentVersion));
                    myXml.WriteEndElement();

                    myXml.WriteEndElement();  //configuration

                    #endregion

                    #region SQL

                    try
                    {
                        var sqlConnection = new SqlConnection(_connectionString);
                        sqlConnection.Open();
                        var spcmd = new SqlCommand("UpDateCutList", sqlConnection) { CommandType = CommandType.StoredProcedure };
                        
                        double workpieceX; double.TryParse(configData.ДлинаГраничнойРамки.Replace('.', ','), out workpieceX);
                        double workpieceY; double.TryParse(configData.ШиринаГраничнойРамки.Replace('.', ','), out workpieceY);
                        int bend; int.TryParse(configData.Сгибы, out bend);
                        double thickness; double.TryParse(configData.ТолщинаЛистовогоМеталла.Replace('.', ','), out thickness);
                            
                        var configuration = configData.Config;

                        var materialId = configData.MaterialId;

                        if (materialId == null)
                        {
                            spcmd.Parameters.AddWithValue("@MaterialID", DBNull.Value);
                        }
                        else
                        {
                            spcmd.Parameters.AddWithValue("@MaterialID", materialId);
                        }
                        
                        spcmd.Parameters.AddWithValue("@PaintX", configData.PaintX);
                        spcmd.Parameters.AddWithValue("@PaintY", configData.PaintY);
                        spcmd.Parameters.AddWithValue("@PaintZ", configData.PaintZ);

                        spcmd.Parameters.AddWithValue("@FILENAME", configData.FileName);
                        spcmd.Parameters.AddWithValue("@IDPDM", configData.IdPdm);

                        #region
                        //double surfaceArea;
                        //double.TryParse(configData.ПлощадьПокрытия.Replace(",", "."), out surfaceArea);

                        //if (Math.Abs(surfaceArea) > 0.1)
                        //{
                        #endregion
                        
                        //MessageBox.Show("SurfaceArea - " + ParseDouble(configData.ПлощадьПокрытия.ToString()));
                        //MessageBox.Show("SurfaceArea - CurrentCulture " + ParseDouble(configData.ПлощадьПокрытия.ToString(CultureInfo.CurrentCulture)));
                        //MessageBox.Show("SurfaceArea InvariantCulture - " + ParseDouble(configData.ПлощадьПокрытия.ToString(CultureInfo.InvariantCulture)));

                        spcmd.Parameters.AddWithValue("@SurfaceArea", ParseDouble(configData.ПлощадьПокрытия.ToString()));//configData.ПлощадьПокрытия);

                        #region
                        //}
                        //else
                        //{
                        //    spcmd.Parameters.AddWithValue("@SurfaceArea", DBNull.Value);
                        //}
                        #endregion

                        spcmd.Parameters.Add("@WorkpieceX", SqlDbType.Float).Value = workpieceX;
                        spcmd.Parameters.Add("@WorkpieceY", SqlDbType.Float).Value = workpieceY;
                        
                        spcmd.Parameters.Add("@Bend", SqlDbType.Int).Value = bend;
                        spcmd.Parameters.Add("@Thickness", SqlDbType.Float).Value = thickness;
                        spcmd.Parameters.Add("@Configuration", SqlDbType.NVarChar).Value = configuration;
                        spcmd.Parameters.Add("@version", SqlDbType.Int).Value = _currentVersion;

                        //MessageBox.Show("configuration - " + configuration + "\nmaterialId - " + materialId
                        //    + "\nPaintX - " + configData.PaintX + "\nPaintY - " + configData.PaintY + "\nPaintY - " + configData.PaintZ
                        //    + "\nFILENAME - " + configData.FileName + "\nIDPDM - " + configData.IdPdm
                        //    + "\nWorkpieceX - " + configData.ДлинаГраничнойРамки + "\nWorkpieceY - " + configData.ШиринаГраничнойРамки
                        //    + "\nBend - " + configData.Сгибы + "\nThickness - " + configData.ТолщинаЛистовогоМеталла
                        //    + "\nConfiguration - " + configuration + "\nversion - " + _currentVersion
                        //    + "\nSurfaceArea - " + configData.ПлощадьПокрытия);


                        spcmd.ExecuteNonQuery();
                        sqlConnection.Close();


                    }
                    catch (Exception exception)
                    {
                        //MessageBox.Show(exception.ToString());
                        LoggerMine.Error(string.Format("Ошибка: {1} Строка: {0}", exception.StackTrace, exception.ToString()), exception.TargetSite.Name, "ExportDataToXmlSql");
                    }

                    #endregion
                }

                //myXml.WriteEndElement();// ' элемент CONFIGURATION
                myXml.WriteEndElement();// ' элемент DOCUMENT
                myXml.WriteEndElement();// ' элемент TRANSACTION
                myXml.WriteEndElement();// ' элемент TRANSACTIONS
                myXml.WriteEndElement();// ' элемент XML
                // заносим данные в myMemoryStream
                myXml.Flush();

                myXml.Close();


                //MessageBox.Show("Выгрузка данных для детали " + swModel.GetTitle() + " завершена.", "ExportDataToXmlSql");
                LoggerMine.Info("Выгрузка данных для детали " + swModel.GetTitle() + " завершена.", "", "ExportDataToXmlSql");
            }
            catch (Exception exception)
            {
                //MessageBox.Show(string.Format("Ошибка: {1} Строка: {0}", exception.StackTrace, exception.ToString()), "ExportDataToXmlSql");
                LoggerMine.Error(string.Format("Ошибка: {1} Строка: {0}", exception.StackTrace, exception.ToString()), exception.TargetSite.Name, "ExportDataToXmlSql");
            }
        }

        #endregion

        /// <summary>
        /// 
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static double ParseDouble(string value)
        {
            value = value.Replace(" ", "");
            if (CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator == ",")
            {
                value = value.Replace(".", ",");
            }
            else
            {
                value = value.Replace(",", ".");
            }
            string[] splited = value.Split(CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator[0]);
            if (splited.Length > 2)
            {
                string r = "";
                for (int i = 0; i < splited.Length; i++)
                {
                    if (i == splited.Length - 1)
                        r += CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;
                    r += splited[i];
                }
                value = r;
            }
            return double.Parse(value);
        }
        

        #region Save As Pdf

        /// <summary>
        /// Saves the DRW as PDF.
        /// </summary>
        /// <param name="filePath">The file path.</param>
        /// <param name="pdfFilePath">The PDF file path.</param>
        public void SaveDrwAsPdf(string filePath, out string pdfFilePath)
        {
            pdfFilePath = "";
            
            //SearchDoc(Path.GetFileNameWithoutExtension(filePath), SwDocType.SwDocDrawing);
            
            SldWorks swApp = null;
            try
            {
                LoggerMine.Info(" Запущен метод сохранения чертежа для " + filePath, "", "SaveDrwAsPdf");

                var vault1 = new EdmVault5();
                IEdmFolder5 oFolder;
                vault1.LoginAuto(PdmBaseName, 0);

                var edmFile5 = vault1.GetFileFromPath(filePath, out oFolder);
                edmFile5.GetFileCopy(0, 0, oFolder.ID, (int)EdmGetFlag.EdmGet_Simple);
                
                try
                {
                    swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
                }
                catch (Exception)
                {
                    swApp = new SldWorks { Visible = true };
                }
                if (swApp == null) { return; }

               
                var swModel = swApp.OpenDoc6(filePath, (int)swDocumentTypes_e.swDocDRAWING,
                                (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "00", 0, 0);
                swModel.Extension.ViewDisplayRealView = false;
                var swDraw = (DrawingDoc)swModel;
                
                try
                {
                    swDraw.ResolveOutOfDateLightWeightComponents();
                    swDraw.ForceRebuild();

                    // Движение по листам
                    var vSheetName = (string[]) swDraw.GetSheetNames();

                    foreach (var name in vSheetName)
                    {
                        swDraw.ResolveOutOfDateLightWeightComponents();
                        var swSheet = swDraw.Sheet[name];
                        swDraw.ActivateSheet(swSheet.GetName());

                        if ((swSheet.IsLoaded()))
                        {
                            try
                            {
                                var sheetviews = (object[]) swSheet.GetViews();
                                var firstView = (View) sheetviews[0];
                                firstView.SetLightweightToResolved();
                                
                                var baseView =  firstView.IGetBaseView();
                                var dispData = (IModelDoc2)baseView.ReferencedDocument;

                            }
                            catch (Exception exception)
                            {
                                LoggerMine.Error(exception.StackTrace + "\n", "", "SaveDrwAsPdf");
                               // MessageToUsr = exception.StackTrace;
                            }
                        }
                        else
                        {
                            return;
                        }

                        //Движение по видам
                        //if (!deep) continue;
                        try
                        {
                            var views = (object[]) swSheet.GetViews();
                            foreach (var drwView in views.Cast<View>())
                            {
                                drwView.SetLightweightToResolved();
                            }
                        }
                        catch (Exception exception)
                        {
                            LoggerMine.Error(string.Format("Ошибка: {1} Строка: {0}", exception.StackTrace, exception.ToString()), exception.TargetSite.Name, "SaveDrwAsPdf");
                        }
                    }
                    
                    #region Saving New Doc (Delete Old)

                    var errors = 0;
                    var warnings = 0;
                    var newpath = Path.GetDirectoryName(swModel.GetPathName()) + "\\" + Path.GetFileNameWithoutExtension(swModel.GetPathName()) + ".pdf";
                    pdfFilePath = newpath;

                    if (new FileInfo(newpath).Exists)
                    {
                        LoggerMine.Info("Файл есть в базе " + Path.GetFileNameWithoutExtension(swModel.GetPathName()) + ".pdf", "", "SaveDrwAsPdf");
                        //CheckInOutPdm(new List<FileInfo> { new FileInfo(eDrwFileName) }, false, PdmBaseName);
                        DeleteFileFromPdm(newpath, PdmBaseName);
                    }

                    swApp.SetUserPreferenceToggle(((int)(swUserPreferenceToggle_e.swPDFExportInColor)), true);
                    swApp.SetUserPreferenceToggle(((int)(swUserPreferenceToggle_e.swPDFExportEmbedFonts)), true);
                    swApp.SetUserPreferenceToggle(((int)(swUserPreferenceToggle_e.swPDFExportUseCurrentPrintLineWeights)), true);

                    var canSave = swModel.Extension.SaveAs(newpath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, ref errors, ref warnings);
                    if (canSave)
                    {
                        DeleteFileFromPdm(newpath, PdmBaseName);
                        AddFileToPdm(newpath, PdmBaseName);
                        swModel.Extension.SaveAs(newpath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, ref errors, ref warnings);
                    }

                    swApp.CloseDoc(Path.GetFileName(new FileInfo(newpath).FullName));

                    #endregion
                }
                catch (Exception exception)
                {
                    LoggerMine.Error(exception + "\n", "", "SaveDrwAsPdf");
                }
                finally
                {
                    swApp.ExitApp();
                    swApp = null;
                    LoggerMine.Info("PDF Сохранен", "", "SaveDrwAsPdf");
                }
            }
            catch (Exception exception)
            {
                swApp?.ExitApp();
                LoggerMine.Error(exception + "\n", "", "SaveDrwAsPdf");
            }
        }

        #endregion

        #endregion
    }
}

#region To Delete

//void AddToPdmByPath(string path, string pdmBase)
//{
//    try
//    {
//        //if (Directory.Exists(path))
//        //{
//        //    return;
//        //}

//        LoggerMine.Info($"Создание папки по пути {path} для сохранения", "", "AddToPdmByPath");
//        var vault1 = new EdmVault5();
//        if (!vault1.IsLoggedIn)
//        {
//            vault1.LoginAuto(pdmBase, 0);
//        }

//        var vault2 = (IEdmVault7)vault1;
//        //try
//        //{
//        //    var directoryInfo = new DirectoryInfo(path);
//        //    if (directoryInfo.Parent == null) return;
//        //    var parentFolder = vault2.GetFolderFromPath(directoryInfo.Parent.FullName);
//        //    parentFolder.AddFolder(0, directoryInfo.Name);
//        //    LoggerMine.Info(string.Format("Создание папки по пути {0} завершено.", path));
//        //}
//        //catch (Exception)
//        //{
//        var fileDirectory = new FileInfo(path).DirectoryName;
//        var parentFolder = vault2.GetFolderFromPath(fileDirectory);


//        parentFolder.AddFile(parentFolder.ID, "", path);
//        //parentFolder.AddFolder(0, directoryInfo.Name);
//        LoggerMine.Info($"Создание файла по пути {path} завершено.", "", "AddToPdmByPath");
//        //}


//    }
//    catch (Exception exception)
//    {
//        LoggerMine.Error($"Не удалось создать папку по пути {path}. Ошибка {exception}", "", "AddToPdmByPath");
//    }
//}

//static void AddFilePdm(string path, string pdmBase)
//{
//    try
//    {
//        //if (Directory.Exists(path))
//        //{
//        //    return;
//        //}

//        LoggerMine.Info($"Создание папки по пути {path} для сохранения", "", "AddToPdmByPath");
//        var vault1 = new EdmVault5();
//        if (!vault1.IsLoggedIn)
//        {
//            vault1.LoginAuto(pdmBase, 0);
//        }
//        var vault2 = (IEdmVault7)vault1;
//        var fileInfo = new FileInfo(path);
//        if (fileInfo.Exists == false) return;
//        var epdmFile = vault2.GetFolderFromPath(fileInfo.FullName);
//        //    epdmFile.AddFolder(0, directoryInfo.Name);
//        //  edmFolder6.AddFile()
//    }
//    catch (Exception exception)
//    {
//        LoggerMine.Error($"Не удалось создать папку по пути {path}. Ошибка {exception}", "", "AddToPdmByPath");
//    }
//}




//Const EMPTY_DRAWING As String = "C:\EmptyDraw.SLDDRW"
//Const OUT_PATH As String = "\\srvkb\DXF"

//Dim swApp As SldWorks.SldWorks
//Dim swModel As SldWorks.ModelDoc2
//Dim swEmptyDraw As SldWorks.ModelDoc2
//Dim swPart As SldWorks.PartDoc

//Sub main()

//    Dim fileName As String
//    fileName = "<Filepath>"

//    Dim ext As String
//    ext = UCase(Right(fileName, 6))
//    If ext <> "SLDPRT" Then
//        Exit Sub
//    End If

//    Set swApp = Application.SldWorks

//    swApp.Visible = True
//    Set swModel = swApp.OpenDoc6(fileName, swDocumentTypes_e.swDocPART, swOpenDocOptions_e.swOpenDocOptions_Silent + swOpenDocOptions_e.swOpenDocOptions_ReadOnly, "", 0, 0)

//    If swModel Is Nothing Then
//        Exit Sub
//    End If

//    If swModel.GetType <> swDocumentTypes_e.swDocPART Then
//        Exit Sub
//    End If

//    Set swPart = swModel

//    If Not IsSheetMetalPart() Then
//        Exit Sub
//    End If

//    Set swEmptyDraw = swApp.OpenDoc6(EMPTY_DRAWING, swDocumentTypes_e.swDocDRAWING, swOpenDocOptions_e.swOpenDocOptions_Silent + swOpenDocOptions_e.swOpenDocOptions_ReadOnly, "", 0, 0)

//    If swEmptyDraw Is Nothing Then
//        Exit Sub
//    End If

//    Dim vConfNames As Variant
//    Dim i As Integer
//    Dim outFile As String

//    vConfNames = swModel.GetConfigurationNames

//    For i = 0 To UBound(vConfNames)

//        Dim swConf As SldWorks.Configuration
//        Set swConf = swModel.GetConfigurationByName(vConfNames(i))

//        If False = swConf.IsDerived Then

//            swModel.ShowConfiguration2 vConfNames(i)
//            swModel.ForceRebuild3 False
//            outFile = GetOutFileName(fileName, CStr(vConfNames(i)))

//            If False = swEmptyDraw.Extension.SaveAs(outFile, swSaveAsVersion_e.swSaveAsCurrentVersion, swSaveAsOptions_e.swSaveAsOptions_Silent, Nothing, 0, 0) Then
//                Exit Sub
//            End If
//            swPart.ExportFlatPatternView outFile, swExportFlatPatternViewOptions_e.swExportFlatPatternOption_RemoveBends
//        End If

//    Next

//    swApp.CloseDoc swEmptyDraw.GetTitle
//    swApp.CloseDoc swModel.GetTitle

//End Sub

//Function GetOutFileName(inputFile As String, confName As String) As String

//    Dim path As String
//    Dim name As String

//    path = Left(inputFile, InStrRev(inputFile, "\"))
//    name = Mid(inputFile, Len(path), Len(inputFile) - Len(path) - 6)

//    GetOutFileName = OUT_PATH + name + "-" + confName + ".dxf"

//End Function

//Function IsSheetMetalPart() As Boolean

//    Dim vBodies As Variant
//    Dim swBody As SldWorks.Body2

//    vBodies = swPart.GetBodies2(swBodyType_e.swSolidBody, False)

//    Dim i As Integer

//    For i = 0 To UBound(vBodies)

//        Set swBody = vBodies(i)

//        If swBody.IsSheetMetal Then
//            IsSheetMetalPart = True
//            Exit Function
//        End If

//    Next

//    IsSheetMetalPart = False

//End Function

#endregion