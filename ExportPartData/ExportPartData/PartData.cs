using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swdocumentmgr;
using System.Windows.Forms;

namespace ExportPartData
{
    public class ExportXmlSql
    {
        private static int CurrentVersion { get; set; }
        

        public static bool ExistXml(string modelName, int documentVersion, out string message)
        {
            message = "";

            #region To Delete

            //var name = new FileInfo(modelName).Name;
            //message = message +"\nName"+ name;

            //if (name != null)
            //{
            //    modelName = Path.GetFileNameWithoutExtension(name).ToUpper();
            //}

            #endregion

            bool exist = false;

            try
            {
                var xmlPartPath = new FileInfo(XmlPath + modelName + ".xml");
                if (!xmlPartPath.Exists) return false;

                var xmlPartVersion = GetXmlVersion(xmlPartPath.FullName);
                exist =  Equals(xmlPartVersion, documentVersion);
            }
            catch (Exception e)
            {
                message = $"Exception 1: {e.Message}";
                exist = false;
            }
            if (exist)
            {
                return true;
            }
            else if (!exist)
                try
                {
                    if (File.Exists(XmlPath + modelName + ".xml"))
                    {
                        var xmlPartVer = GetXmlVersion(XmlPath + modelName + ".xml");
                        exist = Equals(xmlPartVer, documentVersion);
                    }
                }
                catch (Exception e)
                {
                    message = $"Exception 2: {e.Message}";
                }

            if (exist)
            {
                return true;
            }
            else
            {
                return false;
            }


        }

        static int? GetXmlVersion(string xmlPath)
        {
            if (!xmlPath.EndsWith("xml")) return null;

            int? version = null;

            try
            {
                var coordinates = XDocument.Load(xmlPath);

                var enumerable = coordinates.Descendants("attribute")
                    .Select(
                        element =>
                            new
                            {
                                Number = element.FirstAttribute.Value,
                                Values = element.Attribute("value")
                            });
                foreach (var obj in enumerable)
                {
                    if (obj.Number != "Версия") continue;
                    version = Convert.ToInt32(obj.Values.Value);
                    goto m1;
                }
            }
            catch (Exception)
            {
                return 0;
            }

            m1:

            return version;
        }

        public static string XmlPath { get; set; } = @"\\pdmsrv\XML\"; //  @"\\pdmsrv\SolidWorks Admin\XML\"; // @"C:\Temp\"; //

        public static string ConnectionString { get; } = "Data Source=pdmsrv;Initial Catalog=SWPlusDB;Persist Security Info=True;User ID=sa;Password=P@$$w0rd;MultipleActiveResultSets=True";
        
        public static void Export(SldWorks swApp, int lastVer, int idPdm, out Exception exception)
        {
            Export(swApp, lastVer, idPdm, false, out exception);
        }

        public static void GetCurrentConfigPartData(SldWorks swApp, int lastVer, int idPdm, bool closeDoc, bool fixBends, out List<DataToExport> dataList,  out Exception exception)
        //public static void GetCurrentConfigPartData(SldWorks swApp, bool closeDoc, bool fixBends, out List<DataToExport> dataList, out Exception exception)
        {
            // Проход по всем родительским конфигурациям

            exception = null;
            dataList = new List<DataToExport>();

            var swModel = swApp.IActiveDoc2;

            if (swModel == null) return;

            var configName = ((Configuration) swModel.GetActiveConfiguration()).Name;

            swModel.ShowConfiguration2(configName);
            swModel.EditRebuild3();
            var swModelDocExt = swModel.Extension;

            var fileName = swModel.GetTitle().ToUpper().Replace(".SLDPRT", "");

            AddDimentions(swModel, configName, out exception);

            var confiData = new DataToExport
            {
                Config = configName,
                FileName = fileName,
                IdPdm = idPdm,
                Version = lastVer
            };

            #region Разгибание всех сгибов
            fixBends = true;
            if (fixBends)
            {
                swModel.EditRebuild3();
                List<Bends.SolidWorksFixPattern.PartBendInfo> list;
                Bends.Fix(swApp, out list, false); 
            }
            
          
            #endregion

            swModel.ForceRebuild3(false);

            var swCustProp = swModelDocExt.CustomPropertyManager[configName];
            string valOut;
            string materialId;

            // TO DO LOOK

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

            //UpdateCustomPropertyListFromCutList
            const string длинаГраничнойРамкиName = @"Длина граничной рамки";
            const string длинаГраничнойРамкиName2 = @"Bounding Box Length";
            const string ширинаГраничнойРамкиName = @"Ширина граничной рамки";
            const string ширинаГраничнойРамкиName2 = @"Bounding Box Width";
            const string толщинаЛистовогоМеталлаNAme = @"Толщина листового металла";
            const string толщинаЛистовогоМеталлаNAme2 = @"Sheet Metal Thickness";
            const string сгибыName = @"Сгибы";
            const string сгибыName2 = @"Bends";
            const string площадьПокрытияName = @"Площадь покрытия";//const string площадьПокрытияName2 = @"Bounding Box Area";

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

                            if (bodyFolder.GetCutListType() != (int)swCutListType_e.swSheetmetalCutlist)
                            {
                                goto m1;
                            }

                            swSubFeat.Select2(false, -1);
                            bodyFolder.SetAutomaticCutList(true);
                            bodyFolder.UpdateCutList();
                            var swCustPrpMgr = swSubFeat.CustomPropertyManager;
                            swCustPrpMgr.Add("Площадь поверхности", "Текст",
                                "\"SW-SurfaceArea@@@Элемент списка вырезов1@" +
                                Path.GetFileName(swModel.GetPathName()) + "\"");

                            string длинаГраничнойРамки;
                            swCustPrpMgr.Get4(длинаГраничнойРамкиName, true, out valOut,
                                out длинаГраничнойРамки);
                            if (string.IsNullOrEmpty(длинаГраничнойРамки))
                            {
                                swCustPrpMgr.Get4(длинаГраничнойРамкиName2, true, out valOut,
                                    out длинаГраничнойРамки);
                            }
                            swCustProp.Set(длинаГраничнойРамкиName, длинаГраничнойРамки);
                            confiData.ДлинаГраничнойРамки = длинаГраничнойРамки;

                            string ширинаГраничнойРамки;
                            swCustPrpMgr.Get4(ширинаГраничнойРамкиName, true, out valOut,
                                out ширинаГраничнойРамки);
                            if (string.IsNullOrEmpty(ширинаГраничнойРамки))
                            {
                                swCustPrpMgr.Get4(ширинаГраничнойРамкиName2, true, out valOut,
                                    out ширинаГраничнойРамки);
                            }
                            swCustProp.Set(ширинаГраничнойРамкиName, ширинаГраничнойРамки);
                            confiData.ШиринаГраничнойРамки = ширинаГраничнойРамки;

                            string толщинаЛистовогоМеталла;
                            swCustPrpMgr.Get4(толщинаЛистовогоМеталлаNAme, true, out valOut,
                                out толщинаЛистовогоМеталла);
                            if (string.IsNullOrEmpty(толщинаЛистовогоМеталла))
                            {
                                swCustPrpMgr.Get4(толщинаЛистовогоМеталлаNAme2, true, out valOut,
                                out толщинаЛистовогоМеталла);
                            }
                            swCustProp.Set(толщинаЛистовогоМеталлаNAme, толщинаЛистовогоМеталла);
                            confiData.ТолщинаЛистовогоМеталла = толщинаЛистовогоМеталла;

                            string сгибы;
                            swCustPrpMgr.Get4(сгибыName, true, out valOut, out сгибы);
                            if (string.IsNullOrEmpty(сгибы))
                            {
                                swCustPrpMgr.Get4(сгибыName2, true, out valOut, out сгибы);
                            }
                            swCustProp.Set(сгибыName, сгибы);
                            confiData.Сгибы = сгибы;

                            var myMassProp = swModel.Extension.CreateMassProperty();
                            var площадьПоверхности =
                                Convert.ToString(Math.Round(myMassProp.SurfaceArea * 1000) / 1000);

                            swCustProp.Set(площадьПокрытияName, площадьПоверхности);
                            try
                            {
                                confiData.ПлощадьПокрытия =
                                    double.Parse(площадьПоверхности.Replace(".", ","));
                            }
                            catch (Exception e)
                            {
                                exception = e;
                            }
                        }
                        m1:
                        swSubFeat = swSubFeat.GetNextFeature();
                    }
                }
                swFeat2 = swFeat2.GetNextFeature();
            }
            dataList.Add(confiData);

            if (!closeDoc)
            {
                return;
            }
            var namePrt = swApp.IActiveDoc2.GetTitle().ToLower().Contains(".sldprt")
                ? swApp.IActiveDoc2.GetTitle()
                : swApp.IActiveDoc2.GetTitle() + ".sldprt";
            swApp.CloseDoc(namePrt);
       
        }

        public static void Export(SldWorks swApp, int verToExport, int idPdm, bool closeDoc, out Exception exception)
        {
            exception = null;
            CurrentVersion = verToExport;
            //verToExport;

            #region Сбор информации по детали и сохранение разверток

            if (swApp == null)
            {
                try
                {
                    swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
                }
                catch (Exception)
                {
                    swApp = new SldWorks { Visible = false };
                }
            }

            IModelDoc2 swModel;

            try
            {
                swModel = swApp.IActiveDoc2;
            }
            catch (Exception)
            {
                swModel = swApp.IActiveDoc2;
            }

            if (swModel == null) return;

            var modelName = swModel.GetTitle();

            try
            {
                

                IPartDoc partDoc;
                try
                {
                    partDoc = (IPartDoc)((ModelDoc2)swModel);
                }
                catch (Exception)
                {
                    return;
                }

                bool sheetMetal = false;

                try
                {
                    sheetMetal = Part.IsSheetMetal(partDoc);
                }
                catch (Exception)
                {

                }

                if (!sheetMetal)
                {
                    //swApp.CloseDoc(Path.GetFileName(swModel.GetPathName()));

                    //13.10.2016

                    swApp.CloseAllDocuments(true);

                    //swApp.ExitApp();
                    return;
                }

                var activeconfiguration = (Configuration)swModel.GetActiveConfiguration();
                var swModelConfNames = (string[])swModel.GetConfigurationNames();

                foreach (var name in from name in swModelConfNames
                                     let config = (Configuration)swModel.GetConfigurationByName(name)
                                     where config.IsDerived()
                                     select name)
                {
                    swModel.DeleteConfiguration(name);
                }

                var swModelDocExt = swModel.Extension;
                var swModelConfNames2 = (string[])swModel.GetConfigurationNames();

                // Проход по всем родительским конфигурациям

                var dataList = new List<DataToExport>();

                var filePath = swModel.GetPathName();

                foreach (var configName in from name in swModelConfNames2
                                           let config = (Configuration)swModel.GetConfigurationByName(name)
                                           where !config.IsDerived()
                                           select name)
                {
                    // swModel.ShowConfiguration2(configName);
                    swModel.EditRebuild3();

                    AddDimentions(swModel, configName, out exception);

                    var confiData = new DataToExport
                    {
                        Config = configName,
                        FileName = filePath.Substring(filePath.LastIndexOf('\\') + 1),
                        IdPdm = idPdm
                    };

                    #region Разгибание всех сгибов

                    try
                    {
                        swModel.EditRebuild3();
                        List<Bends.SolidWorksFixPattern.PartBendInfo> list;
                        Bends.Fix(swApp, out list, false);
                    }
                    catch (Exception)
                    {
                        //
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

                    //UpdateCustomPropertyListFromCutList
                    const string длинаГраничнойРамкиName = @"Длина граничной рамки";
                    const string длинаГраничнойРамкиName2 = @"Bounding Box Length";
                    const string ширинаГраничнойРамкиName = @"Ширина граничной рамки";
                    const string ширинаГраничнойРамкиName2 = @"Bounding Box Width";
                    const string толщинаЛистовогоМеталлаNAme = @"Толщина листового металла";
                    const string толщинаЛистовогоМеталлаNAme2 = @"Sheet Metal Thickness";//Sheet Metal Thickness
                    const string сгибыName = @"Сгибы";
                    const string сгибыName2 = @"Bends";
                    const string площадьПокрытияName = @"Площадь покрытия";
                    const string площадьПокрытияName2 = @"Bounding Box Area";

                    Feature swFeat2 = swModel.FirstFeature();
                    while (swFeat2 != null)
                    {
                        if (swFeat2.GetTypeName2() == "SolidBodyFolder")
                        {
                            //  List<Bends.SolidWorksFixPattern.PartBendInfo> list;
                            //  Bends.Fix(swApp, out list, false);
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

                                    if (bodyFolder.GetCutListType() != (int)swCutListType_e.swSheetmetalCutlist)
                                    {
                                        goto m1;
                                    }

                                    swSubFeat.Select2(false, -1);
                                    bodyFolder.SetAutomaticCutList(true);
                                    bodyFolder.UpdateCutList();
                                    var swCustPrpMgr = swSubFeat.CustomPropertyManager;
                                    swCustPrpMgr.Add("Площадь поверхности", "Текст",
                                        "\"SW-SurfaceArea@@@Элемент списка вырезов1@" +
                                        Path.GetFileName(swModel.GetPathName()) + "\"");

                                    string длинаГраничнойРамки;
                                    swCustPrpMgr.Get4(длинаГраничнойРамкиName, true, out valOut,
                                        out длинаГраничнойРамки);
                                    if (string.IsNullOrEmpty(длинаГраничнойРамки))
                                    {
                                        swCustPrpMgr.Get4(длинаГраничнойРамкиName2, true, out valOut,
                                            out длинаГраничнойРамки);
                                    }
                                    swCustProp.Set(длинаГраничнойРамкиName, длинаГраничнойРамки);
                                    confiData.ДлинаГраничнойРамки = длинаГраничнойРамки;

                                    string ширинаГраничнойРамки;
                                    swCustPrpMgr.Get4(ширинаГраничнойРамкиName, true, out valOut,
                                        out ширинаГраничнойРамки);
                                    if (string.IsNullOrEmpty(ширинаГраничнойРамки))
                                    {
                                        swCustPrpMgr.Get4(ширинаГраничнойРамкиName2, true, out valOut,
                                            out ширинаГраничнойРамки);
                                    }
                                    swCustProp.Set(ширинаГраничнойРамкиName, ширинаГраничнойРамки);
                                    confiData.ШиринаГраничнойРамки = ширинаГраничнойРамки;

                                    string толщинаЛистовогоМеталла;
                                    swCustPrpMgr.Get4(толщинаЛистовогоМеталлаNAme, true, out valOut,
                                        out толщинаЛистовогоМеталла);
                                    if (string.IsNullOrEmpty(толщинаЛистовогоМеталла))
                                    {
                                        swCustPrpMgr.Get4(толщинаЛистовогоМеталлаNAme2, true, out valOut,
                                        out толщинаЛистовогоМеталла);
                                    }
                                    swCustProp.Set(толщинаЛистовогоМеталлаNAme, толщинаЛистовогоМеталла);
                                    confiData.ТолщинаЛистовогоМеталла = толщинаЛистовогоМеталла;

                                    string сгибы;
                                    swCustPrpMgr.Get4(сгибыName, true, out valOut, out сгибы);
                                    if (string.IsNullOrEmpty(сгибы))
                                    {
                                        swCustPrpMgr.Get4(сгибыName2, true, out valOut, out сгибы);
                                    }
                                    swCustProp.Set(сгибыName, сгибы);
                                    confiData.Сгибы = сгибы;

                                    var myMassProp = swModel.Extension.CreateMassProperty();
                                    var площадьПоверхности =
                                        Convert.ToString(Math.Round(myMassProp.SurfaceArea * 1000) / 1000);

                                    swCustProp.Set(площадьПокрытияName, площадьПоверхности);
                                    try
                                    {
                                        confiData.ПлощадьПокрытия =
                                            double.Parse(площадьПоверхности.Replace(".", ","));
                                    }
                                    catch (Exception e)
                                    {
                                        exception = e;
                                    }
                                }
                                m1:
                                swSubFeat = swSubFeat.GetNextFeature();
                            }
                        }
                        swFeat2 = swFeat2.GetNextFeature();
                    }
                    dataList.Add(confiData);
                }

                swModel.ShowConfiguration2(activeconfiguration.Name);

                ExportDataToXmlSql(swModel.GetTitle().ToUpper().Replace(".SLDPRT", ""), dataList, out exception);

                #endregion

                if (!closeDoc) return;
                var namePrt = swApp.IActiveDoc2.GetTitle().ToLower().Contains(".sldprt")
                    ? swApp.IActiveDoc2.GetTitle()
                    : swApp.IActiveDoc2.GetTitle() + ".sldprt";
                swApp.CloseDoc(namePrt);
            }

            catch (Exception e)
            {
                exception = e;
            }

            finally
            {
                swApp.CloseDoc(Path.GetFileName(swModel.GetPathName()));
            }
        }

        public class DataToExport
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
            public int Version;
        }

        public static void ExportDataToXmlSql(string fileName, IEnumerable<DataToExport> dataToExport, out Exception exception)
        {
            exception = null;
            if (fileName == null || dataToExport == null) return;
            try
            {
                var myXml = new XmlTextWriter(XmlPath + fileName + ".xml", Encoding.UTF8);

                myXml.WriteStartDocument();
                myXml.Formatting = Formatting.Indented;
                myXml.Indentation = 2;

                // создаем элементы
                myXml.WriteStartElement("xml");
                myXml.WriteStartElement("transactions");
                myXml.WriteStartElement("transaction");

                myXml.WriteStartElement("document");

                foreach (var configData in dataToExport)
                {
                    #region XML

                    try
                    { 
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
                        myXml.WriteAttributeString("value", Convert.ToString(configData.ПлощадьПокрытия).Replace(",", "."));
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

                        // PaintX
                        myXml.WriteStartElement("attribute");
                        myXml.WriteAttributeString("name", "PaintX");
                        myXml.WriteAttributeString("value", Convert.ToString(configData.PaintX).Replace(",", "."));
                        myXml.WriteEndElement();

                        // PaintY
                        myXml.WriteStartElement("attribute");
                        myXml.WriteAttributeString("name", "PaintY");
                        myXml.WriteAttributeString("value", Convert.ToString(configData.PaintY).Replace(",", "."));
                        myXml.WriteEndElement();

                        // PaintZ
                        myXml.WriteStartElement("attribute");
                        myXml.WriteAttributeString("name", "PaintZ");
                        myXml.WriteAttributeString("value", Convert.ToString(configData.PaintZ).Replace(",", "."));
                        myXml.WriteEndElement();


                        // Версия последняя
                        myXml.WriteStartElement("attribute");
                        myXml.WriteAttributeString("name", "Версия");
                        myXml.WriteAttributeString("value", configData.Version != 0 ? configData.Version.ToString() : Convert.ToString(CurrentVersion));
                        myXml.WriteEndElement();
                    
                        myXml.WriteEndElement();  //configuration
                    }
                    catch (Exception ex)
                    {
                        //MessageBox.Show($"{ex.ToString()}\n + {ex.StackTrace}\n (Name - {configData.FileName} ID - {configData.IdPdm} Conf - {configData.Config} Ver - {configData.Version})");
                        exception = ex;
                    
                    }

                    #endregion

                    #region SQL                   

                    try
                    {
                        using (var sqlConnection = new SqlConnection(ConnectionString))
                        {
                            sqlConnection.Open();

                            var spcmd = new SqlCommand("UpDateCutList", sqlConnection)
                            {
                                CommandType = CommandType.StoredProcedure
                            };

                            double workpieceX;
                            double.TryParse(configData.ДлинаГраничнойРамки.Replace('.', ','), out workpieceX);
                            double workpieceY;
                            double.TryParse(configData.ШиринаГраничнойРамки.Replace('.', ','), out workpieceY);
                            int bend;
                            int.TryParse(configData.Сгибы, out bend);
                            double thickness;
                            double.TryParse(configData.ТолщинаЛистовогоМеталла.Replace('.', ','), out thickness);

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


                            spcmd.Parameters.AddWithValue("@SurfaceArea", ParseDouble(configData.ПлощадьПокрытия.ToString()));


                            spcmd.Parameters.Add("@WorkpieceX", SqlDbType.Float).Value = workpieceX;
                            spcmd.Parameters.Add("@WorkpieceY", SqlDbType.Float).Value = workpieceY;

                            spcmd.Parameters.Add("@Bend", SqlDbType.Int).Value = bend;
                            spcmd.Parameters.Add("@Thickness", SqlDbType.Float).Value = thickness;
                            spcmd.Parameters.Add("@Version", SqlDbType.Int).Value = configData.Version != 0 ? configData.Version : CurrentVersion;
                            spcmd.Parameters.AddWithValue("@configuration", configuration);

                            #region
                            //spcmd.Parameters.Add("@configuration", SqlDbType.NVarChar).Value = configuration;
                            //query = $"MaterialID- {materialId}\nPaintX- {configData.PaintX}\nFILENAME- {configData.FileName}\nIDPDM- {configData.IdPdm}\nSurfaceArea- {ParseDouble(configData.ПлощадьПокрытия.ToString())}\nWorkpieceX- {workpieceX}\nConfiguration- {configuration}";

                            //MessageBox.Show($"MaterialID- {materialId}\nPaintX- {configData.PaintX}\nFILENAME- {configData.FileName}\nIDPDM- {configData.IdPdm}\nSurfaceArea- {ParseDouble(configData.ПлощадьПокрытия.ToString())}\nWorkpieceX- {workpieceX}\nConfiguration- {configuration}");
                            #endregion

                            spcmd.ExecuteNonQuery();

                            sqlConnection.Close();
                        }
                    }
                    catch (Exception ex)
                    {
                        //MessageBox.Show($"{ex.ToString()}\n + {ex.StackTrace}\n (Name - {configData.FileName} ID - {configData.IdPdm} Conf - {configData.Config} Ver - {configData.Version})");
                        exception = ex;
                        // MessageBox.Show(query);

                        Logger.Add("============ Исключение при добавлении свойств в базу данных  ");
                        Logger.Add(ex.ToString());

                    }

                    #endregion
                }

                myXml.WriteEndElement();// ' элемент DOCUMENT
                myXml.WriteEndElement();// ' элемент TRANSACTION
                myXml.WriteEndElement();// ' элемент TRANSACTIONS
                myXml.WriteEndElement();// ' элемент XML
                // заносим данные в myMemoryStream
                myXml.Flush();

                myXml.Close();
            }
            catch (Exception ex)
            {
                exception = ex;
            }
        }

        private static double ParseDouble(string value)
        {
            value = value.Replace(" ", "");
            value = CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator == "," ? value.Replace(".", ",") : value.Replace(",", ".");
            var splited = value.Split(CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator[0]);
            if (splited.Length <= 2) return double.Parse(value);
            var r = "";
            for (var i = 0; i < splited.Length; i++)
            {
                if (i == splited.Length - 1)
                    r += CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;
                r += splited[i];
            }
            value = r;
            return double.Parse(value);
        }

        private static void AddDimentions(IModelDoc2 swmodel, string configname, out Exception exception)
        {
            const long valueset = 1000;
            exception = null;

            try
            {
                swmodel.GetConfigurationByName(configname);
                swmodel.EditRebuild3();

                var part = (PartDoc) swmodel;
                var box = part.GetPartBox(true);

                swmodel.AddCustomInfo3(configname, "Длина", 30, "");
                swmodel.AddCustomInfo3(configname, "Ширина", 30, "");
                swmodel.AddCustomInfo3(configname, "Высота", 30, "");

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

                swmodel.EditRebuild3();
            
            }
            catch (Exception ex)
            {
                exception = ex;
            }
        }

    }
    
    public class Part
    {
        class SheetMetal
        {
            const string LicenseKeyDM = "Vents:swdocmgr_general-11785-02051-00064-01025-08442-34307-00007-29544-01291-09600-30446-40775-52437-15814-04102-21610-48740-21827-48271-33444-51679-23957-38232-53689-00461-25656-23152-49876-23260-24676-27746-1,swdocmgr_previews-11785-02051-00064-01025-08442-34307-00007-02240-63424-57067-45262-39174-54588-15723-55300-13841-06627-15439-20060-24303-18368-22568-38232-53689-00461-25656-23152-49876-23260-24676-27746-0";
            string filePath;
            public SheetMetal(string filePath)
            {
                this.filePath = filePath;
            }
            public bool ExistSheetMetalDM()
            {
                var result = false;
                try
                {
                    var sLicenseKey = LicenseKeyDM;
                    var nDocType = SwDmDocumentType.swDmDocumentPart;
                    var swClassFact = new SwDMClassFactory();
                    var swDocMgr = swClassFact.GetApplication(sLicenseKey);
                    var nRetVal = default(SwDmDocumentOpenError);
                    var swDocument10 = (SwDMDocument10)swDocMgr.GetDocument(filePath, nDocType, true, out nRetVal); // true - если файл только для чтения
                    if (swDocument10 != null)
                    {
                        var swDocument13 = (SwDMDocument13)swDocument10;
                        if (swDocument13 != null)
                        {
                            var CutListItems = (object[])swDocument13.GetCutListItems2();
                            if (CutListItems != null)
                            { result = true; }
                        }
                    }

                }
                catch (Exception ex)
                {
                    //Logger.ToLog($"ERROR SHEET METAL: {ex.ToString()}, {ex.StackTrace}", 10001);
                }
                return result;
            }
            ~SheetMetal() { }
        }

        //public static bool IsSheetMetal(string swPartPath)
        //{
        //    return new SheetMetal(swPartPath).ExistSheetMetalDM();
        //}

        public static bool IsSheetMetal(IPartDoc swPart)
        {
            return Dxf.IsSheetMetalPart(swPart);
        }
    }


    public class Dxf
    {
        public static bool ExistDxf(int idPdm, int currentVersion, string configuration, out Exception exception)
        {
            exception = null;
            var exist = false;
            string res = "";
            try
            {
                using (var sqlConnection = new SqlConnection(ExportXmlSql.ConnectionString))
                using (var command = new SqlCommand("DXFCheck", sqlConnection)
                {
                    CommandType = CommandType.StoredProcedure
                })
                {
                    command.Parameters.AddWithValue("IDPDM", idPdm);
                    command.Parameters.AddWithValue("Configuration", configuration);
                    command.Parameters.AddWithValue("Version", currentVersion);

                    var returnParameter = command.Parameters.Add("@ReturnVal", SqlDbType.Int);
                    returnParameter.Direction = ParameterDirection.ReturnValue;

                    sqlConnection.Open();
                    command.ExecuteNonQuery();

                    var result = returnParameter.Value;
                    res = result.ToString();
                    exist = result.ToString() == "1";
                }
            }
            catch (Exception ex)
            {
                exception = ex;
                // MessageBox.Show(ex.ToString() + "\n" + ex.StackTrace);
                exist = false;
            }

           // MessageBox.Show(exist.ToString(), res);

            return exist;

        }

        public static void SaveByteToFile(byte[] blob, string varPathToNewLocation)
        {
            Database.SaveFile(blob, varPathToNewLocation);
        }
       
        public static bool Save(string partPath,  string configuration, out Exception exception,  bool fixBends, bool closeAfterSave, string folderToSave = null, bool includeNonSheetParts = false)
        {
            List<DxfFile> dxfList;
           return Save(partPath, folderToSave, configuration, null, out exception, 0, 0, out dxfList, fixBends, closeAfterSave, includeNonSheetParts);
        }

        public static bool Save(SldWorks swApp, out Exception exception, int idPdm, int version, out List<DxfFile> dxfList, bool fixBends, bool closeAfterSave, string configuration = null, bool includeNonSheetParts = false)
        {
           return Save(null, null, configuration, swApp, out exception,  idPdm, version, out dxfList, fixBends, closeAfterSave, includeNonSheetParts);
        }

        public static bool Save(string partPath, string folderToSave, string configuration, out Exception exception, int idPdm, int version, out List<DxfFile> dxfList, bool fixBends, bool closeAfterSave, bool includeNonSheetParts)
        {
           return Save(partPath, folderToSave, configuration, null, out exception, idPdm, version, out dxfList, fixBends, closeAfterSave, includeNonSheetParts);
        }

        public static bool Save(string folderToSave, string configuration, SldWorks swApp, out Exception exception, int idPdm, int version, out List<DxfFile> dxfList, bool fixBends, bool closeAfterSave, bool includeNonSheetParts = false)
        {
           return Save(null, folderToSave, configuration, swApp, out exception, idPdm, version, out dxfList, fixBends, closeAfterSave, includeNonSheetParts);
        }

        public class DxfFile
        {
            public int IdPdm { get; set; }
            public int Version { get; set; }
            public string Configuration { get; set; }
            public string FilePath { get; set; }
        }
          
        static void GetDxf(int idPdm, int version, string configuration)
        {

        }

        public static string TempDxfFolder { get; set; } =  @"C:\DXF\";

        static bool Save(string partPath, string folderToSave, string configuration, SldWorks swApp, out Exception exception, int idPdm, int version, out List<DxfFile> dxfList, bool fixBends, bool closeAfterSave, bool includeNonSheetParts)
        {
            var save = false;

            exception = null;
            dxfList = new List<DxfFile>();

            if (string.IsNullOrEmpty(folderToSave)) folderToSave = TempDxfFolder;

            try
            {
                if (swApp == null)
                {
                    try
                    {
                        swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
                    }
                    catch (Exception)
                    {
                      
                        swApp = new SldWorks { Visible = true };
                    }
                }

                var thisProc = Process.GetProcessesByName("SLDWORKS")[0];
                thisProc.PriorityClass = ProcessPriorityClass.RealTime;

                IModelDoc2 swModel = null;

                if (!string.IsNullOrEmpty(partPath))
                {                    
                    try
                    {
                        swModel = swApp.OpenDoc6(partPath, (int)swDocumentTypes_e.swDocPART,
                            (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", 0, 0);
                    }
                    catch (Exception ex)
                    {
                        exception = ex;
                    }
                }
                else
                {
                    swModel = swApp.IActiveDoc2;
                }

                if (swModel == null) return save;

                bool sheetmetal = true;

                if (!IsSheetMetalPart((IPartDoc)swModel))
                {
                    sheetmetal = false;
                    if (!includeNonSheetParts)
                    {
                        CloseModel(swModel, swApp);
                        return save;
                    }
                }
              

                if (!string.IsNullOrEmpty(configuration))
                {
                    try
                    {
                        string filePath;      
                        save = SaveThisConfigDxf(folderToSave, configuration, fixBends ? swApp : null, swModel, out filePath, sheetmetal);
                        dxfList.Add(new DxfFile
                        {
                            Configuration = configuration,
                            FilePath = filePath,
                            IdPdm = idPdm,
                            Version = version
                        });
                    }
                    catch (Exception ex)
                    {
                        exception = ex;
                    }
                    if (closeAfterSave) { CloseModel(swModel, swApp); }
                }
                else
                {

                    var swModelConfNames2 = (string[])swModel.GetConfigurationNames();

                    foreach (var configName in from name in swModelConfNames2
                                               let config = (Configuration)swModel.GetConfigurationByName(name)
                                               where !config.IsDerived()
                                               select name)
                    {
                        string filePath;
                      //  MessageBox.Show(swModel.GetTitle() + "\nsheetmetal-" + sheetmetal.ToString());
                        save = SaveThisConfigDxf(folderToSave, configName, fixBends ? swApp : null, swModel, out filePath, sheetmetal);
                        dxfList.Add(new DxfFile
                        {
                            Configuration = configName,
                            FilePath = filePath,
                            IdPdm = idPdm,
                            Version = version
                        });
                    }
                    if (closeAfterSave) { CloseModel(swModel, swApp); }
                }
            }
            catch (Exception ex)
            {
                exception = ex;
            }

            return save;
            
        }

       public static  bool SaveThisConfigDxf(string folderToSave, string configuration, SldWorks swApp, IModelDoc2 swModel, out string dxfFilePath, bool sheetmetal)
        {
          
            swModel.ShowConfiguration2(configuration);
            swModel.EditRebuild3();
            
            if (swApp != null && sheetmetal)
            {                
                List<Bends.SolidWorksFixPattern.PartBendInfo> list;
                Bends.Fix(swApp, out list, true);                            
            }

            var sDxfName = DxfName(swModel.GetTitle(), configuration) + ".dxf";

           dxfFilePath = Path.Combine(folderToSave, sDxfName);
          //  dxfFilePath = Path.Combine(@"C:\DXF", sDxfName);

            Directory.CreateDirectory(folderToSave);

            var dataAlignment = new double[12];

            dataAlignment[0] = 0.0;
            dataAlignment[1] = 0.0;
            dataAlignment[2] = 0.0;
            dataAlignment[3] = 1.0;
            dataAlignment[4] = 0.0;
            dataAlignment[5] = 0.0;
            dataAlignment[6] = 0.0;
            dataAlignment[7] = 1.0;
            dataAlignment[8] = 0.0;
            dataAlignment[9] = 0.0;
            dataAlignment[10] = 0.0;
            dataAlignment[11] = 1.0;
            object varAlignment = dataAlignment;

            var swPart = (IPartDoc)swModel;
            int sheetmetalOptions = SheetMetalOptions(true, sheetmetal, false, false, false, true, false);

            //MessageBox.Show(sheetmetalOptions.ToString());                               

            if (sheetmetal)
            {
                return swPart.ExportToDWG(dxfFilePath, swModel.GetPathName(), sheetmetal ? 1 : 2, true, varAlignment, false, false, sheetmetalOptions, sheetmetal ? 0 : 3);
            }
            else
            {                                  
                return swModel.SaveAs4(dxfFilePath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, 0, 0);                
            }
        }
        
        static int SheetMetalOptions(bool ExportGeometry, bool IcnludeHiddenEdges, bool ExportBendLines, bool IncludeScetches, bool MergeCoplanarFaces, bool ExportLibraryFeatures, bool ExportFirmingTools)
        {
            return SheetMetalOptions(
                ExportGeometry ? 1 : 0,
                IcnludeHiddenEdges ? 1 : 0,
                ExportBendLines ? 1 : 0,
                IncludeScetches ? 1 : 0,
                MergeCoplanarFaces ? 1 : 0,
                ExportLibraryFeatures ? 1 : 0,
                ExportFirmingTools ? 1 : 0);
        }
        static int SheetMetalOptions(int p0, int p1, int p2, int p3, int p4, int p5, int p6)
        {
            return p0 * 1 + p1 * 2 + p2 * 4 + p3 * 8 + p4 * 16 + p5 * 32 + p6 * 64;
        }


        static void CloseModel(IModelDoc2 swModel, ISldWorks swApp)
        {
            swApp.CloseDoc(swModel.GetTitle().ToLower().Contains(".sldprt") ? swModel.GetTitle() : swModel.GetTitle() + ".sldprt");
        }

        static string GetFromCutlist(IModelDoc2 swModel, string property)
        {
            string propertyValue = null;

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
                                string valOut;
                                swCustPrpMgr.Get4(property, true, out valOut, out propertyValue);
                            }
                            swSubFeat = swSubFeat.GetNextFeature();
                        }
                    }
                    swFeat2 = swFeat2.GetNextFeature();
                }
            }
            catch (Exception)
            {
                //
            }

            return propertyValue;
        }

        public static bool IsSheetMetalPart(IPartDoc swPart)
        {
            var isSheetMetal = false;            
            try
            {
                var vBodies = swPart.GetBodies2((int)swBodyType_e.swSolidBody, false);

                foreach (Body2 vBody in vBodies)
                {
                    isSheetMetal = vBody.IsSheetMetal();
                    if (isSheetMetal)
                    {
                        return true;
                    }
                }
            }
            catch (Exception)
            {
                return false;
            }
            return isSheetMetal;
        }

        public static string DxfName(string fileName, string config)
        {
            //return $"{fileName.ToLower().Replace(".sldprt", "").ToUpper()}-{config}";
            return $"{fileName.Replace("ВНС-", "").ToLower().Replace(".sldprt", "")}-{config}";
        }

        public class ResultList
        {
            public Exception exc { get; set; }

            public DxfFile dxfFile { get; set; }

        }

        public static void AddToSql(List<DxfFile> dxfList, bool deleteFiles, out List<ResultList> resultList)
        {
            resultList = new List<ResultList>();

            foreach (var file in dxfList)
            {
                Exception ex;
                Database.AddDxf(file.FilePath, file.IdPdm, file.Configuration, file.Version, out ex);
                if (ex != null)
                {
                    resultList.Add(new ResultList { exc = ex , dxfFile = file } );
                }
            }

            if (!deleteFiles) return;

            FileInfo fileInf;
            var files = new DirectoryInfo(TempDxfFolder).GetFiles();
            foreach (var fileInfo in files)
            {
                fileInf = fileInfo;
                try
                {                    
                    fileInfo.Delete();
                }
                catch (Exception ex)
                {
                    resultList.Add(new ResultList { exc = ex, dxfFile = new DxfFile { FilePath = fileInf.FullName } });
                }
            }
        }

    }

    public class Database
    {
        public static void AddDxf(string varFilePath, int idPdm, string configuration, int version, out Exception exception)
        {
            byte[] file;
            exception = null;
            using (var stream = new FileStream(varFilePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = new BinaryReader(stream))
                {
                    file = reader.ReadBytes((int) stream.Length);
                }
            }

            using (var con = new SqlConnection(ExportXmlSql.ConnectionString))
            {
                try
                {
                    con.Open();
                    var sqlCommand = new SqlCommand("ExportDXF", con) {CommandType = CommandType.StoredProcedure};
                    
                    sqlCommand.Parameters.Add("@DXF", SqlDbType.VarBinary, file.Length).Value = file;
                    sqlCommand.Parameters.Add("@IDPDM", SqlDbType.Int).Value = idPdm;
                    sqlCommand.Parameters.Add("@Configuration", SqlDbType.NVarChar).Value = configuration;
                    sqlCommand.Parameters.Add("@Version", SqlDbType.Int).Value = version;
                    
                    sqlCommand.ExecuteNonQuery();
                }
                catch (Exception e)
                {
                    exception = e;
                    //MessageBox.Show(e.Message + "\n" + e.StackTrace + "\n" + "Name - "+ varFilePath + $"\nIDPDM - {idPdm} config {configuration} ver - {version}");
                }
                finally
                {
                    con.Close();
                }
            }
        }

        public static void DeleteDxf(int idPdm, string configuration, int version, out Exception exc)
        {
            exc = null;
            using (var con = new SqlConnection(ExportXmlSql.ConnectionString))
            {
                try
                {
                    con.Open();
                    var sqlCommand = new SqlCommand("DXFDelete", con) { CommandType = CommandType.StoredProcedure };
                    
                    sqlCommand.Parameters.Add("@IDPDM", SqlDbType.Int).Value = idPdm;
                    sqlCommand.Parameters.Add("@Configuration", SqlDbType.NVarChar).Value = configuration;
                    sqlCommand.Parameters.Add("@Version", SqlDbType.Int).Value = version;

                    sqlCommand.ExecuteNonQuery();
                }
                catch (Exception e)
                {
                    exc = e;
                    //  MessageBox.Show(e.Message + "\n" + e.StackTrace + "\n" + $"IDPDM - {idPdm} config {configuration} ver - {version}");
                }
                finally
                {
                    con.Close();
                }
            }
        }

        public static void DatabaseFilePut(string varFilePath)
        {
            byte[] file;
            using (var stream = new FileStream(varFilePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = new BinaryReader(stream))
                {
                    file = reader.ReadBytes((int)stream.Length);
                }
            }
            using (var varConnection = new SqlConnection(ExportXmlSql.ConnectionString))
            using (var sqlWrite = new SqlCommand("INSERT INTO Parts (DXF, WorkpieceX, WorkpieceY, Thickness, ConfigurationID, Version, PaintX, PaintY, PaintZ, NomenclatureID) Values (@File, @WorkpieceX, @WorkpieceY, @Thickness, @ConfigurationID, @Version, @PaintX, @PaintY, @PaintZ, @NomenclatureID)", varConnection))
            {
                varConnection.Open();
                sqlWrite.Parameters.Add("@WorkpieceX", SqlDbType.Float).Value = 1000;
                sqlWrite.Parameters.Add("@WorkpieceY", SqlDbType.Float).Value = 1250;
                sqlWrite.Parameters.Add("@Thickness", SqlDbType.Float).Value = 2;
                sqlWrite.Parameters.Add("@ConfigurationID", SqlDbType.Float).Value = 2;
                sqlWrite.Parameters.Add("@Version", SqlDbType.Float).Value = 3;

                sqlWrite.Parameters.Add("@PaintX", SqlDbType.Float).Value = 3;
                sqlWrite.Parameters.Add("@PaintY", SqlDbType.Float).Value = 4;
                sqlWrite.Parameters.Add("@PaintZ", SqlDbType.Float).Value = 5;

                sqlWrite.Parameters.Add("@NomenclatureID", SqlDbType.Float).Value = 5;

                sqlWrite.Parameters.Add("@File", SqlDbType.VarBinary, file.Length).Value = file;
                sqlWrite.ExecuteNonQuery();
                varConnection.Close();
            }
        }

        public static SqlDataReader GetFile(string configuration, int idpdm, int version, out Exception exc)
        {
            exc = null;
            using (var varConnection = new SqlConnection(ExportXmlSql.ConnectionString))
            using (var sqlQuery = new SqlCommand(@"select p.DXF from Parts p
                                                    inner join Nomenclature n on p.NomenclatureID = n.NomenclatureID
                                                    inner join [Vents-PDM].dbo.DocumentConfiguration dc on p.ConfigurationID = dc.ConfigurationID
                                                    where n.IDPDM = @IDPDM AND p.Version = @Version AND dc.ConfigurationName = @ConfigurationName",
                                                varConnection))
            { 
                varConnection.Open();

                try
                {
                    sqlQuery.Parameters.AddWithValue("@IDPDM", idpdm);
                    sqlQuery.Parameters.AddWithValue("@Version", version);
                    sqlQuery.Parameters.AddWithValue("@ConfigurationName", configuration);

                    return sqlQuery.ExecuteReader();
                }
                catch (Exception e)
                {
                    exc = e;
                    return null;
                }
                finally
                {
                    varConnection.Close();
                }
            }
        }

        public static void SaveFile(byte[] blob, string varPathToNewLocation)
        {
            using (var fs = new FileStream(varPathToNewLocation, FileMode.Create, FileAccess.Write))
                fs.Write(blob, 0, blob.Length);
        }

        public static byte[] DatabaseFileRead(int idpdm, int version, string configuration, out string codeMaterial, out double? thikness, out Exception exc)
        {
            exc = null;
            codeMaterial = null;
            thikness = null;
            using (var varConnection = new SqlConnection(ExportXmlSql.ConnectionString)) 
            using (var sqlQuery = new SqlCommand(
                                                    @"select p.DXF, m.CodeMaterial, p.Thickness
                                                    from Parts p
                                                    inner join Nomenclature n on p.NomenclatureID = n.NomenclatureID
                                                    Left JOIN MaterialsProp m on p.MaterialID = m.LevelID
                                                    inner join [Vents-PDM].dbo.DocumentConfiguration dc on p.ConfigurationID = dc.ConfigurationID                                                    
                                                    where n.IDPDM = @IDPDM AND p.Version = @Version AND dc.ConfigurationName = @ConfigurationName ",                                                
                                                 varConnection))
            {
                varConnection.Open();

                sqlQuery.Parameters.AddWithValue("@IDPDM", idpdm);
                sqlQuery.Parameters.AddWithValue("@Version", version);
                sqlQuery.Parameters.AddWithValue("@ConfigurationName", configuration);

                try
                {
                    using (var sqlQueryResult = sqlQuery.ExecuteReader())
                    {
                        try
                        {
                            sqlQueryResult.Read();
                            //MessageBox.Show("RecordsAffected" + sqlQueryResult.RecordsAffected);
                        }
                        catch (Exception ex)
                        {
                          //  MessageBox.Show(ex.ToString() + "\n" + ex.StackTrace + "\n" + idpdm + "\n" + version + "\n" + configuration);
                        }

                        byte[] blob = null;

                        try
                        {
                            blob = new byte[(sqlQueryResult.GetBytes(0, 0, null, 0, int.MaxValue))];
                            sqlQueryResult.GetBytes(0, 0, blob, 0, blob.Length);
                        }
                        catch (Exception ex)
                        {
                     //       MessageBox.Show(ex.ToString() + "\n" + idpdm + "\n" + version + "\n" + configuration, "Нажмите ОК для продолжения");
                          //  MessageBox.Show(ex.ToString());
                            return null;
                        }                        

                        try
                        {
                            codeMaterial = sqlQueryResult.GetString(1);
                        }
                        catch (Exception ex)
                        {
                           // MessageBox.Show(ex.ToString());
                        }
                            try
                        {
                            //Convert.ToDecimal()
                            thikness = Convert.ToDouble(sqlQueryResult["Thickness"]);
                            //thikness = (double?)sqlQueryResult.GetValue(2);
                        }
                        catch (Exception ex)
                        {
                           // MessageBox.Show(ex.ToString());
                        }
                        return blob;
                    }
                }
                catch (Exception e)
                {
                    exc = e;
                    return null;
                }
                finally
                {
                    varConnection.Close();
                }
            }
        }

        public static void DatabaseFileRead(string varId, string varPathToNewLocation)
        {
            using (var varConnection = new SqlConnection(ExportXmlSql.ConnectionString))
            using (var sqlQuery = new SqlCommand(@"SELECT [DXF] FROM [Parts] WHERE [PartID] = @varId", varConnection))

            {
                varConnection.Open();
                sqlQuery.Parameters.AddWithValue("@varId", varId);
                using (var sqlQueryResult = sqlQuery.ExecuteReader())
            {
                sqlQueryResult.Read();
                var blob = new byte[(sqlQueryResult.GetBytes(0, 0, null, 0, int.MaxValue))];
                sqlQueryResult.GetBytes(0, 0, blob, 0, blob.Length);
                using (var fs = new FileStream(varPathToNewLocation, FileMode.Create, FileAccess.Write))
                    fs.Write(blob, 0, blob.Length);
            }
            varConnection.Close();
        }


    }

        public static MemoryStream DatabaseFileRead(string varId)
        {
            var memoryStream = new MemoryStream();
            using (var varConnection = new SqlConnection(ExportXmlSql.ConnectionString))
            using (var sqlQuery = new SqlCommand(@"SELECT [RaportPlik] FROM [dbo].[Raporty] WHERE [RaportID] = @varID", varConnection))
            {
                sqlQuery.Parameters.AddWithValue("@varID", varId);
                using (var sqlQueryResult = sqlQuery.ExecuteReader())
                {
                    sqlQueryResult.Read();
                    var blob = new byte[(sqlQueryResult.GetBytes(0, 0, null, 0, int.MaxValue))];
                    sqlQueryResult.GetBytes(0, 0, blob, 0, blob.Length);
                    //using (var fs = new MemoryStream(memoryStream, FileMode.Create, FileAccess.Write)) {
                    memoryStream.Write(blob, 0, blob.Length);
                    //}
                }
            }
            return memoryStream;
        }

        public static int DatabaseFilePut(MemoryStream fileToPut)
        {
            var varId = 0;
            var file = fileToPut.ToArray();
            const string preparedCommand = @"
                    INSERT INTO [dbo].[Raporty]
                               ([RaportPlik])
                         VALUES
                               (@File)
                        SELECT [RaportID] FROM [dbo].[Raporty]
            WHERE [RaportID] = SCOPE_IDENTITY()
                    ";
            using (var varConnection = new SqlConnection(ExportXmlSql.ConnectionString))
            using (var sqlWrite = new SqlCommand(preparedCommand, varConnection))
            {
                sqlWrite.Parameters.Add("@File", SqlDbType.VarBinary, file.Length).Value = file;

                using (var sqlWriteQuery = sqlWrite.ExecuteReader())
                    while (sqlWriteQuery.Read())
                    {
                        varId = sqlWriteQuery["RaportID"] as int? ?? 0;
                    }
            }
            return varId;
        }

        void Temp()
        {
            byte[] data = null;
            var xml = Encoding.UTF8.GetString(data);
            var xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(xml);
            // TODO: do something with the resulting XmlDocument


            DataTable dataTable = null;
            using (var stream = new MemoryStream(data))
            {
                dataTable.ReadXml(stream);
            }
        }

    }

    public class Bends
    {
        public static void Fix(SldWorks swApp, out List<SolidWorksFixPattern.PartBendInfo> partBendInfos, bool makeFlat)
        {
            var solidWorksMacro = new SolidWorksFixPattern
            {
                SwApp = swApp ?? (SldWorks)Marshal.GetActiveObject("SldWorks.Application")
            };
            solidWorksMacro.FixFlatPattern(makeFlat);
            partBendInfos = solidWorksMacro.PartBendInfos;
        }

        public class SolidWorksFixPattern
        {
            public void FixFlatPattern(bool makeDxf)
            {
                try
                {
                    _swModel = (ModelDoc2)SwApp.ActiveDoc;
                    var getActiveConfig = (Configuration)_swModel.GetActiveConfiguration();
                   _swModel.EditRebuild3();
                    GetBendsInfo(getActiveConfig.Name);
                    FixOneBand(getActiveConfig.Name, makeDxf);
                }
                catch (Exception exception)
                {
                     //MessageBox.Show(exception.StackTrace);
                }
            }

            public SldWorks SwApp;
            ModelDoc2 _swModel;
            Feature _swFeat;
            Feature _swSubFeat;

            public class PartBendInfo
            {
                public string Config { get; set; }
                public string EdgeFlange { get; set; }
                public string OneBend { get; set; }
                public bool IsSupressed { get; set; }
            }

            public List<PartBendInfo> PartBendInfos = new List<PartBendInfo>();
           
            static bool IsSheetFeature(string name)
            {
                switch (name)
                {
                    case "EdgeFlange":
                    case "FlattenBends":
                    case "SMBaseFlange":
                    case "SheetMetal":
                    case "SM3dBend":
                    case "SMMiteredFlange":
                    case "ProcessBends":
                    case "FlatPattern":
                    case "Hem":
                    case "Jog":
                    case "LoftedBend":
                    case "Rip":
                    case "CornerFeat":
                        return true;
                    default: return false;
                        
                }
            }

            private void GetBendsInfo(string config)
            {
                var swPart = (IPartDoc)_swModel;
                _swFeat = (Feature)_swModel.FirstFeature();
                while ((_swFeat != null))
                {
                    if (IsSheetFeature(_swFeat.GetTypeName()))
                    {
                        var parentFeatureName = _swFeat.Name;
                        var stateOfEdgeFlange = IsSuppressedEdgeFlange(parentFeatureName);
                        _swSubFeat = _swFeat.IGetFirstSubFeature();
                        
                        while ((_swSubFeat != null))
                        {
                            if (_swSubFeat.GetTypeName() == "OneBend" || _swSubFeat.GetTypeName() == "SketchBend")
                            {
                                PartBendInfos.Add(new PartBendInfo
                                {
                                    Config = config,
                                    EdgeFlange = parentFeatureName,
                                    OneBend = _swSubFeat.Name,
                                    IsSupressed = stateOfEdgeFlange
                                });

                                _swSubFeat.Select(false);

                                if (stateOfEdgeFlange)
                                {
                                    swPart.EditSuppress();
                                }
                                else
                                {
                                    swPart.EditUnsuppress();
                                }

                                _swSubFeat.DeSelect();


                            }
                            _swSubFeat = (Feature)_swSubFeat.GetNextSubFeature();
                        }
                    }
                    _swFeat = (Feature)_swFeat.GetNextFeature();
                }
            }

            private void FixOneBand(string config, bool makeDxf)
            {
                var swPart = (IPartDoc) _swModel;

                _swFeat = (Feature) _swModel.FirstFeature();
                Feature flatPattern = null;

                while ((_swFeat != null))
                {
                    if (_swFeat.GetTypeName() == "FlatPattern")
                    {
                         flatPattern = _swFeat;
                        flatPattern.Select(true);
                        swPart.EditUnsuppress();
                        flatPattern.DeSelect();

                        _swSubFeat = (Feature)flatPattern.GetFirstSubFeature();

                        while ((_swSubFeat != null))
                        {
                            if (_swSubFeat.GetTypeName() == "UiBend")
                            {
                                try
                                {                                   

                                    var supression =
                                        PartBendInfos.Where(x => x.OneBend == GetOneBandName(_swSubFeat.Name))
                                        .Single(x => x.Config == config)
                                        .IsSupressed;                                    

                                    _swSubFeat.Select(false);                                   

                                    if (supression)
                                    {
                                        swPart.EditSuppress();
                                    }
                                    else
                                    {
                                        swPart.EditUnsuppress();
                                    }

                                    _swSubFeat.DeSelect();

                                  
                                }
                                catch (Exception)
                                {
                                    //MessageBox.Show(exception.StackTrace);
                                }
                            }

                            _swSubFeat = (Feature) _swSubFeat.GetNextSubFeature();
                        }
                    }

                    _swFeat = (Feature) _swFeat.GetNextFeature();
                }

                if (makeDxf)
                {
                    flatPattern?.Select(true);
                    swPart.EditUnsuppress();
                    flatPattern?.DeSelect();
                }
                else
                {
                    flatPattern?.Select(true);
                    swPart.EditSuppress();
                    flatPattern?.DeSelect();
                }

                _swModel.EditRebuild3();
            }

            static string GetOneBandName(string uiName)
            {
                if (!string.IsNullOrEmpty(uiName))
                    return uiName.Substring(
                        uiName.IndexOf('<') + 1,
                        uiName.IndexOf('>') - uiName.IndexOf('<') - 1);
                return null;
            }

            bool IsSuppressedEdgeFlange(string featureName)
            {
                var state = false;
                try
                {
                    _swModel.Extension.SelectByID2(featureName, "BODYFEATURE", 0, 0, 0, false, 0, null, 0);
                    var swSelMgr = (SelectionMgr)_swModel.SelectionManager;
                    var swFeat = (Feature)swSelMgr.GetSelectedObject6(1, -1);
                    var states = (bool[])swFeat.IsSuppressed2(1, _swModel.GetConfigurationNames());
                    state = states[0];
                    _swModel.ClearSelection2(true);
                }
                catch (Exception)
                {
                   
                }
                return state;
            }

        }
    }
}
