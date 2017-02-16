using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using MakeDxfUpdatePartData;
using VentsMaterials;

namespace BomPartList
{
    /// <summary>
    /// 
    /// </summary>
    public partial class BomPartListClass
    {
        #region Поля 

        /// <summary>
        /// The exception
        /// </summary>
        public ArgumentException Exception;

        readonly SetMaterials _swMaterials = new SetMaterials();

        private readonly ToSQL _connectToSql = new ToSQL();

        /// <summary>
        /// Gets or sets the connection to SQL.
        /// </summary>
        /// <value>
        /// The connection to SQL. Строка подключения к базе данных.
        /// </value>
        public string ConnectionToSql{ get; set; }

        #endregion

        #region Методы получения листа

        /// <summary>
        /// Accepts all changes.
        /// </summary>
        /// <param name="chengedParts">The chenged parts.</param>
        public void AcceptAllChanges(IEnumerable<BomCells> chengedParts)
        {
            ToSQL.Conn = ConnectionToSql;
            foreach (var chengedPart in chengedParts)
            {
                var part = chengedPart;

                var materialId = _connectToSql.GetSheetMetalMaterialsName()
                    .Where(x => x.MatName == part.Материал);
                foreach (var columnMatPropse in materialId)
                {
                    SetMaterialsProperty(chengedPart.Путь, chengedPart.Конфигурация, Convert.ToInt32(columnMatPropse.LevelID));
                }
            }
        }

        /// <summary>
        /// Accepts all changes.
        /// </summary>
        /// <param name="currentBomList">The current bom list.</param>
        /// <param name="newBomList">The new bom list.</param>
        public void AcceptAllChanges(List<BomCells> currentBomList, List<BomCells> newBomList)
        {
            AcceptAllChanges(СравнитьМассивы(currentBomList, newBomList));
            UpdateCutList(currentBomList);
        }

        static IEnumerable<BomCells> СравнитьМассивы(List<BomCells> currentBomList, List<BomCells> newBomList)
        {
            var newList = new List<BomCells>();

            if (newBomList == null) return null;
            newBomList = newBomList.OrderBy(x => x.ОбозначениеDocMgr).ToList();
            if (currentBomList == null) return null;
            currentBomList = currentBomList.OrderBy(x => x.ОбозначениеDocMgr).ToList();

            for (var i = 0; i < currentBomList.Count; i++)
            {
                if (newBomList[i].Материал != currentBomList[i].Материал ||
                    newBomList[i].ТолщинаЛистаDocMgr != currentBomList[i].ТолщинаЛистаDocMgr ||
                    newBomList[i].Конфигурация != currentBomList[i].Конфигурация)
                {
                    newList.Add(newBomList[i]);
                }
            }
            return newList;
        }

        void SetMaterialsProperty(string partPath, string partConfig, int materialId)
        {
            var pdmBase = PdmBaseName;
            
            try
            {
                if (!IsSheetMetalPart(partPath, pdmBase))
                {
                    //MessageBox.Show("Не листовая деталь");
                    return;
                }
            }
            catch (Exception exception)
            {
              Message = "Ошибка во время определения листовой детали" + exception.ToString() + "по пути" + partPath;
              throw Exception = new ArgumentException(exception.ToString()); 
            }

            try
            {
                CheckInOutPdm(partPath, false, pdmBase);
            }
            catch (Exception exception)
            {
                Message = "Ошибка во время разрегистрации " + exception.ToString() + "по пути" + partPath;
                throw Exception = new ArgumentException(exception.ToString()); 
            }

            try
            {
              //  _swMaterials.ApplyMaterial(partPath, partConfig, materialId);
            }
            catch (Exception exception)
            {
                Message = "Ошибка во время применения материала " + exception.ToString() + "по пути" + partPath;
                throw Exception = new ArgumentException(exception.ToString()); 
            }

            finally
            {
                try
                {
                    CheckInOutPdm(partPath, true, pdmBase);
                }
                catch (Exception exception)
                {
                    Message = "Ошибка во время регистрации " + exception.ToString() + "по пути" + partPath;
                    throw Exception;
                }
            }
        }

        void UpdateCutList(IEnumerable<BomCells> currentBomList)
        {
            var partDataClass = new MakeDxfExportPartDataClass {PdmBaseName = PdmBaseName};
            foreach (var selectedPrt in currentBomList)
            {
                try
                {
                    // TODO: Проверку по бызе о наличии CutList'а

                    const string xmlPath = @"\\srvkb\SolidWorks Admin\XML";

                    if (new FileInfo(xmlPath +"\\"+ Path.GetFileNameWithoutExtension(selectedPrt.Путь)+".xml").Exists)
                    {
                        Message = String.Format("Файл {0} уже есть в базе ", Path.GetFileNameWithoutExtension(selectedPrt.Путь)+".xml");
                    }
                    else
                    {
                       // partDataClass.CreateFlattPatternUpdateCutlist(selectedPrt.Путь, false);
                    }
                        
                 
                        //partDataClass.CreateFlattPatternUpdateCutlist(selectedPrt.Путь, true);
                }
                catch (Exception)
                {
                    return;
                }
            }
        }

        #endregion


    }
}
