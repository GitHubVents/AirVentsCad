using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using EPDM.Interop.epdm;

namespace BomPartList
{
    /// <summary>
    /// 
    /// </summary>
    public partial class BomPartListClass
    {
        #region Поля 

        /// <summary>
        /// Gets or sets the name of the PDM base. For example: "Vents-PDM"
        /// </summary>
        public string PdmBaseName { get; set; }

        /// <summary>
        /// Gets or sets the bom identifier.
        /// </summary>
        /// <value>
        /// The bom identifier. Для "Vents-Pdm" - 8
        /// </value>
        public int BomId { get; set; }

        /// <summary>
        /// Gets or sets the name of the user. For example: "kb64"
        /// </summary>
        public string UserName { get; set; }

        /// <summary>
        /// Gets or sets the user password. For example: "12345"
        /// </summary>
        public string UserPassword { get; set; }
        
        /// <summary>
        /// The message
        /// </summary>
        public string Message;

        /// <summary>
        /// The assembly information label
        /// </summary>
        public string AssemblyInfoLabel;

        /// <summary>
        /// Gets or sets the assembly path.
        /// </summary>
        /// <value>
        /// The assembly fullpath in pdm.
        /// </value>
        public string AssemblyPath { get; set; }
        
        /// <summary>
        /// Gets or sets the asm configuration.
        /// </summary>
        /// <value>
        /// The asm configuration.
        /// </value>
        public string AsmConfiguration { get; set; }

        readonly IEdmVault10 _mVault = new  EdmVault5() as IEdmVault10;
        IEdmFile7 _edmFile7;

        /// <summary>
        /// Поля спецификации
        /// </summary>
        public class BomCells
        {
            // Pdm
            public string Количество { get; set; }
            public string ТипФайла { get; set; }
            public string Конфигурация { get; set; }
            public string ПоследняяВерсия { get; set; }

            public string Путь { get; set; }
            public string Уровень { get; set; }
            public string Состояние { get; set; }

            public string Раздел { get; set; }
            public string Обозначение { get; set; }
            public string Наименование { get; set; }
            public string Материал { get; set; }
            public string МатериалЦми { get; set; }
            public string ТолщинаЛиста { get; set; }
            public string ERPCode { get; set; }

            public string КодERP { get; set; }
            
            /// <summary>
            /// Gets or sets the errors.
            /// </summary>
            /// <value>
            /// Поле заполняется по результатам проверки других полей автоматически. Если пусто - все данные устраивают.
            /// </value>
            public string Errors { get; set; }

            // DocMgr
            public string РазделDocMgr { get; set; }
            public string ОбозначениеDocMgr { get; set; }
            public string НаименованиеDocMgr { get; set; }
            public string МатериалЦмиDocMgr { get; set; }
            public string ТолщинаЛистаDocMgr { get; set; }
            public string MaterialDocMgr { get; set; }

            public string МассаDocMgr { get; set; }
            public string НаименованиеERPDocMgr { get; set; }
            public string КодМатериалаDocMgr { get; set; }
            public string КодERPDocMgr { get; set; }
            public string КоличествоМатериалаDocMgr { get; set; }
            public string КодДокументаDocMgr { get; set; }
        }

        /// <summary>
        /// Полная спецификация сборки
        /// </summary>
        /// <returns></returns>
        public List<BomCells> BomList()
        {
           // LoggerInfo(string.Format("Получение спецификации на заказ для {0}, конфигурация - {1}", AssemblyPath, AsmConfiguration));
            try
            {
                Login();
            }
            catch (Exception exception)
            {
            //    LoggerError(exception.ToString());
                Message = "Ошибка входа в хранилище " + exception.ToString();
            }
            
            if (GetFile(AssemblyPath))
            {
                Getbom(BomId, AsmConfiguration);
           //     LoggerInfo("Cпецификация на заказ успешно получена.");

            }
            return _bomList;
        }

        /// <summary>
        /// Boms the list.
        /// </summary>
        /// <param name="config">The configuration.</param>
        /// <returns></returns>
        public List<BomCells> BomList(string config)
        {
          //  LoggerInfo(string.Format("Получение спецификации на заказ для {0}, конфигурация - {1}", AssemblyPath, config));
            try
            {
                Login();
            }
            catch (Exception exception)
            {
          //      LoggerError(exception.ToString());
                Message = "Ошибка входа в хранилище " + exception.ToString();
            }

            if (GetFile(AssemblyPath))
            {
                Getbom(BomId, config);
          //      LoggerInfo("Cпецификация на заказ успешно получена.");
            }
            return _bomList;
        }

        /// <summary>
        /// Boms the list updated with DocMgr .
        /// </summary>
        /// <returns></returns>
        public List<BomCells> BomListDocMgr()
        {
         //   LoggerInfo(string.Format("Получение дополненой из Document Manager'a спецификации на заказ для {0}, конфигурация - {1}", AssemblyPath, AsmConfiguration));
            
            try
            {
                Login();
            }
            catch (Exception exception)
            {
          //      LoggerError(exception.ToString());
                Message = "Ошибка входа в хранилище " + exception.ToString();
            }

            if (GetFile(AssemblyPath))
            {
                UpdateBomListFromDocMgr(BomList());
           //     LoggerInfo("Дополненная из Document Manager'a спецификация на заказ успешно получена.");
            }
            return _bomList;
        }

        /// <summary>
        /// Boms the list document MGR only all asms.
        /// </summary>
        /// <returns></returns>
        public List<BomCells> BomListDocMgrOnlyAllAsms()
        {
           // LoggerInfo(string.Format("Получение списка всех подсборок для сборки по пути {0}", AssemblyPath));

            try
            {
                Login();
            }
            catch (Exception exception)
            {
          //      LoggerError(exception.ToString());
                Message = "Ошибка входа в хранилище " + exception.ToString();
            }

            if (GetFile(AssemblyPath))
            {
                UpdateBomListFromDocMgr(BomList().Where(x => x.ТипФайла == "sldasm").Where(x => x.Раздел == "" || x.Раздел == "Сборочные единицы").OrderBy(x => x.ОбозначениеDocMgr).ToList());
         //       LoggerInfo("Список всех подсборок успешно получена.");
            }
            return _bomList;
        }

        /// <summary>
        /// Boms the list updated with DocMgr .
        /// </summary>
        /// <returns></returns>
        //public List<ДанныеДляВыгрузки> BomListDocMgrTopLevelFiltered(string config)
        //{
        //    LoggerInfo(string.Format("Получение списка компонентов верхнего уровня для сборки по пути {0}, конфигурация {1}", ПутьКСборке, config));

        //    try
        //    {
        //        Login();
        //    }
        //    catch (Exception exception)
        //    {
        //        Message = "Ошибка входа в хранилище " + exception.ToString();
        //    }

        //    if (GetFile(ПутьКСборке))
        //    {
        //        UpdateBomListFromDocMgr(BomList(config).Where(x => x.Уровень == "1" || x.Уровень == "0").ToList());
        //    }
        //    return _bomList;
        //}

        /// <summary>
        /// Boms the list updated with DocMgr .
        /// </summary>
        /// <returns></returns>
        public List<BomCells> BomListDocMgrTopLevel(string config)
        {
          //  LoggerInfo(string.Format("Получение списка компонентов верхнего уровня для сборки по пути {0}, конфигурация {1}", AssemblyPath, config));
            try
            {
                Login();
            }
            catch (Exception exception)
            {
                Message = "Ошибка входа в хранилище " + exception.ToString();
         //       LoggerError(exception.ToString());
            }

            if (GetFile(AssemblyPath))
            {
                UpdateBomListFromDocMgr(BomList(config).Where(x => x.Уровень == "1" || x.Уровень == "0").ToList());
          //      LoggerInfo("Список компонентов верхнего уровня для сборки успешно получена.");
            }
            return _bomList;
        }

        /// <summary>
        /// Список сборочных единиц в сборке.
        /// </summary>
        /// <returns></returns>
        public List<BomCells> BomListAsm()
        {
            return BomList().Where(x => x.ТипФайла == "sldasm" & x.Уровень != "0" & x.Раздел == "Сборочные единицы").OrderBy(x => x.Обозначение).ToList();
        }

        /// <summary>
        /// Список деталей в сборке. 
        /// </summary>
        /// <param name="onlySheetMetal">if set to <c>true</c> [только интересующие детали - изготавливаемые на заводе*].</param>
        /// <returns></returns>
        public List<BomCells> BomListPrt(bool onlySheetMetal)
        {
         //   LoggerInfo(string.Format("Получение списка информации по деталям для сборки по пути {0}", AssemblyPath));
            return
                      UpdateBomListFromDocMgr(BomList().ToList());

            //NormalizeBomList(
            //NormalizeBomList(
            //NormalizeBomList(
            //NormalizeBomList(
            //NormalizeBomList(
            //UpdateBomListFromDocMgr(onlySheetMetal ? (BomList().Where(x => x.ТипФайла == "sldprt" & x.Раздел == "" || x.ТипФайла == "sldprt" & x.Раздел == "Детали")).OrderBy(x => x.Обозначение).ToList()
            //: (BomList().Where(x => x.ТипФайла == "sldprt").Where(x => x.Раздел == "" || x.Раздел == "Детали")).OrderBy(x => x.Обозначение).ToList()))))));

            //return BomList().Where(x => x.ТипФайла == "sldprt" & x.Раздел == "" || x.Раздел == "Детали").OrderBy(x => x.Обозначение).ToList();
        }

        #endregion

        #region Методы получения листа

        void Login()
        {
       //     LoggerInfo("Осуществление входа в хранилище.");
            if (_mVault == null) return;
            if (_mVault.IsLoggedIn) goto m1;
            try
            {
                try
                {
                    _mVault.LoginAuto(PdmBaseName, 0);
                }
                catch (Exception)
                {
                    _mVault.Login(UserName, UserPassword, PdmBaseName);
                }
               
                _mVault.CreateSearch();
                if (_mVault.IsLoggedIn)
                {
                    Message = "Logged in " + PdmBaseName;
                }
            }
            catch (COMException ex)
            {
             //   LoggerError(string.Format("Ошибка входа в хранилище {0} пользователь - {1} пароль - {2}. {3}", PdmBaseName, UserName, UserPassword, ex.ToString()));
                Message = "Failed to connect to " + PdmBaseName + " - " + ex.ToString();
            }
            m1:
            Message = "System.Exception";
            //  LoggerInfo("Вход в хранилище осуществлен.");
        }
        
        bool GetFile(string assemblyPath)
        {
          //  LoggerInfo("Получение файла сборки для дальнейшей обработки.");
            if (assemblyPath == null)
            {
                Message = "Введите путь к сборке";
           //     LoggerInfo("Введите путь к сборке");
                return false;
            }
            if (assemblyPath.Contains("sldasm"))
            {
                Message = "Выбрана не сборка! Выберете файл с расширением .sldasm";
           //     LoggerInfo("Выбрана не сборка! Выберете файл с расширением .sldasm");
                return false;
            }
            var filename = assemblyPath;
            try
            {
                IEdmFolder5 folder;
                _edmFile7 = (IEdmFile7)_mVault.GetFileFromPath(filename, out folder);
            }
            catch (Exception exception)
            {
                Message = exception.ToString();
                return false;
            }
         //   LoggerInfo("Файл сборки для дальнейшей обработки получен успешно.");
            return true;
        }

        void Getbom(int bomId, string bomConfiguration)
        {
         //   LoggerInfo(String.Format("Получение спецификации на заказ из хранилища {0} по id={1} для конфигурации - {2}", PdmBaseName, bomId, bomConfiguration));

            if (_edmFile7 == null) return;

            var bomView = _edmFile7.GetComputedBOM(Convert.ToInt32(bomId), Convert.ToInt32(-1), bomConfiguration,
                (int) EdmBomFlag.EdmBf_ShowSelected);

            if (bomView == null) return;
            object[] bomRows;
           EdmBomColumn[] bomColumns;
            bomView.GetRows(out bomRows);
            bomView.GetColumns(out bomColumns);

            var bomTable = new DataTable();

            foreach (EdmBomColumn bomColumn in bomColumns)
            {
                bomTable.Columns.Add(new DataColumn {ColumnName = bomColumn.mbsCaption});
            }

            bomTable.Columns.Add(new DataColumn{ ColumnName = "Путь" });
            bomTable.Columns.Add(new DataColumn{ ColumnName = "Уровень" });

            for (var i = 0; i < bomRows.Length; i++)
            {
                var cell = (IEdmBomCell) bomRows.GetValue(i);

                bomTable.Rows.Add();
                for (var j = 0; j < bomColumns.Length; j++)
                {
                    var column = (EdmBomColumn) bomColumns.GetValue(j);
                    object value;
                    object computedValue;
                    String config;
                    bool readOnly;
                    cell.GetVar(column.mlVariableID, column.meType, out value, out computedValue, out config,
                        out readOnly);
                    if (value != null)
                    {
                        bomTable.Rows[i][j] = value;
                    }
                    bomTable.Rows[i][j + 1] = cell.GetPathName();
                    bomTable.Rows[i][j + 2] = cell.GetTreeLevel();
                }
            }

            _bomList = BomTableToBomList(bomTable);

        //    LoggerInfo("Спецификация на заказ успешно получена");

            #region LabelContent

            var labelContentList = _bomList.Where(x => x.Уровень == "0").ToList();
            
            foreach (var bomCells in labelContentList)
            {
                AssemblyInfoLabel = " Сборка - " + Path.GetFileNameWithoutExtension(bomCells.Путь) +
                    " " + bomCells.Наименование + 
                    " (конфигурация - " + bomCells.Конфигурация + ")";
            }

            #endregion

        }
        
        #endregion

        #region ДанныеДляВыгрузки - Поля с х-ками деталей

        List<BomCells> _bomList = new List<BomCells>();

        readonly static SwDocMgr Swdocmgr = new SwDocMgr();
        static string _properetyValue;
        static string _properetyType;
             
        private static List<BomCells> BomTableToBomList(DataTable table)
        {
            //LoggerInfo("Заполнение списка из полученой таблицы с количеством колонок:" + table.Rows.Count);
            //var countrows = table.Rows.Count;
            var bomList = new List<BomCells>(table.Rows.Count);

            //foreach (var bomCellse in bomList)
            //{
            //    //bomCellse.Раздел = table.Rows..ColumnName[""]
            //}

            for (var i = 0; i < table.Rows.Count; i++)
            {
                //table.Rows[i].
            }

            bomList.AddRange(from DataRow row in table.Rows
                select row.ItemArray
                into values
                select new BomCells
                {
                    Раздел = values[0].ToString(),
                    Обозначение = values[1].ToString(),
                    Наименование = values[2].ToString(),
                    Материал = values[3].ToString(),
                    МатериалЦми = values[4].ToString(),
                    ТолщинаЛиста = values[5].ToString(),
                    Количество = values[6].ToString(),
                    ТипФайла = values[7].ToString(),
                    Конфигурация = values[8].ToString(),
                    ПоследняяВерсия = values[9].ToString(),
                    ERPCode = values[10].ToString(),
                    Состояние = values[11].ToString(),
                    Путь = values[12].ToString(),
                    Уровень = values[13].ToString()
                });
            foreach (var bomCells in bomList)
            {
                bomCells.Errors = ErrorMessageForParts(bomCells);
            }
            //LoggerInfo("Список из полученой таблицы успешно заполнен элементами в количестве" + bomList.Count);
            return bomList;
        }

        private static List<BomCells> NormalizeBomList(List<BomCells> listToNormalize)
        {
            //LoggerInfo("Нормализация списка(объединение повторяющихся строк).");
            listToNormalize = listToNormalize.OrderBy(x => x.ОбозначениеDocMgr).ToList();

            for (var i = 0; i < listToNormalize.Count - 1; i++)
            {
                if (listToNormalize[i].ОбозначениеDocMgr != listToNormalize[i + 1].ОбозначениеDocMgr) continue;
                if (listToNormalize[i].Конфигурация != listToNormalize[i + 1].Конфигурация) continue;

                listToNormalize[i].Количество = Convert.ToString(
                    (Convert.ToInt32(listToNormalize[i].Количество) +
                     Convert.ToInt32(listToNormalize[i + 1].Количество)));
                listToNormalize.RemoveAt(i + 1);
            }
            return listToNormalize;
        }

        private List<BomCells> UpdateBomListFromDocMgr(List<BomCells> listToUpdate)
        {
         //   LoggerInfo("Дополнение списка даными из Document Menedger'a");

            var count = 0;
            var count2 = 0;
            foreach (var bomCells in listToUpdate)
            {
                #region "Материал"

                try
                {
                    Swdocmgr.GetResolvedConfigCustomProperty(bomCells.Путь, bomCells.Конфигурация, "Материал", out _properetyValue, out _properetyType);
                    bomCells.MaterialDocMgr = _properetyValue;
                }
                catch (Exception)
                {
                    count = count + 1;
                    bomCells.MaterialDocMgr = "";// "Поле отсутствует - " + count;
                }

                #endregion
                
                #region Раздел

                try
                {
                    try
                    {
                        Swdocmgr.GetCustomProperty(bomCells.Путь, "Раздел", out _properetyValue, out _properetyType);
                    }
                    catch (Exception)
                    {
                        Swdocmgr.GetConfigCustomProperty(bomCells.Путь, bomCells.Конфигурация, "Раздел", out _properetyValue, out _properetyType);
                    }
                    
                    bomCells.РазделDocMgr = _properetyValue;
                }
                catch (Exception)
                {
                    count2 = count2 + 1;
                    bomCells.РазделDocMgr = "";//"Поле отсутствует - " + count;
                }
                
                #endregion

                #region Обозначение

                try
                {
                    try
                    {
                        Swdocmgr.GetResolvedConfigCustomProperty(bomCells.Путь, bomCells.Конфигурация, "Обозначение", out _properetyValue, out _properetyType);
                       // Swdocmgr.GetCustomProperty(bomCells.Путь, "Обозначение", out _properetyValue, out _properetyType);
                    }
                    catch (Exception)
                    {
                        Swdocmgr.GetConfigCustomProperty(bomCells.Путь, bomCells.Конфигурация, "Обозначение", out _properetyValue, out _properetyType);
                    }

                    bomCells.ОбозначениеDocMgr = _properetyValue;
                }
                catch (Exception)
                {
                    count2 = count2 + 1;
                    bomCells.ОбозначениеDocMgr = "";//"Поле отсутствует - " + count;
                }

                #endregion

                #region Толщина листового металла

                try
                {
                    try
                    {
                        Swdocmgr.GetResolvedConfigCustomProperty(bomCells.Путь, bomCells.Конфигурация, "Толщина листового металла", out _properetyValue, out _properetyType);
                        // Swdocmgr.GetCustomProperty(bomCells.Путь, "Обозначение", out _properetyValue, out _properetyType);
                    }
                    catch (Exception)
                    {
                        Swdocmgr.GetConfigCustomProperty(bomCells.Путь, bomCells.Конфигурация, "Толщина листового металла", out _properetyValue, out _properetyType);
                    }

                    bomCells.ТолщинаЛиста = _properetyValue;
                }
                catch (Exception)
                {
                    count2 = count2 + 1;
                    bomCells.ТолщинаЛиста = "";//"Поле отсутствует - " + count;
                }

                #endregion
                
                #region Наименование

                try
                {
                    try
                    {
                        Swdocmgr.GetCustomProperty(bomCells.Путь, "Наименование", out _properetyValue, out _properetyType);
                    }
                    catch (Exception)
                    {
                        Swdocmgr.GetConfigCustomProperty(bomCells.Путь, bomCells.Конфигурация, "Наименование", out _properetyValue, out _properetyType);
                    }

                    bomCells.НаименованиеDocMgr = _properetyValue;
                }
                catch (Exception)
                {
                    count2 = count2 + 1;
                    bomCells.НаименованиеDocMgr = "";//"Поле отсутствует - " + count;
                }

                #endregion

                #region Масса

                try
                {
                    try
                    {
                        Swdocmgr.GetResolvedConfigCustomProperty(bomCells.Путь, bomCells.Конфигурация, "Масса_Таблица", out _properetyValue, out _properetyType);
                    }
                    catch (Exception)
                    {
                        Swdocmgr.GetConfigCustomProperty(bomCells.Путь, bomCells.Конфигурация, "Масса_Таблица", out _properetyValue, out _properetyType);
                    }

                    bomCells.МассаDocMgr = _properetyValue;
                }
                catch (Exception)
                {
                    bomCells.МассаDocMgr = "";
                }

                #endregion

                #region ERP_Code

                try
                {
                    try
                    {
                        Swdocmgr.GetResolvedConfigCustomProperty(bomCells.Путь, bomCells.Конфигурация, "ERP code", out _properetyValue, out _properetyType);
                    }
                    catch (Exception)
                    {
                        Swdocmgr.GetConfigCustomProperty(bomCells.Путь, bomCells.Конфигурация, "ERP code", out _properetyValue, out _properetyType);
                    }

                    bomCells.КодERPDocMgr = _properetyValue;
                }
                catch (Exception)
                {
                    bomCells.КодERPDocMgr = "";
                }

                #endregion

                #region Код материала

                try
                {
                    try
                    {
                        Swdocmgr.GetResolvedConfigCustomProperty(bomCells.Путь, bomCells.Конфигурация, "Код материала", out _properetyValue, out _properetyType);
                    }
                    catch (Exception)
                    {
                        Swdocmgr.GetConfigCustomProperty(bomCells.Путь, bomCells.Конфигурация, "Код материала", out _properetyValue, out _properetyType);
                    }

                    bomCells.КодМатериалаDocMgr = _properetyValue;
                }
                catch (Exception)
                {
                    bomCells.КодМатериалаDocMgr = "";
                }

                #endregion

                #region Количество материала

                try
                {
                    try
                    {
                        Swdocmgr.GetResolvedConfigCustomProperty(bomCells.Путь, bomCells.Конфигурация, "Количество", out _properetyValue, out _properetyType);
                    }
                    catch (Exception)
                    {
                        Swdocmgr.GetConfigCustomProperty(bomCells.Путь, bomCells.Конфигурация, "Количество", out _properetyValue, out _properetyType);
                    }

                    bomCells.КоличествоМатериалаDocMgr = _properetyValue;
                }
                catch (Exception)
                {
                    bomCells.КоличествоМатериалаDocMgr = "";
                }

                #endregion
               

                #region Код документа

                try
                {
                    try
                    {
                        Swdocmgr.GetResolvedConfigCustomProperty(bomCells.Путь, bomCells.Конфигурация, "Сборка1_ФБ", out _properetyValue, out _properetyType);
                    }
                    catch (Exception)
                    {
                        Swdocmgr.GetConfigCustomProperty(bomCells.Путь, bomCells.Конфигурация, "Сборка1_ФБ", out _properetyValue, out _properetyType);
                    }

                    bomCells.КодДокументаDocMgr = _properetyValue;
                }
                catch (Exception)
                {
                    bomCells.КодДокументаDocMgr = "";
                }

                #endregion

                #region Состояние PDM

                try
                {
                    IEdmFolder5 folder;
                    IEdmVault10 mVault = new EdmVault5() as IEdmVault10;
                    mVault.LoginAuto(PdmBaseName, 0);
                    var edmFile7 = (IEdmFile7)mVault.GetFileFromPath(bomCells.Путь, out folder);
                    _properetyValue = edmFile7.CurrentState.Name;
                    bomCells.Состояние = _properetyValue;
                    
                }
                catch (Exception)
                {
                    bomCells.Состояние = "";
                }

                #endregion
                
            }

            foreach (var bomCells in listToUpdate)
            {
                bomCells.Errors = ErrorMessageForPartsDocMgr(bomCells);
            }

       //     LoggerInfo("Завершено дополнение списка даными из Document Menedger'a");

            return listToUpdate;
        }

        static string ErrorMessageForPartsDocMgr(BomCells bomCells)
        {
          //  LoggerInfo("Дополнение колонки Errors списка данными об необходимых исправлениях.");

            var обозначениеErr = "";
            var материалЦмиErr = "";
            var материалErr = "";
            var наименованиеErr = "";
            var разделErr = "";
            var толщинаЛистаErr = "";
            var конфигурацияErr = "";

            var messageErr = String.Format("Необходимо заполнить: {0}{1}{2}{3}{4}{5}{6}",
               обозначениеErr,
                материалErr,
               материалЦмиErr,
               наименованиеErr,
               разделErr,
               толщинаЛистаErr,
               конфигурацияErr);

            if (bomCells.ОбозначениеDocMgr == "")
            {
                обозначениеErr = "\n Обозначение";
            }

            var regex = new Regex("[^0-9]+");
            if (regex.IsMatch(bomCells.Конфигурация))
            {
                конфигурацияErr = "\n Изменить имя конфигурации на численное значение";
            }

            if (bomCells.НаименованиеDocMgr == "")
            {
                наименованиеErr = "\n Наименование";
            }

            if (bomCells.РазделDocMgr == "")
            {
                разделErr = "\n Раздел";
            }


            if (bomCells.ТипФайла == "sldprt")
            {

                if (bomCells.ТолщинаЛиста == "")
                {
                    толщинаЛистаErr = "\n ТолщинаЛиста";
                }

                //if (bomCells.МатериалЦми == "")
                //{
                //    материалЦмиErr = "\n Материал Цми";
                //}

                if (bomCells.MaterialDocMgr == "")
                {
                    материалErr = "\n Материал";
                }

            }

            var message = bomCells.Errors = String.Format("Необходимо заполнить: {0}{1}{2}{3}{4}{5}{6}",
                обозначениеErr,
                материалErr,
                материалЦмиErr,
                наименованиеErr,
                разделErr,
                толщинаЛистаErr,
                конфигурацияErr);

            return bomCells.Errors == messageErr ? "" : message;
        }

        static string ErrorMessageForParts(BomCells bomCells)
        {
            var обозначениеErr = "";
            var материалЦмиErr = "";
            var наименованиеErr = "";
            var разделErr = "";
            var толщинаЛистаErr = "";
            var конфигурацияErr = "";

            var messageErr = String.Format("Необходимо заполнить: {0}{1}{2}{3}{4}{5}",
               обозначениеErr,
               материалЦмиErr,
               наименованиеErr,
               разделErr,
               толщинаЛистаErr,
               конфигурацияErr);

            if (bomCells.Обозначение == "")
            {
                обозначениеErr = "\n Обозначение";
            }

            var regex = new Regex("[^0-9]+");
            if (regex.IsMatch(bomCells.Конфигурация))
            {
                конфигурацияErr = "\n Изменить имя конфигурации на численное значение";
            }

            if (bomCells.Наименование == "")
            {
                наименованиеErr = "\n Наименование";
            }

            if (bomCells.Раздел == "")
            {
                разделErr = "\n Раздел";
            }

            if (bomCells.ТипФайла == "sldprt")
            {

                if (bomCells.ТолщинаЛиста == "")
                {
                    толщинаЛистаErr = "\n ТолщинаЛиста";
                }

                //if (bomCells.МатериалЦми == "")
                //{
                //    материалЦмиErr = "\n Материал Цми";
                //}
            }

            var message = bomCells.Errors = String.Format("Необходимо заполнить: {0}{1}{2}{3}{4}{5}",
                обозначениеErr,
                материалЦмиErr,
                наименованиеErr,
                разделErr,
                толщинаЛистаErr,
                конфигурацияErr);

            return bomCells.Errors == messageErr ? "" : message;
        }

        #endregion

    }
}
