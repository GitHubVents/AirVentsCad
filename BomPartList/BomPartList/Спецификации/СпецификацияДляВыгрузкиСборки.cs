using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using BomPartList.Properties;
using EPDM.Interop.epdm;
using VentsPDM_dll;

namespace BomPartList.Спецификации
{

    /// <summary>
    /// 
    /// </summary>
    public class СпецификацияДляВыгрузкиСборки
    {

        #region Поля 

        #region to delete
        private static string _connectionString = Settings.Default.ConnectionToSQL;

        /// <summary>
        /// Gets or sets the connection string.
        /// </summary>
        /// <value>
        /// The connection string.
        /// </value>
        public static string ConnectionString
        {
            get { return _connectionString; }
            set { _connectionString = value; }
        }

        #endregion

        //Data Source=192.168.14.11;Initial Catalog=SWPlusDB;User ID=sa;Password=PDMadmin;Persist Security Info=True;

        /// <summary>
        /// Gets or sets the data source.
        /// </summary>
        /// <value>
        /// The data source.
        /// </value>
        public string DataSource { get; set; }

        /// <summary>
        /// Gets or sets the initial catalog.
        /// </summary>
        /// <value>
        /// The initial catalog.
        /// </value>
        public string InitialCatalog { get; set; }

        /// <summary>
        /// Gets the SQL connection string.
        /// </summary>
        /// <value>
        /// The SQL connection string.
        /// </value>
        //public static string SqlConnectionString
        //{
        //    get
        //    {
        //        var builder = new SqlConnectionStringBuilder();
        //        builder["Data Source"] = "192.168.14.11";
        //        builder["Initial Catalog"] = "SWPlusDB";
        //        builder["User ID"] = "sa";
        //        builder["Password"] = "PDMadmin";
        //        builder["Persist Security Info"] = true;
        //        return builder.ConnectionString;
        //    }
        //}
        

        /// <summary>
        /// Gets or sets the name of the PDM base. For example: "Vents-PDM"
        /// </summary>
        public string repositoryName { get; set; }

        /// <summary>
        /// Gets or sets the name of the user. For example: "kb64"
        /// </summary>
        public string ИмяПользователя { get; set; }

        /// <summary>
        /// Gets or sets the user password. For example: "12345"
        /// </summary>
        public string ПарольПользователя { get; set; }

        readonly IEdmVault10 _mVault = new EdmVault5() as IEdmVault10;
        IEdmFile7 _edmFile7;

        /// <summary>
        /// Поля спецификации
        /// </summary>
        public class UploadData
        {
            public string Раздел { get; set; } 
            public string Обозначение { get; set; }
            public string КодДокумента { get; set; }
            public string Наименование { get; set; }
            public string Количество { get; set; }
            public string КодERP { get; set; }
            public string КодМатериала { get; set; }
            public string КолМатериала { get; set; }
            public string Масса { get; set; }
            public string Версия { get; set; }
            public string Состояние { get; set; }
            public string Конфиг { get; set; }
            public string ПлощадьПокрытия { get; set; }

            public string Путь { get; set; }
            public string Уровень { get; set; }
            
            /// <summary>
            /// Gets or sets the errors.
            /// </summary>
            /// <value>
            /// Поле заполняется по результатам проверки других полей автоматически. Если пусто - все данные устраивают.
            /// </value>
            public string Errors { get; set; }
        }


        /// <summary>
        /// Спецификация Сборки.
        /// </summary>
        /// <param name="pathToAssembly">Путь к Сборке.</param>
        /// <param name="specificationId">id Спецификации.</param>
        /// <param name="configuretion">Конфигурация.</param>
        /// <returns></returns>
        public List<UploadData> Specification(string pathToAssembly, int specificationId, string configuretion)
        {
            ПолучениеСпецификации(pathToAssembly, specificationId, configuretion);
            return _bomList;
        }

        /// <summary>
        /// Спецификация Сборки.
        /// </summary>
        /// <param name="путьКСборке">Путь к Сборке.</param>
        /// <param name="конфигурация">Конфигурация.</param>
        /// <returns></returns>
        public List<UploadData> Specification(string путьКСборке, string конфигурация)
        {
            ПолучениеСпецификации(путьКСборке, 10, конфигурация);
            //return NormalizeBomList(_bomList);
            return _bomList;
        }
       

        #endregion

        #region Методы получения листа

        /// <summary>
        /// Получение пути к сборке.
        /// </summary>
        /// <param name="имяСборки">Имя сборки в базе.</param>
        /// <returns></returns>
        public string ПолучениеПутиКСборке(string имяСборки)
        {
            Отладка(string.Format("Получение пути к сборке по имени {0} в хранилище {1}", имяСборки, repositoryName), "", "ПолучениеПутиКСборке", "СпецификацияДляВыгрузкиСборки");
          
            try
            {
                var ventsPdMdll = new PDM { vaultname = repositoryName };
                var pathToAssembly = ventsPdMdll.SearchDoc(имяСборки).First();
                Отладка(string.Format("Получен путь {2}  к сборке по имени {0} в хранилище {1}", имяСборки, repositoryName, pathToAssembly), "", "ПолучениеПутиКСборке", "СпецификацияДляВыгрузкиСборки");
                return new FileInfo(pathToAssembly.Path).Exists ? pathToAssembly.Path : "";
            }
            catch (Exception exception)
            {
                Ошибка(string.Format("Не удалось найти сборку c именем {0}. Возникла ошибка {1}", имяСборки, exception.StackTrace), exception.ToString(), "ПолучениеПутиКСборке", "СпецификацияДляВыгрузкиСборки");
                return "";
            }
        }
        
        /// <summary>
        /// Получение списка конфигураций по пути к файлу сборки.
        /// </summary>
        /// <param name="путьКСборке">Путь к файлу сборки.</param>
        /// <returns></returns>
        public IEnumerable<string> ПолучениеКонфигураций(string путьКСборке)
        {
            try
            {
                var vault1 = new EdmVault5();
                vault1.LoginAuto(repositoryName, 0);
                IEdmFolder5 oFolder;
                var edmFile5 = vault1.GetFileFromPath(new FileInfo(путьКСборке).FullName, out oFolder);
                var configs = edmFile5.GetConfigurations("");

                var headPosition = configs.GetHeadPosition();

                var configsList = new List<string>();

                while (!headPosition.IsNull)
                {
                    var configName = configs.GetNext(headPosition);
                    if (configName != "@")
                    {
                        configsList.Add(configName);
                    }
                }
                return configsList;
            }
            catch (Exception)
            {
                return null;
            }
        }

        void ПолучениеСпецификации(string путьКСборке, int idСпецификации, string конфигурация)
        {
            Отладка(string.Format("Запущен метод ПолучениеСпецификации(int idСпецификации = {0}, string конфигурация = {1})", idСпецификации, конфигурация), "", "ПолучениеСпецификации", "СпецификацияДляВыгрузкиСборки");
            try
            {
                ВходВХранилище();
            }
            catch (Exception exception)
            {
                Ошибка(exception.StackTrace, exception.ToString(), "ПолучениеПутиКСборке", "СпецификацияДляВыгрузкиСборки");
            }

            if (!ПолучениеСборки(путьКСборке)) return;
            
            Отладка(String.Format("Получение спецификации на заказ из хранилища {0} по id={1} для конфигурации - {2}", repositoryName, idСпецификации, конфигурация), "", "ПолучениеСпецификации", "СпецификацияДляВыгрузкиСборки");

            if (_edmFile7 == null) return;

            EdmBomView bomView = null;

            try
            {
                bomView = _edmFile7.GetComputedBOM(Convert.ToInt32(idСпецификации), Convert.ToInt32(-1), конфигурация,
                (int)EdmBomFlag.EdmBf_ShowSelected);
            }
            catch (Exception)
            {
                Ошибка(String.Format("Получение спецификации на заказ из хранилища {0} по id={1} для конфигурации - {2}", repositoryName, idСпецификации, конфигурация), "", "ПолучениеСпецификации", "СпецификацияДляВыгрузкиСборки");
            }

            

            if (bomView == null) return;
            object [] bomRows;
            EdmBomColumn [] bomColumns;
            bomView.GetRows(out bomRows);
            bomView.GetColumns(out bomColumns);

            var bomTable = new DataTable();

            foreach (EdmBomColumn bomColumn in bomColumns)
            {
                bomTable.Columns.Add(new DataColumn { ColumnName = bomColumn.mbsCaption });
            }

            bomTable.Columns.Add(new DataColumn { ColumnName = "Путь" });
            bomTable.Columns.Add(new DataColumn { ColumnName = "Уровень" });

            for (var i = 0; i < bomRows.Length; i++)
            {
                var cell = (IEdmBomCell)bomRows.GetValue(i);

                bomTable.Rows.Add();
                for (var j = 0; j < bomColumns.Length; j++)
                {
                    var column = (EdmBomColumn)bomColumns.GetValue(j);
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

            Отладка("Спецификация для выгрузки успешно получена", "", "ПолучениеСпецификации", "СпецификацияДляВыгрузкиСборки");
            
        }
        
        void ВходВХранилище()
        {
            Отладка("Осуществление входа в хранилище.", "", "ВходВХранилище", "СпецификацияДляВыгрузкиСборки");
            

            if (_mVault == null) return;
            if (_mVault.IsLoggedIn) goto m1;
            try
            {
                try
                {
                    _mVault.LoginAuto(repositoryName, 0);
                }
                catch (Exception)
                {
                    _mVault.Login(ИмяПользователя, ПарольПользователя, repositoryName);
                }
               
                _mVault.CreateSearch();
                if (_mVault.IsLoggedIn)
                {
                    Отладка("Осуществление входа в хранилище " + repositoryName, "", "ВходВХранилище", "СпецификацияДляВыгрузкиСборки");
                }
            }
            catch (COMException ex)
            {
                Ошибка(string.Format("Ошибка входа в хранилище {0} пользователь - {1} пароль - {2}. {3}", repositoryName, ИмяПользователя, ПарольПользователя, ex.ToString()), ex.ToString(), "ВходВХранилище", "СпецификацияДляВыгрузкиСборки");
            }
            m1:
            Отладка("Осуществление входа в хранилище.", "", "ВходВХранилище", "СпецификацияДляВыгрузкиСборки");
            
        }
        
        bool ПолучениеСборки(string путьКСборке)
        {
            Отладка("Получение файла сборки для дальнейшей обработки.", "", "ПолучениеСборки", "СпецификацияДляВыгрузкиСборки");
            
            if (путьКСборке == null)
            {
                Отладка("Не введен путь к сборке", "", "ПолучениеСборки", "СпецификацияДляВыгрузкиСборки");
                return false;
            }
            //if (путьКСборке.Contains("sldasm"))
            //{
            //    //LoggerInfo("Выбрана не сборка! Выберете файл с расширением .sldasm");
            //    return false;
            //}
            var filename = путьКСборке;
            try
            {
                IEdmFolder5 folder;
                _edmFile7 = (IEdmFile7)_mVault.GetFileFromPath(filename, out folder);
            }
            catch (Exception exception)
            {
                Отладка(exception.StackTrace, "", "ПолучениеСборки", "СпецификацияДляВыгрузкиСборки");
                return false;
            }
            
            Отладка("Файл сборки для дальнейшей обработки получен успешно.", "", "ПолучениеСборки", "СпецификацияДляВыгрузкиСборки");
            
            return true;
        }
        
        #endregion

        #region ДанныеДляВыгрузки - Поля с х-ками 

        List<UploadData> _bomList = new List<UploadData>();
        
        private static List<UploadData> BomTableToBomList(DataTable table)
        {

            Отладка("Заполнение списка из полученой таблицы с количеством колонок:" + table.Columns.Count, "", "ПолучениеСборки", "СпецификацияДляВыгрузкиСборки");
            
            //var countrows = table.Rows.Count;
            var bomList = new List<UploadData>(table.Rows.Count);
            
            bomList.AddRange(from DataRow row in table.Rows
                select row.ItemArray
                into values
                select new UploadData
                {
                    Раздел = values[0].ToString(),
                    Обозначение = values[1].ToString(),
                    КодДокумента = values[2].ToString(),
                    Наименование = values[3].ToString(),
                    Количество = values[4].ToString(),
                    КодERP = values[5].ToString(),
                    КодМатериала = values[6].ToString(),
                    КолМатериала = values[7].ToString(),
                    Масса = values[8].ToString(),
                    Версия = values[9].ToString(),
                    Состояние = values[10].ToString(),
                    Конфиг = values[11].ToString(),
                    ПлощадьПокрытия = values[12].ToString(),
                    Путь = values[13].ToString(),
                    Уровень = values[14].ToString()
                });
            foreach (var bomCells in bomList)
            {
                bomCells.Errors = ErrorMessageForParts(bomCells);
            }

            Отладка("Список из полученой таблицы успешно заполнен элементами в количестве" + bomList.Count, "", "ПолучениеСборки", "СпецификацияДляВыгрузкиСборки");
            
            return bomList;
        }

        private static List<UploadData> NormalizeBomList(List<UploadData> listToNormalize)
        {
            listToNormalize = listToNormalize.OrderBy(x => x.Обозначение).ToList();

            for (var i = 0; i < listToNormalize.Count - 1; i++)
            {
                if (listToNormalize[i].Обозначение != listToNormalize[i + 1].Обозначение) continue;
                if (listToNormalize[i].Конфиг != listToNormalize[i + 1].Конфиг) continue;

                listToNormalize[i].Количество = Convert.ToString(
                    (Convert.ToInt32(listToNormalize[i].Количество) +
                     Convert.ToInt32(listToNormalize[i + 1].Количество)));
                listToNormalize.RemoveAt(i + 1);
            }
            return listToNormalize;
        }
        
        static string ErrorMessageForParts(UploadData данныеДляВыгрузки)
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

            if (данныеДляВыгрузки.Обозначение == "")
            {
                обозначениеErr = "\n Обозначение";
            }

            var regex = new Regex("[^0-9]+");
            if (regex.IsMatch(данныеДляВыгрузки.Конфиг))
            {
                конфигурацияErr = "\n Изменить имя конфигурации на численное значение";
            }

            if (данныеДляВыгрузки.Наименование == "")
            {
                наименованиеErr = "\n Наименование";
            }

            if (данныеДляВыгрузки.Раздел == "")
            {
                разделErr = "\n Раздел";
            }
            
            var message = данныеДляВыгрузки.Errors = String.Format("Необходимо заполнить: {0}{1}{2}{3}{4}{5}",
                обозначениеErr,
                материалЦмиErr,
                наименованиеErr,
                разделErr,
                толщинаЛистаErr,
                конфигурацияErr);

            return данныеДляВыгрузки.Errors == messageErr ? "" : message;
        }

        #endregion

        #region Логгер
        
        static void Отладка(string logText, string код, string функция, string className)
        {
            Лог.Debug(logText, код, функция, className);
        }

        static void Ошибка(string logText, string код, string функция, string className)
        {
            Лог.Error(logText, код, функция, className);
        }

        static void Информация(string logText, string код, string функция, string className)
        {
            Лог.Info(logText, код, функция, className);
        }

        static class Лог
        {
            //private const string ConnectionString = "Data Source=192.168.14.11;Initial Catalog=SWPlusDB;User ID=sa;Password=PDMadmin";//Settings.Default.ConnectionToSQL;//
            //private const string ClassName = "ModelSw";
            //----------------------------------------------------------
            // Статический метод записи строки в файл лога без переноса
            //----------------------------------------------------------
            public static void Write(string text)
            {
                //using (var streamWriter = new StreamWriter("C:\\log.txt", true))  //Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + 
                //{
                //    streamWriter.Write(text);
                //}
            }

            //---------------------------------------------------------
            // Статический метод записи строки в файл лога с переносом
            //---------------------------------------------------------
            public static void WriteLine(string message)
            {
                //using (var streamWriter = new StreamWriter("C:\\log.txt", true))
                //{
                //    streamWriter.WriteLine("{0,-23} {1}", DateTime.Now + ":", message);
                //}
            }


            public static void Debug(string message, string код, string функция, string className)
            {
                //using (var streamWriter = new StreamWriter("C:\\log.txt", true))
                //{
                //    streamWriter.WriteLine("{0,-20}  {2,-7} {3,-20} {1}", DateTime.Now + ":", message, "Error", className);
                //}
                WriteToBase(message, "Debug", код, className, функция);
            }


            public static void Error(string message, string код, string функция, string className)
            {
                //using (var streamWriter = new StreamWriter("C:\\log.txt", true))
                //{
                //    streamWriter.WriteLine("{0,-20}  {2,-7} {3,-20} {1}", DateTime.Now + ":", message, "Error", className);
                //}
                WriteToBase(message, "Error", код, className, функция);
            }

            public static void Info(string message, string код, string функция, string className)
            {
                //using (var streamWriter = new StreamWriter("C:\\log.txt", true))
                //{
                //    streamWriter.WriteLine("{0,-20}  {2,-7} {3,-20} {1}", DateTime.Now + ":", message, "Info", className);
                //}
                WriteToBase(message, "Info", код, className, функция);
            }

            static void WriteToBase(string описание, string тип, string код, string модуль, string функция)
            {
                //Data Source=192.168.14.11;Initial Catalog=SWPlusDB;User ID=sa;Password=PDMadmin

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


                        //var returnParameter = sqlCommand.Parameters.Add("@ProjectNumber", SqlDbType.Int);
                        //returnParameter.Direction = ParameterDirection.ReturnValue;

                        sqlCommand.ExecuteNonQuery();

                        //var result = Convert.ToInt32(returnParameter.Value);

                        //switch (result)
                        //{
                        //    case 0:
                        //        MessageBox.Show("Подбор №" + Номерподбора.Text + " уже существует!");
                        //        break;
                        //}

                    }
                    catch (Exception exception)
                    {
                      //  MessageBox.Show("Введите корректные данные! " + exception.ToString());
                    }
                    finally
                    {
                        con.Close();
                    }
                }
            }
        }



        #endregion
    }
}
