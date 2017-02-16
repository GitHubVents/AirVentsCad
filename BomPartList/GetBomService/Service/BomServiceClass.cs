using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.ServiceModel;
using BomPartList;
using EdmLib;

namespace GetBomService.Service
{
    public class BomServiceClass
    {
        [ServiceContract]
        public interface IBomTableService
        {
            [OperationContract]
            string PathById(int id);

            [OperationContract]
            string PathByNameAsm(string name);

            [OperationContract]
            IEnumerable<string> AsmNames();

            [OperationContract]
            IEnumerable<string> SearchAsmNames(string someLettersInName);

            [OperationContract]
            IList<BomPartListClass.BomCells> Bom(int type, string assemblyPath);

            [OperationContract]
            IList<BomPartListClass.BomCells> BomParts(string assemblyPath, string pdmBaseName, string userName,
                string password, string constring);
        }

        public class BomTableService : IBomTableService
        {
            #region Fields

            const string Constring = "Data Source=192.168.14.11;Initial Catalog=SWPlusDB;User ID=sa;Password=PDMadmin"; 
            const string PdmBaseName = "Tets_debag";// "Vents-PDM"
            const string UserName = "kb81";
            const string Password = "1";
            const int BomId = 7;//8

            #endregion

            public string PathById(int id)
            {
                try
                {
                    var vault1 = new EdmVault5();
                    vault1.LoginAuto(PdmBaseName, 0);

                    var search = vault1.CreateSearch();
                    search.FindFiles = true;
                    search.FindFolders = false;
                    var edmFile5 = vault1.GetObject(EdmObjectType.EdmObject_File, id);

                    search.FileName = edmFile5.Name;
                    var result = search.GetFirstResult();
                    return string.Format("Path to file - {0}", result.Path);
                }
                catch (Exception)
                {
                    return "Ошибка во время подключения!";
                }
            }

            public string PathByNameAsm(string name)
            {
                try
                {
                    var vault1 = new EdmVault5();
                    vault1.LoginAuto(PdmBaseName, 0);

                    var search = vault1.CreateSearch();
                    search.FindFiles = true;
                    search.FindFolders = false;

                    try
                    {
                        search.FileName = name + ".SLDASM";
                        var result = search.GetFirstResult();
                        return result.Path;
                    }
                    catch (Exception)
                    {
                        search.FileName = name;
                        var result = search.GetFirstResult();
                        return result.Path;
                    }
                }

                catch (Exception)
                {
                    return "Ошибка во время подключения!";
                }

            }

            public IEnumerable<string> SearchAsmNames(string someLettersInName)
            {
                try
                {
                    var vault1 = new EdmVault5();
                    vault1.LoginAuto(PdmBaseName, 0);

                    var search = vault1.CreateSearch();
                    search.FindFiles = true;
                    search.FindFolders = false;

                    var list = new List<string>();

                    search.FileName = someLettersInName;
                    list.Add(search.GetFirstResult().Name);


                    while (search.GetNextResult() != null)
                    {
                        list.Add(search.GetNextResult().Name);
                    }

                    return list.AsEnumerable();
                }

                catch (Exception)
                {
                    return null;
                }

            }

            public IEnumerable<string> AsmNames()
            {
                try
                {
                    var vault1 = new EdmVault5();
                    vault1.LoginAuto(PdmBaseName, 0);
                    var searchDir = new DirectoryInfo(vault1.RootFolderPath);
                    var files = searchDir.GetFiles("*.sldasm", SearchOption.AllDirectories);
                    return files.Select(file => Path.GetFileNameWithoutExtension(file.FullName));
                }
                catch (Exception)
                {
                    return null;
                }
            }

            public IList<BomPartListClass.BomCells> Bom(int type, string assemblyPath)
            {
                var bomClass = new BomPartListClass
                {
                    ConnectionToSql = Constring,
                    BomId = BomId,
                    PdmBaseName = PdmBaseName,
                    UserName = UserName,
                    UserPassword = Password,
                    AssemblyPath = assemblyPath,
                };

                switch (type)
                {
                    case 0:
                        return bomClass.BomListAsm().ToList();
                    case 1:
                        return bomClass.BomListPrt(true).ToList();
                    case 2:
                        return bomClass.BomListPrt(false).ToList();
                    default:
                        return bomClass.BomList();
                }
            }

            public IList<BomPartListClass.BomCells> BomParts(string assemblyPath, string pdmBaseName, string userName, string password, string constring)
            {
                var bomId = 7;
                if (pdmBaseName == "Vents-PDM")
                {
                    bomId = 8;
                }

                var bomClass = new BomPartListClass
                {
                    ConnectionToSql = constring,
                    BomId = bomId,
                    PdmBaseName = pdmBaseName,
                    UserName = userName,
                    UserPassword = password,
                    AssemblyPath = assemblyPath
                };

                return bomClass.BomListPrt(true).ToList();
            }

        }
    }
}
