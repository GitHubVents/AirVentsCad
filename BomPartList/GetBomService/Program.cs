using System;
using System.ServiceModel;
using System.ServiceModel.Description;
using GetBomService.Service;


namespace GetBomService
{
    #region Service

    //[ServiceContract]
    //public interface IBomTableService
    //{
    //    [OperationContract]
    //    string PathById(int id);

    //    [OperationContract]
    //    string PathByNameAsm(string name);
        
    //    [OperationContract]
    //    IEnumerable<string> AsmNames();

    //    [OperationContract]
    //    IList<BomPartListClass.ДанныеДляВыгрузки> Bom(int type, string assemblyPath);

    //    [OperationContract]
    //    IList<BomPartListClass.ДанныеДляВыгрузки> BomParts(string assemblyPath, string pdmBaseName, string userName,
    //        string password, string constring);
    //}

    //public class BomTableService : IBomTableService
    //{
    //    #region Fields

    //    const string Constring = "Data Source=192.168.14.11;Initial Catalog=SWPlusDB;User ID=sa;Password=PDMadmin";    //srvkb
    //    const string ИмяХранилища = "Tets_debag";// "Vents-PDM"
    //    const string ИмяПользователя = "kb81";
    //    const string Password = "1";
    //    const int IdСпецификации = 7;//8

    //    #endregion

    //    public string PathById(int id)
    //    {
    //        try
    //        {
    //            var vault1 = new EdmVault5();
    //            vault1.LoginAuto(ИмяХранилища, 0);

    //            var search = vault1.CreateSearch();
    //            search.FindFiles = true;
    //            search.FindFolders = false;
    //            var edmFile5 = vault1.GetObject(EdmObjectType.EdmObject_File, id);

    //            search.FileName = edmFile5.Name;
    //            var result = search.GetFirstResult();
    //            return string.Format("Path to file - {0}", result.Path);
    //        }
    //        catch (Exception)
    //        {
    //            return "Ошибка во время подключения!";
    //        }
    //    }

    //    public string PathByNameAsm(string name)
    //    {
    //        try
    //        {
    //            var vault1 = new EdmVault5();
    //            vault1.LoginAuto(ИмяХранилища, 0);

    //            var search = vault1.CreateSearch();
    //            search.FindFiles = true;
    //            search.FindFolders = false;

    //            try
    //            {
    //                search.FileName = name + ".SLDASM";
    //                var result = search.GetFirstResult();
    //                return result.Path;
    //            }
    //            catch (Exception)
    //            {
    //                search.FileName = name;
    //                var result = search.GetFirstResult();
    //                return result.Path;
    //            }
    //        }

    //        catch (Exception)
    //        {
    //            return "Ошибка во время подключения!";
    //        }

    //    }

    //    List<string> SearchAsmNames(string nameSomeLetters)
    //    {
    //        try
    //        {
    //            var vault1 = new EdmVault5();
    //            vault1.LoginAuto(ИмяХранилища, 0);

    //            var search = vault1.CreateSearch();
    //            search.FindFiles = true;
    //            search.FindFolders = false;

    //            var list = new List<string>();


    //            search.FileName = nameSomeLetters;
    //            list.Add(search.GetFirstResult().Name);


    //            while (search.GetNextResult() != null)
    //            {
    //                list.Add(search.GetNextResult().Name);
    //            }

    //            return list;
                
    //            //try
    //            //{
    //            //    search.FileName = nameSomeLetters + ".SLDASM";
    //            //    var result = search.GetFirstResult();

    //            //    return result.Name;
    //            //}
    //            //catch (Exception)
    //            //{
    //            //    search.FileName = nameSomeLetters;
    //            //    var result = search.GetFirstResult();
    //            //    return result.Path;
    //            //}
    //        }

    //        catch (Exception)
    //        {
    //            return null;
    //        }

    //    }

    //    public IEnumerable<string> AsmNames()
    //    {
    //        try
    //        {
    //            var vault1 = new EdmVault5();
    //            vault1.LoginAuto(ИмяХранилища, 0);
    //            var searchDir = new DirectoryInfo(vault1.RootFolderPath);
    //            var files = searchDir.GetFiles("*.sldasm", SearchOption.AllDirectories);
    //            return files.Select(file => Path.GetFileNameWithoutExtension(file.FullName));
    //        }
    //        catch (Exception)
    //        {
    //            return null;
    //        }
    //    }

    //    public IList<BomPartListClass.ДанныеДляВыгрузки> Bom(int type, string assemblyPath)
    //    {
    //        var bomClass = new BomPartListClass
    //        {
    //            ConnectionToSql = Constring,
    //            IdСпецификации = IdСпецификации,
    //            ИмяХранилища = ИмяХранилища,
    //            ИмяПользователя = ИмяПользователя,
    //            ПарольПользователя = Password,
    //            ПутьКСборке = assemblyPath,
    //        };

    //        switch (type)
    //        {
    //            case 0:
    //                return bomClass.BomListAsm().ToList();
    //            case 1:
    //                return bomClass.BomListPrt(true).ToList();
    //            case 2:
    //                return bomClass.BomListPrt(false).ToList();
    //            default:
    //                return bomClass.BomList();
    //        }
    //    }

    //    public IList<BomPartListClass.ДанныеДляВыгрузки> BomParts(string assemblyPath, string pdmBaseName, string userName, string password, string constring)
    //    {
    //        var bomId = 7;
    //        if (pdmBaseName == "Vents-PDM")
    //        {
    //            bomId = 8;
    //        }

    //        var bomClass = new BomPartListClass
    //        {
    //            ConnectionToSql = constring,
    //            IdСпецификации = bomId,
    //            ИмяХранилища = pdmBaseName,
    //            ИмяПользователя = userName,
    //            ПарольПользователя = password,
    //            ПутьКСборке = assemblyPath
    //        };

    //        return bomClass.BomListPrt(true).ToList();
    //    }

    //}

    #endregion

    class Program
    {
        static void Main()
        {
            
            //todo: 

            // HACK
            var baseAddress = new Uri("http://192.168.14.86/bomService");  //http://192.168.14.11:8085/bomtable   http://192.168.14.86:8085/bomService   http://srvkb:8085/bomtable
            Console.Title = "BomService.v1";

            // Create the ServiceHost.
            using (var host = new ServiceHost(typeof(BomServiceClass.BomTableService), baseAddress))
            {
               // var httpBinding = new HttpTransportBindingElement
               // {
               //     MaxBufferSize = 2147483647,
               //     MaxReceivedMessageSize = 2147483647
               // };
               // var binding = new CustomBinding();
               //// binding.Elements.Add(new TextMessageEncodingBindingElement(MessageVersion.Soap11, System.Text.Encoding.UTF8));
               // binding.Elements.Add(httpBinding);


              //   Enable metadata publishing.
                var smb = new ServiceMetadataBehavior
                {
                    HttpGetEnabled = true,
                    MetadataExporter = { PolicyVersion = PolicyVersion.Policy15 },

                    //HttpGetBinding = new BasicHttpBinding
                    //{
                    //    MaxBufferSize = 2147483647,
                    //    MaxReceivedMessageSize = 2147483647
                    //}
                };
                host.Description.Behaviors.Add(smb);

                // Open the ServiceHost to start listening for messages. Since
                // no endpoints are explicitly configured, the runtime will create
                // one endpoint per base address for each service contract implemented
                // by the service.
                host.Open();

                Console.WriteLine("The service is ready at {0}", baseAddress);
                Console.WriteLine("Press <Enter> to stop the service.");
                Console.ReadLine();

                // Close the ServiceHost.
                host.Close();
            }
        }
    }
}
