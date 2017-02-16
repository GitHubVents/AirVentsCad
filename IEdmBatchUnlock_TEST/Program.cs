using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EPDM.Interop.epdm;

namespace IEdmBatchUnlock_TEST
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                EdmVault5 edmVault5 = new EdmVault5();
                if (!edmVault5.IsLoggedIn)
                {
                    edmVault5.LoginAuto("Vents-PDM", 0);
                }


                IEdmFolder5 folder = edmVault5.GetFolderFromPath(@"D:\Vents-PDM\Test");
                folder.AddFile(0, @"C:\Users\Antonyk\Desktop\МенюКафе.xls");

                //VentsCadLibrary.SwEpdm.CheckInOutPdm(@"D:\Vents-PDM\Test\МенюКафе.xls", false);


                IEdmFile5 edmFile5 = edmVault5.GetFileFromPath(@"D:\Vents-PDM\Test\МенюКафе.xls", out folder);
                //  edmFile5.GetFileCopy(0, 0, folder.ID, (int)EdmGetFlag.EdmGet_Simple);
                // Разрегистрировать

                Console.Write(edmFile5.Name);
                VentsCadLibrary.SwEpdm.CheckInOutPdm(@"D:\Vents-PDM\Test\МенюКафе.xls",true);
                //try
                //{


                //    if (!edmFile5.IsLocked)
                //    {
                //        Console.WriteLine("no locked");
                //        edmFile5.LockFile(folder.ID, 0);
                //    }
                //}
                //catch (Exception ex)
                //{
                //    Console.WriteLine("no Locked: " + ex);
                //}

                //try
                //{
                //    if (edmFile5.IsLocked)
                //    {
                //        Console.WriteLine("Locked");
                //        edmFile5.UnlockFile(0, "");

                //    }
                //}

                //catch (Exception ex)
                //{
                //    Console.WriteLine("Locked: " + ex);
                //}
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }


    }





    //IEdmBatchUnlock2 batchUnlocker;
    //IEdmSelectionList6 fileList = null;
    //EdmSelectionObject poSel;
    //EdmSelItem[] ppoSelection = new EdmSelItem[11];
    //int fileCount = 0;
    //IEdmFile5 aFile;
    //IEdmFolder5 aFolder;
    //IEdmFolder5 ppoRetParentFolder;
    //IEdmPos5 aPos;
    //string str;

    //bool retVal;      
    //batchUnlocker = (IEdmBatchUnlock2)edmVault5.CreateUtility(EdmUtility.EdmUtil_BatchUnlock);
    //int i = 0;
    //string[] filenames = { @"D:\Vents-PDM\Проекты\Blauberg\02-01-Panels\02-01-01-345-567-50-AZ.SLDPRT", @"D:\Vents-PDM\Проекты\Blauberg\02-01-Panels\02-01-02-345-567-50-AZ.SLDPRT" };
    //string [] paths = { @"D:\Vents-PDM\Проекты\Blauberg\02-01-Panels", @"D:\Vents-PDM\Проекты\Blauberg\02-01-Panels" };
    //IEdmBatchAdd poAdder = (IEdmBatchAdd)edmVault5.CreateUtility(EdmUtility.EdmUtil_BatchAdd);

    //for (int ij = 0; ij < 2; ij++)
    //{


    //    poAdder.AddFileFromPathToPath(filenames[ij], paths[ij], 0, "", 0);
    //    aFile = edmVault5.GetFileFromPath(filenames[ij], out ppoRetParentFolder);
    //    aPos = aFile.GetFirstFolderPosition();
    //    aFolder = aFile.GetNextFolder(aPos);
    //    ppoSelection[i] = new EdmSelItem();
    //    ppoSelection[i].mlDocID = aFile.ID;
    //    ppoSelection[i].mlProjID = aFolder.ID;
    //    i = i + 1;
    //}
    //// Add selections to the batch of files to check in
    //batchUnlocker.AddSelection(edmVault5, ref ppoSelection);

    //if ((batchUnlocker != null))
    //{
    //    batchUnlocker.CreateTree(0, (int)EdmUnlockBuildTreeFlags.Eubtf_ShowCloseAfterCheckinOption + (int)EdmUnlockBuildTreeFlags.Eubtf_MayUnlock);
    //    batchUnlocker.Comment = "Updates";
    //    fileList = (IEdmSelectionList6)batchUnlocker.GetFileList((int)EdmUnlockFileListFlag.Euflf_GetUnlocked + (int)EdmUnlockFileListFlag.Euflf_GetUndoLocked + (int)EdmUnlockFileListFlag.Euflf_GetUnprocessed);
    //    aPos = fileList.GetHeadPosition();
    //    batchUnlocker.UnlockFiles(0, null);
    //    IEdmBatchAdd2 statuses = batchUnlocker.GetStatus((int)EdmUnlockStatusFlag.Eusf_CloseAfterCheckinFlag);
    //    Console.ReadLine();
    //}

}
 