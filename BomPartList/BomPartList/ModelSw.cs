using System;
using System.Runtime.InteropServices;
using System.Threading;
using EPDM.Interop.epdm;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace BomPartList
{
    /// <summary>
    /// 
    /// </summary>
    public partial class BomPartListClass
    {
        #region Поля 
        
        SldWorks _swApp;
     
        #endregion

        #region Методы 

        #region Registration
      
        void GetLastVersionPdm(string path, string pdmBase)
        {
            try
            {
                //LoggerInfo(string.Format("Получение последней версии по пути {0}\nБаза - {1}", path, pdmBase));
                var vault1 = new EdmVault5();
                IEdmFolder5 oFolder;
                vault1.LoginAuto(pdmBase, 0);

                var edmFile5 = vault1.GetFileFromPath(path, out oFolder);
                edmFile5.GetFileCopy(1, 0, oFolder.ID, (int)EdmGetFlag.EdmGet_Simple);
            }
            catch (Exception exception)
            {
                //LoggerError(string.Format("Во время получения последней версии по пути {0} возникла ошибка{2}\nБаза - {1}", path, pdmBase, exception.ToString()));
                throw Exception = new ArgumentException(string.Format("Во время получения последней версии по пути {0} возникла ошибка{2}\nБаза - {1}", path, pdmBase, exception.ToString()));
                //Logger.Log(LogLevel.Error, string.Format("Во время получения последней версии по пути {0} возникла ошибка\nБаза - {1}", path, pdmBase), exception);
            }
        }

        void CheckInOutPdm(string filePath, bool registration, string pdmBase)
        {
            var retryCount = 2;
            var success = false;
            while (!success && retryCount > 0)
            {
                try
                {
                    var vault1 = new EdmVault5();
                    IEdmFolder5 oFolder;
                    vault1.LoginAuto(pdmBase, 0);
                    var edmFile5 = vault1.GetFileFromPath(filePath, out oFolder);

                    // Разрегистрировать
                    if (registration == false)
                    {
                        //  edmFile5.GetFileCopy(1, 0, oFolder.ID, (int)EdmGetFlag.EdmGet_Simple);
                        // edmFile5.GetFileCopy(0, 0, oFolder.ID, (int)EdmGetFlag.EdmGet_Simple);
                        edmFile5.LockFile(oFolder.ID, 0);
                    }

                    // Зарегистрировать
                    if (registration)
                    {
                        //edmFile5.UnlockFile(8, "", (int)EdmUnlockFlag.EdmUnlock_Simple, null);
                        edmFile5.UnlockFile(oFolder.ID, "");
                        Thread.Sleep(50);
                    }

                    //Logger.Log(LogLevel.Debug, string.Format("В базе PDM - {1}, зарегестрирован документ по пути {0}", filePath, pdmBase));

                    success = true;
                }
                catch (Exception exception)
                {
                    retryCount--;
                    var ex = exception;
                    Thread.Sleep(200);
                    if (retryCount == 0)
                    {
                        // throw; //or handle error and break/return
                    }
                    throw Exception = new ArgumentException(ex.ToString());
                }
            }
            if (!success)
            {
             //   Logger.Log(LogLevel.Error, string.Format("Во время регистрации документа по пути {0} возникла ошибка\nБаза - {1}", filePath, pdmBase), ex);
            }
        }

        #endregion

        #region AdditionalMethods
  
        bool IsSheetMetalPart(string partPath, string pdmBase)
        {
            InitializeSw(true);
            IModelDoc2 swDoc = null;

            try
            {
                GetLastVersionPdm(partPath, pdmBase);
                swDoc = _swApp.OpenDoc6(partPath, (int)swDocumentTypes_e.swDocPART,
                    (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", 0, 0);
                return IsSheetMetalPart((IPartDoc)swDoc);
            }
            catch (Exception)
            {
                throw Exception;
            }
            finally
            {
                if (swDoc != null) _swApp.CloseDoc(swDoc.GetTitle());
                _swApp.ExitApp();
                _swApp = null;
            }
        }

        static bool IsSheetMetalPart(IPartDoc swPart)
        {
            var isSheetMetal = false;
            var vBodies = swPart.GetBodies2((int)swBodyType_e.swSolidBody, false);

            foreach (Body2 vBody in vBodies)
            {
                isSheetMetal = vBody.IsSheetMetal();
            }
            return isSheetMetal;
        }

        internal bool InitializeSw(bool visible)
        {
            try
            {
                _swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
            }
            catch (Exception)
            {
                _swApp = new SldWorks { Visible = visible };
            }
            return _swApp != null;
        }

        #endregion
        
        #endregion

    }
}
