using SolidWorks.Interop.swdocumentmgr;

namespace AddinConvertTo.Classes
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
                    swDocument13.CloseDoc();
                }
            }
            catch (System.Exception ex)
            {
                Logger.ToLog($"ERROR SHEET METAL: {ex.ToString()}, {ex.StackTrace}", 10001);
            }
            return result;
        }
        ~SheetMetal() { }
    }
}