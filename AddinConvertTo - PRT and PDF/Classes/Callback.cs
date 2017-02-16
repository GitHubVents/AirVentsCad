
using EPDM.Interop.epdm;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddinConvertTo.Classes
{
    public class Callback
    {
        #region IEdmUnlockOpCallback Members
        public EdmUnlockOpReply MsgBox(EdmUnlockOpMsg mssge, int docId, int projID, string path, ref EdmUnlockErrInfo err)
        {
            return EdmUnlockOpReply.Euor_OK;
        }
        public void ProgressBegin(EdmProgressType type, EdmUnlockEvent events, int steps)
        {
            return;
        }
        public void ProgressEnd(EdmProgressType type)
        {
            //Demonstrates callback
            Logger.ToLog("ProgressEnd called. " + type, 1000);
            //Logger.ToLog("ProgressEnd called.", 1000);
            return;
        }
        public bool ProgressStep(EdmProgressType type, string msgText, int progressPos)
        {
            return true;
        }
        public bool ProgressStepEvent(EdmProgressType type, EdmUnlockEventMsg opText, int progressPos)
        {
            return true;
        }
        #endregion
    }
}