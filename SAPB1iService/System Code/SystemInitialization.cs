using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;
using SAPB1iService;
using System.IO;

namespace SAPB1iService
{
    class SystemInitialization
    {
        public static bool initTables()
        {

            /******************************* TOUCH ME NOT PLEASE *****************************************************/

            if (SystemFunction.createUDT("FTPISL", "FT Payroll Integration Log", SAPbobsCOM.BoUTBTableType.bott_NoObject) == false)
                return false;

            if (SystemFunction.createUDT("FTISSP", "FT Integration SetUp", SAPbobsCOM.BoUTBTableType.bott_NoObject) == false)
                return false;

            /****************************** UNTIL HERE - THANK YOU ***************************************************/

            return true;
        }
        public static bool initFields()
        {

            /******************************* TOUCH ME NOT PLEASE *****************************************************/

            #region "FRAMEWORK UDF"

            /******************************* INTEGRATION SERVICE LOG ***********************************************/

            if (SystemFunction.isUDFexists("@FTPISL", "Process") == false)
                if (SystemFunction.createUDF("@FTPISL", "Process", "Process", SAPbobsCOM.BoFieldTypes.db_Alpha, 250, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTPISL", "TransType") == false)
                if (SystemFunction.createUDF("@FTPISL", "TransType", "Transaction Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 250, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTPISL", "ObjType") == false)
                if (SystemFunction.createUDF("@FTPISL", "ObjType", "Object Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTPISL", "TransDate") == false)
                if (SystemFunction.createUDF("@FTPISL", "TransDate", "Transaction Date", SAPbobsCOM.BoFieldTypes.db_Date, 0, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTPISL", "FileName") == false)
                if (SystemFunction.createUDF("@FTPISL", "FileName", "FileName", SAPbobsCOM.BoFieldTypes.db_Memo, 0, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTPISL", "TrgtDocKey") == false)
                if (SystemFunction.createUDF("@FTPISL", "TrgtDocKey", "Base Document Key", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTPISL", "TrgtDocNum") == false)
                if (SystemFunction.createUDF("@FTPISL", "TrgtDocNum", "Base Document No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTPISL", "StartTime") == false)
                if (SystemFunction.createUDF("@FTPISL", "StartTime", "StartTime", SAPbobsCOM.BoFldSubTypes.st_Time, 0, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTPISL", "EndTime") == false)
                if (SystemFunction.createUDF("@FTPISL", "EndTime", "EndTime", SAPbobsCOM.BoFldSubTypes.st_Time, 0, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTPISL", "Status") == false)
                if (SystemFunction.createUDF("@FTPISL", "Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTPISL", "ErrorCode") == false)
                if (SystemFunction.createUDF("@FTPISL", "ErrorCode", "Error Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTPISL", "Remarks") == false)
                if (SystemFunction.createUDF("@FTPISL", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, 0, "", "", "") == false)
                    return false;

            /******************************* INTEGRATION SETUP ***********************************************/

            if (SystemFunction.isUDFexists("@FTISSP", "ExportFile") == false)
                if (SystemFunction.createUDF("@FTISSP", "ExportFile", "Export File Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTISSP", "ExportPath") == false)
                if (SystemFunction.createUDF("@FTISSP", "ExportPath", "Export Path", SAPbobsCOM.BoFieldTypes.db_Alpha, 200, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTISSP", "ImportFile") == false)
                if (SystemFunction.createUDF("@FTISSP", "ImportFile", "Import File Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTISSP", "ImportPath") == false)
                if (SystemFunction.createUDF("@FTISSP", "ImportPath", "Import Path", SAPbobsCOM.BoFieldTypes.db_Alpha, 200, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTISSP", "Delimiter") == false)
                if (SystemFunction.createUDF("@FTISSP", "Delimiter", "Delimiter", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTISSP", "ProcessTime") == false)
                if (SystemFunction.createUDF("@FTISSP", "ProcessTime", "Process Time", SAPbobsCOM.BoFldSubTypes.st_Time, 0, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTISSP", "AlwaysRun") == false)
                if (SystemFunction.createUDF("@FTISSP", "AlwaysRun", "Services Always Running?", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "N", "N - No, Y - Yes", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTISSP", "ProcSer") == false)
                if (SystemFunction.createUDF("@FTISSP", "ProcSer", "Process Service", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "N", "N - No, Y - Yes", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTISSP", "RunRep") == false)
                if (SystemFunction.createUDF("@FTISSP", "RunRep", "Reprocess Error File?", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "N", "N - No, Y - Yes", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("@FTISSP", "RepDate") == false)
                if (SystemFunction.createUDF("@FTISSP", "RepDate", "Reprocess Error Date", SAPbobsCOM.BoFieldTypes.db_Date, 0, "", "", "") == false)
                    return false;

            /************************** MARKETING DOCUMENTS ****************************************************************/

            if (SystemFunction.isUDFexists("OINV", "isExtract") == false)
                if (SystemFunction.createUDF("OINV", "isExtract", "Extracted", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "N", "N - No, E - Error, Y -Yes", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("OINV", "FileName") == false)
                if (SystemFunction.createUDF("OINV", "FileName", "FileName", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("OINV", "RefNum") == false)
                if (SystemFunction.createUDF("OINV", "RefNum", "Reference Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("INV1", "RefNum") == false)
                if (SystemFunction.createUDF("INV1", "RefNum", "Reference Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("INV1", "BaseLine") == false)
                if (SystemFunction.createUDF("INV1", "BaseLine", "Base Line", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("INV1", "BaseRef") == false)
                if (SystemFunction.createUDF("INV1", "BaseRef", "Base Reference", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("INV1", "BaseType") == false)
                if (SystemFunction.createUDF("INV1", "BaseType", "Base Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("INV3", "RefNum") == false)
                if (SystemFunction.createUDF("INV3", "RefNum", "Reference Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("INV3", "BaseLine") == false)
                if (SystemFunction.createUDF("INV3", "BaseLine", "Base Line", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("INV3", "BaseRef") == false)
                if (SystemFunction.createUDF("INV3", "BaseRef", "Base Reference", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("INV3", "BaseType") == false)
                if (SystemFunction.createUDF("INV3", "BaseType", "Base Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("INV5", "RefNum") == false)
                if (SystemFunction.createUDF("INV5", "RefNum", "Reference Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", "", "") == false)
                    return false;

            /************************** ITEM MASTER DATA ***************************************************************/

            if (SystemFunction.isUDFexists("OITM", "isExtract") == false)
                if (SystemFunction.createUDF("OITM", "isExtract", "Extracted", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "N", "N - No, E - Error, Y -Yes", "") == false)
                    return false;

            /************************** BUSINESS PARTNER DATA **********************************************************/

            if (SystemFunction.isUDFexists("OCRD", "isExtract") == false)
                if (SystemFunction.createUDF("OCRD", "isExtract", "Extracted", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "N", "N - No, E - Error, Y -Yes", "") == false)
                    return false;

            /************************** ADMINISTRATION ****************************************************************/

            if (SystemFunction.isUDFexists("OUSR", "IntMsg") == false)
                if (SystemFunction.createUDF("OUSR", "IntMsg", "Integration Message", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "N", "N - No, Y - Yes", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("OADM", "Company") == false)
                if (SystemFunction.createUDF("OADM", "Company", "Company", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "", "", "") == false)
                    return false;

            #endregion

            /****************************** UNTIL HERE - THANK YOU ***************************************************/

            if (SystemFunction.isUDFexists("OINV", "ActDelDate") == false)
                if (SystemFunction.createUDF("OINV", "ActDelDate", "Actual Delivery Date", SAPbobsCOM.BoFieldTypes.db_Date, 0, "", "", "") == false)
                    return false;

            if (SystemFunction.isUDFexists("OINV", "CounterDate") == false)
                if (SystemFunction.createUDF("OINV", "CounterDate", "Counter Date", SAPbobsCOM.BoFieldTypes.db_Date, 0, "", "", "") == false)
                    return false;

            return true;
        }
        public static bool initUDO()
        {

            return true;
        }
        public static bool initFolders()
        {
            try
            {
                string strDate = DateTime.Today.ToString("MMddyyyy") + @"\";

                string strExp = @"Export\" + strDate;
                string strImp = @"Import\" + strDate;

                GlobalVariable.strErrLogPath = GlobalVariable.strFilePath + @"\Error Log";
                if (!Directory.Exists(GlobalVariable.strErrLogPath))
                    Directory.CreateDirectory(GlobalVariable.strErrLogPath);

                GlobalVariable.strSQLScriptPath = GlobalVariable.strFilePath + @"\SQL Scripts\";
                if (!Directory.Exists(GlobalVariable.strSQLScriptPath))
                    Directory.CreateDirectory(GlobalVariable.strSQLScriptPath);

                GlobalVariable.strSAPScriptPath = GlobalVariable.strFilePath + @"\SAP Scripts\";
                if (!Directory.Exists(GlobalVariable.strSAPScriptPath))
                    Directory.CreateDirectory(GlobalVariable.strSAPScriptPath);

                GlobalVariable.strExpSucPath = GlobalVariable.strFilePath + @"\Success Files\" + strExp;
                if (!Directory.Exists(GlobalVariable.strExpSucPath))
                    Directory.CreateDirectory(GlobalVariable.strExpSucPath);

                GlobalVariable.strExpErrPath = GlobalVariable.strFilePath + @"\Error Files\" + strExp;
                if (!Directory.Exists(GlobalVariable.strExpErrPath))
                    Directory.CreateDirectory(GlobalVariable.strExpErrPath);

                GlobalVariable.strImpSucPath = GlobalVariable.strFilePath + @"\Success Files\" + strImp;
                if (!Directory.Exists(GlobalVariable.strImpSucPath))
                    Directory.CreateDirectory(GlobalVariable.strImpSucPath);

                GlobalVariable.strImpErrPath = GlobalVariable.strFilePath + @"\Error Files\" + strImp;
                if (!Directory.Exists(GlobalVariable.strImpErrPath))
                    Directory.CreateDirectory(GlobalVariable.strImpErrPath);

                GlobalVariable.strImpPath = GlobalVariable.strFilePath + @"\Import Files\";
                if (!Directory.Exists(GlobalVariable.strImpPath))
                    Directory.CreateDirectory(GlobalVariable.strImpPath);

                GlobalVariable.strExpPath = GlobalVariable.strFilePath + @"\Export Files\";
                if (!Directory.Exists(GlobalVariable.strExpPath))
                    Directory.CreateDirectory(GlobalVariable.strExpPath);

                GlobalVariable.strConPath = GlobalVariable.strFilePath + @"\Connection Path\";
                if (!Directory.Exists(GlobalVariable.strConPath))
                    Directory.CreateDirectory(GlobalVariable.strConPath);

                GlobalVariable.strTempPath = GlobalVariable.strFilePath + @"\Temp Files\";
                if (!Directory.Exists(GlobalVariable.strTempPath))
                    Directory.CreateDirectory(GlobalVariable.strTempPath);

                GlobalVariable.strAttImpPath = GlobalVariable.strFilePath + @"\Attachment\" + strImp;
                if (!Directory.Exists(GlobalVariable.strAttImpPath))
                    Directory.CreateDirectory(GlobalVariable.strAttImpPath);

                GlobalVariable.strAttExpPath = GlobalVariable.strFilePath + @"\Attachment\" + strExp;
                if (!Directory.Exists(GlobalVariable.strAttExpPath))
                    Directory.CreateDirectory(GlobalVariable.strAttExpPath);

                GlobalVariable.strArcExpPath = GlobalVariable.strFilePath + @"\Archive Files\Export\";
                if (!Directory.Exists(GlobalVariable.strArcExpPath))
                    Directory.CreateDirectory(GlobalVariable.strArcExpPath);

                GlobalVariable.strArcImpPath = GlobalVariable.strFilePath + @"\Archive Files\Import\";
                if (!Directory.Exists(GlobalVariable.strArcImpPath))
                    Directory.CreateDirectory(GlobalVariable.strArcImpPath);

                //Import._Import();
                return true;
            }
            catch(Exception ex)
            {
                SystemFunction.errorAppend(string.Format("Error initializing program directory. {0}", ex.Message.ToString()));
                return false;
            }
        }
        public static bool initStoreProcedure()
        {
            if (!(SystemFunction.initStoredProcedures(GlobalVariable.strSAPScriptPath)))
                return false;

            return true;
        }
        public static bool initSQLConnection()
        {
            if (File.Exists(GlobalVariable.strSQLSettings))
            {
                if (SystemFunction.connectSQL(GlobalVariable.strSQLSettings))
                    return true;
                else
                    return false;
            }
            else
                return false;
        }

    }
}
