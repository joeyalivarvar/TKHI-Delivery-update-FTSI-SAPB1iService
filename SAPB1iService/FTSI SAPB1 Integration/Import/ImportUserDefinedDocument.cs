using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using SAPbobsCOM;
using System.Windows;
using System.Windows.Forms;
using System.Data;
using System.Text.RegularExpressions;
using System.Xml;
using System.Data.SqlClient;
using Microsoft.VisualBasic.FileIO;
using Renci.SshNet.Common;

namespace SAPB1iService
{
    class ImportUserDefinedDocument
    {
        private static DateTime dteStart;
        private static string strTransType = "Documents - Update Delivery Date";
        private static string strMsgBod;
        private static string errorType;

        private static DataTable oDTTrackNo, oDTFreight, oDTXML;

        private static DataTable oDataTable, oDataTable2;

        public static void _ImportUserDefinedDocument()
        {
            importFromFile();
        }
        private static void importFromFile()
        {
            string strStatus = "";

            try
            {
                string[] strFileImport = new string[] { string.Format("*.xlsx") };

                foreach (string fileimport in strFileImport)
                {
                    foreach (var strFile in Directory.GetFiles(GlobalVariable.strImpPath, fileimport))
                    {

                        GlobalVariable.strFileName = Path.GetFileName(strFile);
                        errorType = "";
                        dteStart = DateTime.Now;

                        string strExtension = Path.GetExtension(GlobalVariable.strFileName);
                        if (Path.GetExtension(GlobalVariable.strFileName) == ".xlsx")
                        {
                            if (importDIAPIPostDocumentFExcel(strFile))
                                strStatus = "S";
                            else
                            {
                                if (errorType == "Invalid File Name")
                                    strStatus = "I";
                                else
                                    strStatus = "E";
                            }
                        }

                        if (strStatus == "S" || strStatus == "E")
                        {
                            TransferFile.transferProcFiles("Import", strStatus, Path.GetFileName(strFile));
                            GlobalFunction.sendAlert(strStatus, "Import", strMsgBod, GlobalVariable.oObjectType, GlobalVariable.oCompany.GetNewObjectKey().ToString());

                        }

                    }
                }

                GC.Collect();
            }
            catch (Exception ex)
            {
                SystemFunction.transHandler("Import", strTransType, "28", "", "", "", dteStart, "E", "-111", ex.Message.ToString());
            }
        }

        private static bool importDIAPIPostDocumentFExcel(string strFile)
        {
            string strQuery = null;
            SAPbobsCOM.Recordset oRecordset;
            try
            {
                if (strFile.Contains("POD"))
                {
                    strTransType = "Documents - Update Delivery Date";

                    if (GlobalFunction.importXLSX(Path.GetFullPath(strFile), "YES", "Sheet2"))

                        if (!(importPOD()))
                            return false;
                        else
                            return true;
                    else
                        return false;
                }
                else
                {
                    errorType = "Invalid File Name";
                    strMsgBod = string.Format("File Name is Not Valid! Error Posting {0}.\rError Code: {1}\rDescription: {2} ", GlobalVariable.strFileName, GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                    strQuery = string.Format("SELECT \"U_FileName\" FROM \"@FTPISL\" WHERE CAST(\"U_FileName\" AS NVARCHAR) = '{0}'", GlobalVariable.strFileName.ToString());
                    //strQuery = string.Format("SELECT TO_NVARCHAR(\"U_FileName\", TEXT) FROM \"@FTPISL\" WHERE \"U_FileName\" = '{0}'", GlobalVariable.strFileName);

                    oRecordset = null;
                    oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                    oRecordset.DoQuery(strQuery);
                    if (oRecordset.RecordCount == 0)
                        SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), strMsgBod);

                    GC.Collect();
                    return false;
                }

                GC.Collect();

                return true;
            }
            catch (Exception ex)
            {

                GlobalVariable.intErrNum = -111;
                GlobalVariable.strErrMsg = ex.Message.ToString();

                SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), "", "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                GC.Collect();

                return false;
            }
        }
        private static bool importPOD()
        {

            string strTrckNo, strCntrDte = "", strActlDte = "", strQuery = "", strStatus = "",
                   strPrinted = "", strDocEntry, strDocNum, strCancel, strDocDate, strCurDat;

            int intDay;

            int intCtr = 0;

            bool blWithErr = false;

            DataTable oDTDate;
            DataRow[] oDRTrck;

            DateTime dteDoc, dteAct, dteCnt, dteCur;

            SAPbobsCOM.Recordset oRecordset;

            try
            {
                initTrackDT();

                strCurDat = DateTime.Today.ToString("MM/dd/yyyy");
                dteCur = DateTime.ParseExact(strCurDat, "MM/dd/yyyy",
                                      System.Globalization.CultureInfo.InvariantCulture);
                //= Glob/*a*/lFunction.getDateTime(.Replace("/", ""), "MMDDYYYY", "DD/MM/YYYY");

                GlobalFunction.getObjType(13);

                oDTDate = GlobalVariable.oDTImpData.DefaultView.ToTable(true, GlobalVariable.oDTImpData.Columns[2].ColumnName,
                                                                              GlobalVariable.oDTImpData.Columns[5].ColumnName,
                                                                              GlobalVariable.oDTImpData.Columns[7].ColumnName);

                for (int intRow = 0; intRow <= oDTDate.Rows.Count - 1; intRow++)
                {
                    strTrckNo = oDTDate.Rows[intRow][0].ToString();

                    strQuery = string.Format(string.Format("{0} = '{1}' ", GlobalVariable.oDTImpData.Columns[2].ColumnName, strTrckNo));
                    oDRTrck = GlobalVariable.oDTImpData.Select(strQuery);

                    if (oDRTrck.Length > 1)
                    {
                        blWithErr = true;

                        GlobalVariable.intErrNum = -999;
                        GlobalVariable.strErrMsg = string.Format("Error updating Delivery Date / Counter Date with Tracking No {0}.  Duplicate Tracking No.", strTrckNo);

                        SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                        continue;
                    }

                    strQuery = string.Format("SELECT \"DocEntry\", \"DocDate\", \"DocNum\", \"CANCELED\", \"DocStatus\", \"Printed\" FROM OINV WHERE \"TrackNo\" = '{0}' ", strTrckNo);

                    oRecordset = null;
                    oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                    oRecordset.DoQuery(strQuery);

                    if (oRecordset.RecordCount > 0)
                    {
                        strDocEntry = oRecordset.Fields.Item("DocEntry").Value.ToString();
                        strDocNum = oRecordset.Fields.Item("DocNum").Value.ToString();
                        strStatus = oRecordset.Fields.Item("DocStatus").Value.ToString();
                        strPrinted = oRecordset.Fields.Item("Printed").Value.ToString();
                        strCancel = oRecordset.Fields.Item("CANCELED").Value.ToString();
                        strDocDate = oRecordset.Fields.Item("DocDate").Value.ToString("MM/dd/yyyy");

                        //dteDoc = GlobalFunction.getDateTime(strDocDate.Replace("/", ""), "MMDDYYYY", "DD/MM/YYYY");

                        dteDoc = DateTime.ParseExact(strDocDate, "MM/dd/yyyy",
                                     System.Globalization.CultureInfo.InvariantCulture);

                        SystemFunction.errorAppend(dteDoc.ToString());

                        if (strStatus != "C" && strCancel == "N")
                        {
                            if (!(string.IsNullOrEmpty(oDTDate.Rows[intRow][1].ToString())))
                            {
                                strActlDte = Convert.ToDateTime(oDTDate.Rows[intRow][1].ToString()).ToString("MM/dd/yyyy");

                                //dteAct = GlobalFunction.getDateTime(strActlDte.Replace("/", ""), "MMDDYYYY", "DD/MM/YYYY");
                                dteAct = DateTime.ParseExact(strActlDte, "MM/dd/yyyy",
                                      System.Globalization.CultureInfo.InvariantCulture);

                                SystemFunction.errorAppend(dteAct.ToString());

                                if (dteAct < dteDoc || dteAct > dteCur)
                                {
                                    blWithErr = true;

                                    GlobalVariable.intErrNum = -999;
                                    GlobalVariable.strErrMsg = string.Format("Error updating Delivery Date with Tracking No {0}.  Please Check Actual Delivery Date.", strTrckNo);

                                    SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                                    continue;
                                }
                            }
                            else
                            {
                                GlobalVariable.intErrNum = -999;
                                GlobalVariable.strErrMsg = string.Format("Error updating Tracking No {0}.  Please Check Actual Delivery Date.", strTrckNo);

                                SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                                continue;
                            }

                            if (!(string.IsNullOrEmpty(oDTDate.Rows[intRow][2].ToString())))
                                strCntrDte = Convert.ToDateTime(oDTDate.Rows[intRow][2].ToString()).ToString("MM/dd/yyyy");
                            else
                                strCntrDte = Convert.ToDateTime(oDTDate.Rows[intRow][1].ToString()).ToString("MM/dd/yyyy");

                            //dteCnt = GlobalFunction.getDateTime(strCntrDte.Replace("/", ""), "MMDDYYYY", "DD/MM/YYYY");

                            dteCnt = DateTime.ParseExact(strCntrDte, "MM/dd/yyyy",
                                     System.Globalization.CultureInfo.InvariantCulture);

                            if (dteCnt < dteDoc || dteCnt > dteCur)
                            {
                                blWithErr = true;

                                GlobalVariable.intErrNum = -999;
                                GlobalVariable.strErrMsg = string.Format("Error updating Counter Date with Tracking No {0}.  Please Check Counter Date.", strTrckNo);

                                SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                                continue;
                            }

                            strQuery = string.Format("UPDATE OINV SET \"U_ActDelDate\" = to_date('{0}', 'MM/dd/yyyy'), \"U_CounterDate\" = to_date('{1}', 'MM/dd/yyyy') WHERE \"TrackNo\" = '{2}' ", strActlDte, strCntrDte, strTrckNo);

                            if (!(string.IsNullOrEmpty(strQuery)))
                            {
                                if (!(SystemFunction.executeQuery(strQuery)))
                                {
                                    blWithErr = true;

                                    GlobalVariable.intErrNum = -999;
                                    GlobalVariable.strErrMsg = string.Format("Error updating Delivery Date / Counter Date with Tracking No {0}.", strTrckNo);

                                    SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);
                                }
                                else
                                {
                                    oDTTrackNo.Rows.Add(strTrckNo, strDocEntry, strDocNum);
                                }
                            }
                        }
                        else
                        {
                            blWithErr = true;

                            GlobalVariable.intErrNum = -999;
                            GlobalVariable.strErrMsg = string.Format("Tracking No {0} Status is not Open.", strTrckNo);

                            SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                            continue;

                        }
                    }
                    else
                    {
                        blWithErr = true;

                        GlobalVariable.intErrNum = -999;
                        GlobalVariable.strErrMsg = string.Format("Tracking No {0} not exist in SAP AR Invoices.", strTrckNo);

                        SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                        continue;
                    }
                }

                if (blWithErr == false)
                {
                    for (int intRow = 0; intRow <= oDTTrackNo.Rows.Count - 1; intRow++)
                    {
                        strTrckNo = oDTTrackNo.Rows[intRow]["TrackNo"].ToString();
                        strDocEntry = oDTTrackNo.Rows[intRow]["DocEntry"].ToString();
                        strDocNum = oDTTrackNo.Rows[intRow]["DocNum"].ToString();

                        strMsgBod = string.Format("Successfully updated delivery date / counter date from {0} with Tracking No {1} and Invoice No {2}. ", GlobalVariable.strFileName, strTrckNo, strDocNum);

                        SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, strDocEntry, strDocNum, dteStart, "S", "0", strMsgBod);

                        GlobalFunction.sendAlert("S", "Import", strMsgBod, GlobalVariable.oObjectType, strDocEntry);
                    }

                    return true;
                }
                else
                {
                    strMsgBod = string.Format("Error updating delivery date / counter date from {0}. Please check error logs for information/remarks.", GlobalVariable.strFileName);

                    GlobalFunction.sendAlert("E", "Import", strMsgBod, GlobalVariable.oObjectType, "");

                    return false;
                }

                GC.Collect();

            }
            catch (Exception ex)
            {

                GlobalVariable.intErrNum = -111;
                GlobalVariable.strErrMsg = ex.Message.ToString();

                SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                GC.Collect();

                return false;
            }
        }
        private static void initTrackDT()
        {
            oDTTrackNo = new DataTable("UpdateTrackNo");
            oDTTrackNo.Columns.Add("TrackNo", typeof(System.String));
            oDTTrackNo.Columns.Add("DocEntry", typeof(System.String));
            oDTTrackNo.Columns.Add("DocNum", typeof(System.String));
        }

    }
}
