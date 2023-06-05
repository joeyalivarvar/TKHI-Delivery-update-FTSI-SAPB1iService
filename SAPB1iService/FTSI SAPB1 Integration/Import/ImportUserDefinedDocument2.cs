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

namespace SAPB1iService
{
    class ImportUserDefinedDocument
    {
        private static DateTime dteStart;
        private static string strTransType = "Documents - Update Delivery Date";
        private static string strMsgBod;

        private static DataTable oDTTrackNo, oDTFreight, oDTXML;

        public static void _ImportUserDefinedDocument()
        {
            importFromFile();
        }
        private static void importFromFile()
        {
            string strStatus = "";

            try
            {
                string[] strFileImport = new string[] { string.Format("*.xlsx"), string.Format("*.xml"), string.Format("*.csv"), };

                foreach (string fileimport in strFileImport)
                {
                    foreach (var strFile in Directory.GetFiles(GlobalVariable.strImpPath, fileimport))
                    {
                        GlobalVariable.strFileName = Path.GetFileName(strFile);

                        dteStart = DateTime.Now;

                        if (fileimport == "*.xlsx")
                        {
                            if (importDIAPIPostDocumentFExcel(strFile))
                                strStatus = "S";
                            else
                                strStatus = "E";
                        }
                        else if (fileimport == "*.csv")
                        {
                            if (importDIAPIPostDocumentFCSV(strFile))
                                strStatus = "S";
                            else
                                strStatus = "E";
                        }
                        else if (fileimport == "*.xml")
                        {
                            if (importDIAPIPostDocumentFXML(strFile))
                                strStatus = "S";
                            else
                                strStatus = "E";
                        }

                        TransferFile.transferProcFiles("Import", strStatus, Path.GetFileName(strFile));

                        GlobalFunction.sendAlert(strStatus, "Import", strMsgBod, GlobalVariable.oObjectType, GlobalVariable.oCompany.GetNewObjectKey().ToString());

                        //EmailSender._EmailSender("Import", strStatus, GlobalVariable.strFileName, strPostDocNum, string.Format("Error Code : {0} Description : {1} ", GlobalVariable.intErrNum, GlobalVariable.strErrMsg));
                    }
                }

                GC.Collect();
            }
            catch (Exception ex)
            {
                SystemFunction.transHandler("Import", strTransType, "28", "", "", "", dteStart, "E", "-111", ex.Message.ToString());
            }
        }
        private static bool importDIAPIPostDocumentFXML(string strFile)
        {

            string strQuery, strCardCode = "C000200", strNumAtCard, strCodeBars, strShipTo = "", strWhsCode = "", strVatGroup = "";

            string strStatus, strPostDocNum, strTransType = "Document - Sales Order from XML";

            string strBasePrice, strDiscType, strExpCode;

            string strPONumber;

            int intFRow, intDRow = 0, intPONum = 0;

            double dblQuantity, dblPriceAfVat, dblGTotal, dblFAmount, dblDiscRate, dblNetDiscAmt;
            double dblDiscPItem1, dblDiscPItem2, dblDiscPItem3, dblDiscPItem4, dblDiscPItem5;
            double dblTotDisc1, dblTotDisc2, dblTotDisc3, dblTotDisc4, dblTotDisc5;
            double dblNetDiscAmt1, dblNetDiscAmt2, dblNetDiscAmt3, dblNetDiscAmt4, dblNetDiscAmt5;

            DataRow[] oDRFreight;

            char chrExt = ' ';

            SAPbobsCOM.Documents oDocuments;
            SAPbobsCOM.Recordset oRecordset, oRSDiscount;

            DataTable oDTHeader;

            try
            {
                GlobalFunction.getObjType(17);
                if (importXMLData(strFile))
                {
                    if (oDTXML.Rows.Count > 0)
                    {
                        if (!(GlobalVariable.oCompany.InTransaction))
                            GlobalVariable.oCompany.StartTransaction();

                        DataTable[] oDTSplit = oDTXML.AsEnumerable().Select((row, index) => new { row, index }).GroupBy(x => x.index / 7)
                                                                         .Select(g => g.Select(x => x.row).CopyToDataTable()).ToArray();


                        foreach (DataTable oDTSO in oDTSplit)
                        {
                            initFreightDT();
                            oDTHeader = oDTSO.DefaultView.ToTable(true, "OrderId", "OrderDate");

                            if (oDTHeader.Rows.Count > 0)
                            {
                                intDRow = 0;
                                strPONumber = oDTHeader.Rows[0]["OrderId"].ToString();

                                if (intPONum == 0)
                                    strNumAtCard = strPONumber;
                                else if (intPONum == 1)
                                {
                                    chrExt = 'A';
                                    strNumAtCard = strPONumber + "_" + chrExt.ToString();
                                }
                                else
                                {
                                    chrExt++;
                                    strNumAtCard = strPONumber + "_" + chrExt.ToString();
                                }

                                oDocuments = null;
                                oDocuments = (SAPbobsCOM.Documents)GlobalVariable.oCompany.GetBusinessObject(GlobalVariable.oObjectType);

                                oDocuments.CardCode = strCardCode;
                                oDocuments.NumAtCard = strNumAtCard;

                                oDocuments.DocDate = DateTime.Today;
                                oDocuments.DocDueDate = Convert.ToDateTime(oDTHeader.Rows[0]["OrderDate"].ToString());
                                oDocuments.TaxDate = DateTime.Today;

                                oDocuments.UserFields.Fields.Item("U_PODate").Value = Convert.ToDateTime(oDTHeader.Rows[0]["OrderDate"].ToString());

                                for (int intRowD = 0; intRowD <= oDTSO.Rows.Count - 1; intRowD++)
                                {
                                    strCodeBars = oDTSO.Rows[intRowD]["BarCode"].ToString();
                                    strShipTo = oDTSO.Rows[intRowD]["Shipto"].ToString();
                                    dblQuantity = Convert.ToDouble(oDTSO.Rows[intRowD]["Quantity"].ToString());

                                    strQuery = string.Format("CALL \"FTSI_IMPORT_SO_MULTIDISCOUNT_BASE\" ('{0}', '{1}', '{2}');", strCodeBars, strCardCode, strShipTo);

                                    oRecordset = null;
                                    oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                    oRecordset.DoQuery(strQuery);

                                    if (oRecordset.RecordCount > 0)
                                    {
                                        strWhsCode = oRecordset.Fields.Item("U_Whscode").Value.ToString();
                                        strVatGroup = oRecordset.Fields.Item("ECVatGroup").Value.ToString();

                                        dblPriceAfVat = Convert.ToDouble(oRecordset.Fields.Item("PriceAfVat").Value.ToString());

                                        dblGTotal = dblQuantity * dblPriceAfVat;

                                        if (intRowD > 0)
                                            oDocuments.Lines.Add();

                                        oDocuments.Lines.BarCode = strCodeBars;
                                        oDocuments.Lines.Quantity = dblQuantity;
                                        oDocuments.Lines.WarehouseCode = strWhsCode;
                                        oDocuments.Lines.UnitPrice = dblPriceAfVat;
                                        oDocuments.Lines.VatGroup = strVatGroup;
                                        oDocuments.Lines.UserFields.Fields.Item("U_WhseName").Value = strShipTo;

                                        #region "LEVEL 1"

                                        strQuery = string.Format("CALL \"FTSI_IMPORT_SO_MULTIDISCOUNT_LEVEL\" ('{0}', '{1}', '{2}', to_date('{3}', 'MM/dd/yyyy'))", strCodeBars, strCardCode, "1", DateTime.Today.ToString("MM/dd/yyyy"));


                                        oRSDiscount = null;
                                        oRSDiscount = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        oRSDiscount.DoQuery(strQuery);

                                        if (oRSDiscount.RecordCount > 0)
                                        {
                                            strDiscType = oRSDiscount.Fields.Item("U_DiscType").Value.ToString();
                                            strExpCode = oRSDiscount.Fields.Item("U_FrghtCode").Value.ToString();
                                            dblDiscRate = Convert.ToDouble(oRSDiscount.Fields.Item("U_DiscRate").Value.ToString());

                                            dblDiscPItem1 = dblPriceAfVat * dblDiscRate / 100;
                                            dblTotDisc1 = dblQuantity * dblDiscPItem1;
                                            dblNetDiscAmt1 = dblGTotal - dblTotDisc1;
                                            dblNetDiscAmt = dblNetDiscAmt1;
                                        }
                                        else
                                        {
                                            strDiscType = "";
                                            strExpCode = "";
                                            dblDiscRate = 0;
                                            dblDiscPItem1 = 0;
                                            dblTotDisc1 = 0;
                                            dblNetDiscAmt1 = dblGTotal - dblTotDisc1;
                                            dblNetDiscAmt = dblNetDiscAmt1;
                                        }

                                        if (dblTotDisc1 > 0)
                                        {
                                            oDocuments.Lines.UserFields.Fields.Item("U_DiscType1").Value = strDiscType;
                                            oDocuments.Lines.UserFields.Fields.Item("U_DiscRate1").Value = dblDiscRate;
                                            oDocuments.Lines.UserFields.Fields.Item("U_DiscPerItem1").Value = dblDiscPItem1;
                                            oDocuments.Lines.UserFields.Fields.Item("U_DiscAmt1").Value = dblTotDisc1;
                                            oDocuments.Lines.UserFields.Fields.Item("U_NetofDisc1").Value = dblNetDiscAmt1;

                                            oDRFreight = oDTFreight.Select("ExpCode = '" + strExpCode + "' ");

                                            if (oDRFreight.Length > 0)
                                            {
                                                intFRow = Convert.ToInt32(oDRFreight[0]["Row"]);
                                                oDTFreight.Rows[intFRow]["Amount"] = dblTotDisc1 +
                                                                                               Convert.ToDouble(oDTFreight.Rows[intFRow]["Amount"]);
                                            }
                                            else
                                            {
                                                oDTFreight.Rows.Add(intDRow,
                                                                    strExpCode,
                                                                    strVatGroup,
                                                                    dblTotDisc1);

                                                intDRow++;
                                            }
                                        }

                                        #endregion

                                        #region "LEVEL 2"

                                        strQuery = string.Format("CALL \"FTSI_IMPORT_SO_MULTIDISCOUNT_LEVEL\" ('{0}', '{1}', '{2}', to_date('{3}', 'MM/dd/yyyy'))", strCodeBars, strCardCode, "2", DateTime.Today.ToString("MM/dd/yyyy"));

                                        oRSDiscount = null;
                                        oRSDiscount = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        oRSDiscount.DoQuery(strQuery);

                                        if (oRSDiscount.RecordCount > 0)
                                        {
                                            strDiscType = oRSDiscount.Fields.Item("U_DiscType").Value.ToString();
                                            dblDiscRate = Convert.ToDouble(oRSDiscount.Fields.Item("U_DiscRate").Value.ToString());
                                            strExpCode = oRSDiscount.Fields.Item("U_FrghtCode").Value.ToString();
                                            strBasePrice = oRSDiscount.Fields.Item("U_BasePrice").Value.ToString();

                                            if (strBasePrice == "G")
                                                dblDiscPItem2 = dblPriceAfVat * dblDiscRate / 100;
                                            else
                                                dblDiscPItem2 = dblNetDiscAmt / dblQuantity * dblDiscRate / 100;

                                            dblTotDisc2 = dblQuantity * dblDiscPItem2;
                                            dblNetDiscAmt2 = dblNetDiscAmt - dblTotDisc2;
                                            dblNetDiscAmt = dblNetDiscAmt2;


                                        }
                                        else
                                        {
                                            strDiscType = "";
                                            strExpCode = "";
                                            dblDiscRate = 0;
                                            dblDiscPItem2 = 0;
                                            dblTotDisc2 = 0;
                                            dblNetDiscAmt2 = dblNetDiscAmt - dblTotDisc2;
                                            dblNetDiscAmt = dblNetDiscAmt2;
                                        }

                                        if (dblTotDisc2 > 0)
                                        {
                                            oDocuments.Lines.UserFields.Fields.Item("U_DiscType2").Value = strDiscType;
                                            oDocuments.Lines.UserFields.Fields.Item("U_DiscRate2").Value = dblDiscRate;
                                            oDocuments.Lines.UserFields.Fields.Item("U_DiscPerItem2").Value = dblDiscPItem2;
                                            oDocuments.Lines.UserFields.Fields.Item("U_DiscAmt2").Value = dblTotDisc2;
                                            oDocuments.Lines.UserFields.Fields.Item("U_NetofDisc2").Value = dblNetDiscAmt2;

                                            oDRFreight = oDTFreight.Select("ExpCode = '" + strExpCode + "' ");

                                            if (oDRFreight.Length > 0)
                                            {
                                                intFRow = Convert.ToInt32(oDRFreight[0]["Row"]);
                                                oDTFreight.Rows[intFRow]["Amount"] = dblTotDisc2 +
                                                                                               Convert.ToDouble(oDTFreight.Rows[intFRow]["Amount"]);
                                            }
                                            else
                                            {
                                                oDTFreight.Rows.Add(intDRow,
                                                                              strExpCode,
                                                                              strVatGroup,
                                                                              dblTotDisc2);

                                                intDRow++;
                                            }
                                        }

                                        #endregion

                                        #region "LEVEL 3"

                                        strQuery = string.Format("CALL \"FTSI_IMPORT_SO_MULTIDISCOUNT_LEVEL\" ('{0}', '{1}', '{2}', to_date('{3}', 'MM/dd/yyyy'))", strCodeBars, strCardCode, "3", DateTime.Today.ToString("MM/dd/yyyy"));

                                        oRSDiscount = null;
                                        oRSDiscount = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        oRSDiscount.DoQuery(strQuery);
                                        int check = oRSDiscount.RecordCount;
                                        if (oRSDiscount.RecordCount > 0)
                                        {
                                            strDiscType = oRSDiscount.Fields.Item("U_DiscType").Value.ToString();
                                            dblDiscRate = Convert.ToDouble(oRSDiscount.Fields.Item("U_DiscRate").Value.ToString());

                                            strExpCode = oRSDiscount.Fields.Item("U_FrghtCode").Value.ToString();

                                            strBasePrice = oRSDiscount.Fields.Item("U_BasePrice").Value.ToString();

                                            if (strBasePrice == "G")
                                                dblDiscPItem3 = dblPriceAfVat * dblDiscRate / 100;
                                            else
                                                dblDiscPItem3 = dblNetDiscAmt / dblQuantity * dblDiscRate / 100;

                                            dblTotDisc3 = dblQuantity * dblDiscPItem3;
                                            dblNetDiscAmt3 = dblNetDiscAmt - dblTotDisc3;
                                            dblNetDiscAmt = dblNetDiscAmt3;


                                        }
                                        else
                                        {
                                            strDiscType = "";
                                            strExpCode = "";
                                            dblDiscRate = 0;
                                            dblDiscPItem3 = 0;
                                            dblTotDisc3 = 0;
                                            dblNetDiscAmt3 = dblNetDiscAmt - dblTotDisc3;
                                            dblNetDiscAmt = dblNetDiscAmt3;
                                        }

                                        if (dblTotDisc3 > 0)
                                        {
                                            oDocuments.Lines.UserFields.Fields.Item("U_DiscType3").Value = strDiscType;
                                            oDocuments.Lines.UserFields.Fields.Item("U_DiscRate3").Value = dblDiscRate;
                                            oDocuments.Lines.UserFields.Fields.Item("U_DiscPerItem3").Value = dblDiscPItem3;
                                            oDocuments.Lines.UserFields.Fields.Item("U_DiscAmt3").Value = dblTotDisc3;
                                            oDocuments.Lines.UserFields.Fields.Item("U_NetofDisc3").Value = dblNetDiscAmt3;


                                            oDRFreight = oDTFreight.Select("ExpCode = '" + strExpCode + "' ");

                                            if (oDRFreight.Length > 0)
                                            {
                                                intFRow = Convert.ToInt32(oDRFreight[0]["Row"]);
                                                oDTFreight.Rows[intFRow]["Amount"] = dblTotDisc3 +
                                                                                               Convert.ToDouble(oDTFreight.Rows[intFRow]["Amount"]);
                                            }
                                            else
                                            {
                                                oDTFreight.Rows.Add(intDRow,
                                                                              strExpCode,
                                                                              strVatGroup,
                                                                              dblTotDisc3);

                                                intDRow++;
                                            }
                                        }

                                        #endregion

                                        #region "LEVEL 4"

                                        strQuery = string.Format("CALL \"FTSI_IMPORT_SO_MULTIDISCOUNT_LEVEL\" ('{0}', '{1}', '{2}', to_date('{3}', 'MM/dd/yyyy'))", strCodeBars, strCardCode, "4", DateTime.Today.ToString("MM/dd/yyyy"));

                                        oRSDiscount = null;
                                        oRSDiscount = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        oRSDiscount.DoQuery(strQuery);

                                        if (oRSDiscount.RecordCount > 0)
                                        {
                                            strDiscType = oRSDiscount.Fields.Item("U_DiscType").Value.ToString();
                                            dblDiscRate = Convert.ToDouble(oRSDiscount.Fields.Item("U_DiscRate").Value.ToString());

                                            strExpCode = oRSDiscount.Fields.Item("U_FrghtCode").Value.ToString();

                                            strBasePrice = oRSDiscount.Fields.Item("U_BasePrice").Value.ToString();

                                            if (strBasePrice == "G")
                                                dblDiscPItem4 = dblPriceAfVat * dblDiscRate / 100;
                                            else
                                                dblDiscPItem4 = dblNetDiscAmt / dblQuantity * dblDiscRate / 100;

                                            dblTotDisc4 = dblQuantity * dblDiscPItem4;
                                            dblNetDiscAmt4 = dblNetDiscAmt - dblTotDisc4;
                                            dblNetDiscAmt = dblNetDiscAmt4;


                                        }
                                        else
                                        {
                                            strDiscType = "";
                                            strExpCode = "";
                                            dblDiscRate = 0;
                                            dblDiscPItem4 = 0;
                                            dblTotDisc4 = 0;
                                            dblNetDiscAmt4 = dblNetDiscAmt - dblTotDisc4;
                                            dblNetDiscAmt = dblNetDiscAmt4;
                                        }
                                        if (dblTotDisc4 > 0)
                                        {
                                            oDocuments.Lines.UserFields.Fields.Item("U_DiscType4").Value = strDiscType;
                                            oDocuments.Lines.UserFields.Fields.Item("U_DiscRate4").Value = dblDiscRate;
                                            oDocuments.Lines.UserFields.Fields.Item("U_DiscPerItem4").Value = dblDiscPItem4;
                                            oDocuments.Lines.UserFields.Fields.Item("U_DiscAmt4").Value = dblTotDisc4;
                                            oDocuments.Lines.UserFields.Fields.Item("U_NetofDisc4").Value = dblNetDiscAmt4;

                                            oDRFreight = oDTFreight.Select("ExpCode = '" + strExpCode + "' ");

                                            if (oDRFreight.Length > 0)
                                            {
                                                intFRow = Convert.ToInt32(oDRFreight[0]["Row"]);
                                                oDTFreight.Rows[intFRow]["Amount"] = dblTotDisc4 +
                                                                                               Convert.ToDouble(oDTFreight.Rows[intFRow]["Amount"]);
                                            }
                                            else
                                            {
                                                oDTFreight.Rows.Add(intDRow,
                                                                              strExpCode,
                                                                              strVatGroup,
                                                                              dblTotDisc4);

                                                intDRow++;
                                            }
                                        }


                                        #endregion

                                        #region "LEVEL 5"

                                        strQuery = string.Format("CALL \"FTSI_IMPORT_SO_MULTIDISCOUNT_LEVEL\" ('{0}', '{1}', '{2}', to_date('{3}', 'MM/dd/yyyy'))", strCodeBars, strCardCode, "5", DateTime.Today.ToString("MM/dd/yyyy"));

                                        oRSDiscount = null;
                                        oRSDiscount = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        oRSDiscount.DoQuery(strQuery);

                                        if (oRSDiscount.RecordCount > 0)
                                        {
                                            strDiscType = oRSDiscount.Fields.Item("U_DiscType").Value.ToString();
                                            dblDiscRate = Convert.ToDouble(oRSDiscount.Fields.Item("U_DiscRate").Value.ToString());

                                            strExpCode = oRSDiscount.Fields.Item("U_FrghtCode").Value.ToString();

                                            strBasePrice = oRSDiscount.Fields.Item("U_BasePrice").Value.ToString();

                                            if (strBasePrice == "G")
                                                dblDiscPItem5 = dblPriceAfVat * dblDiscRate / 100;
                                            else
                                                dblDiscPItem5 = dblNetDiscAmt / dblQuantity * dblDiscRate / 100;

                                            dblTotDisc5 = dblQuantity * dblDiscPItem5;
                                            dblNetDiscAmt5 = dblNetDiscAmt - dblTotDisc5;
                                            dblNetDiscAmt = dblNetDiscAmt4;


                                        }
                                        else
                                        {
                                            strDiscType = "";
                                            strExpCode = "";
                                            dblDiscRate = 0;
                                            dblDiscPItem5 = 0;
                                            dblTotDisc5 = 0;
                                            dblNetDiscAmt5 = dblNetDiscAmt - dblTotDisc5;
                                            dblNetDiscAmt = dblNetDiscAmt5;
                                        }
                                        if (dblTotDisc5 > 0)
                                        {
                                            oDocuments.Lines.UserFields.Fields.Item("U_DiscType5").Value = strDiscType;
                                            oDocuments.Lines.UserFields.Fields.Item("U_DiscRate5").Value = dblDiscRate;
                                            oDocuments.Lines.UserFields.Fields.Item("U_DiscPerItem5").Value = dblDiscPItem5;
                                            oDocuments.Lines.UserFields.Fields.Item("U_DiscAmt5").Value = dblTotDisc5;
                                            oDocuments.Lines.UserFields.Fields.Item("U_NetofDisc5").Value = dblNetDiscAmt5;

                                            oDRFreight = oDTFreight.Select("ExpCode = '" + strExpCode + "' ");

                                            if (oDRFreight.Length > 0)
                                            {
                                                intFRow = Convert.ToInt32(oDRFreight[0]["Row"]);
                                                oDTFreight.Rows[intFRow]["Amount"] = dblTotDisc5 +
                                                                                               Convert.ToDouble(oDTFreight.Rows[intFRow]["Amount"]);
                                            }
                                            else
                                            {
                                                oDTFreight.Rows.Add(intDRow,
                                                                              strExpCode,
                                                                              strVatGroup,
                                                                              dblTotDisc5);

                                                intDRow++;
                                            }
                                        }

                                        #endregion
                                    }

                                    else
                                    {

                                        strQuery = string.Format("SELECT CRD1.\"Address\", CRD1.\"U_Whscode\", OCRD.\"ECVatGroup\" FROM CRD1 INNER JOIN OCRD ON CRD1.\"CardCode\" =  OCRD.\"CardCode\" WHERE CRD1.\"CardCode\" = '{0}' AND CRD1.\"Address\"= '{1}'", strCardCode, strShipTo);

                                        oRecordset = null;
                                        oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        oRecordset.DoQuery(strQuery);

                                        if (oRecordset.RecordCount == 0)
                                        {
                                            //to do raise error
                                            GlobalVariable.intErrNum = GlobalVariable.oCompany.GetLastErrorCode();
                                            GlobalVariable.strErrMsg = GlobalVariable.oCompany.GetLastErrorDescription();

                                            strStatus = "E";
                                            strMsgBod = string.Format("Error Posting {0}, Invalid ShipTo Address " + strShipTo + " ShipTo Address is Not found in database ", GlobalVariable.strFileName);

                                            SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), strMsgBod);

                                            return false;
                                        }


                                        strQuery = string.Format("SELECT OITM.\"ItemCode\", OITM.\"ItemName\" FROM OITM INNER JOIN OBCD ON OITM.\"ItemCode\" = OBCD.\"ItemCode\" WHERE OBCD.\"BcdCode\" = '{0}'", strCodeBars);

                                        oRecordset = null;
                                        oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        oRecordset.DoQuery(strQuery);

                                        if (oRecordset.RecordCount == 0)
                                        {
                                            //to do raise error
                                            GlobalVariable.intErrNum = GlobalVariable.oCompany.GetLastErrorCode();
                                            GlobalVariable.strErrMsg = GlobalVariable.oCompany.GetLastErrorDescription();

                                            strStatus = "E";
                                            strMsgBod = string.Format("Error Posting {0}, Invalid Barcode No." + strCodeBars + " Barcode Not found in database ", GlobalVariable.strFileName);

                                            SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), strMsgBod);

                                            return false;
                                        }
                                    }
                                }

                                intPONum++;

                                if (oDTFreight.Rows.Count > 0)
                                {
                                    for (int intRow = 0; intRow < oDTFreight.Rows.Count; intRow++)
                                    {

                                        dblFAmount = Convert.ToDouble(oDTFreight.Rows[intRow]["Amount"]);
                                        if (dblFAmount > 0)
                                        {

                                            oDocuments.Expenses.ExpenseCode = Convert.ToInt32(oDTFreight.Rows[intRow]["ExpCode"].ToString());
                                            oDocuments.Expenses.LineGross = dblFAmount * -1;
                                            oDocuments.Expenses.VatGroup = oDTFreight.Rows[intRow]["VatGroup"].ToString();
                                            oDocuments.Expenses.Add();

                                        }
                                    }

                                }

                                oDocuments.ShipToCode = strShipTo;

                                if (oDocuments.Add() != 0)
                                {
                                    GlobalVariable.intErrNum = GlobalVariable.oCompany.GetLastErrorCode();
                                    GlobalVariable.strErrMsg = GlobalVariable.oCompany.GetLastErrorDescription();

                                    strStatus = "E";
                                    strMsgBod = string.Format("Error Posting {0}.\rError Code: {1}\rDescription: {2} ", GlobalVariable.strFileName, GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                                    SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                                    return false;
                                }
                                else
                                {
                                    GlobalVariable.intErrNum = 0;
                                    GlobalVariable.strErrMsg = GlobalVariable.oCompany.GetLastErrorDescription();

                                    strPostDocNum = GlobalVariable.oCompany.GetNewObjectKey().ToString();

                                    strMsgBod = string.Format("Successfully Posted Sales Order from {0}. Posted Document Number: {1} ", GlobalVariable.strFileName, strPostDocNum);

                                    SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, GlobalVariable.oCompany.GetNewObjectKey(), strPostDocNum, dteStart, "S", "0", strMsgBod);

                                    strStatus = "S";
                                }
                            }

                        }

                        if (GlobalVariable.oCompany.InTransaction)
                            GlobalVariable.oCompany.EndTransaction(BoWfTransOpt.wf_Commit);

                    }

                    GC.Collect();

                    return true;
                }
                else
                    return false;


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
        private static bool importDIAPIPostDocumentFCSV(string strFile)
        {

            string strQuery, strCardCode = "C000121", strItemCode, strNumAtCard, strCodeBars, strShipTo = "", strWhsCode = "", strVatGroup = "";

            string strStatus, strPostDocNum, strTransType = "Document - Sales Order from CSV";

            string strBasePrice, strDiscType, strExpCode;

            string strPONumber;

            int intFRow, intDRow = 0, intPONum = 0;

            double dblQuantity, dblPriceAfVat, dblGTotal, dblFAmount, dblDiscRate, dblNetDiscAmt;
            double dblDiscPItem1, dblDiscPItem2, dblDiscPItem3, dblDiscPItem4, dblDiscPItem5;
            double dblTotDisc1, dblTotDisc2, dblTotDisc3, dblTotDisc4, dblTotDisc5;
            double dblNetDiscAmt1, dblNetDiscAmt2, dblNetDiscAmt3, dblNetDiscAmt4, dblNetDiscAmt5;

            char chrExt = ' ';

            DataRow[] oDRFreight;

            SAPbobsCOM.Documents oDocuments;
            SAPbobsCOM.Recordset oRecordset, oRSDiscount;

            DataTable oDTHeader;

            try
            {
                GlobalFunction.getObjType(17);

                if (GlobalFunction.importCSV(Path.GetDirectoryName(strFile), Path.GetFileName(strFile), "YES", GlobalVariable.chrDlmtr.ToString()))
                {

                    if (GlobalVariable.oDTImpData.Rows.Count > 0)
                    {

                        if (!(GlobalVariable.oCompany.InTransaction))
                            GlobalVariable.oCompany.StartTransaction();

                        DataTable[] oDTSplit = GlobalVariable.oDTImpData.AsEnumerable().Select((row, index) => new { row, index }).GroupBy(x => x.index / 7)
                                                                                       .Select(g => g.Select(x => x.row).CopyToDataTable()).ToArray();


                        foreach (DataTable oDTSO in oDTSplit)
                        {

                            initFreightDT();
                            oDTHeader = oDTSO.DefaultView.ToTable(true, "PO Number", "Receipt Date", "Cancel Date");

                            if (oDTHeader.Rows.Count > 0)
                            {
                                strPONumber = oDTHeader.Rows[0]["PO Number"].ToString();

                                intDRow = 0;

                                if (intPONum == 0)
                                    strNumAtCard = strPONumber;
                                else if (intPONum == 1)
                                {
                                    chrExt = 'A';
                                    strNumAtCard = strPONumber + "_" + chrExt.ToString();
                                }
                                else
                                {
                                    chrExt++;
                                    strNumAtCard = strPONumber + "_" + chrExt.ToString();
                                }

                                oDocuments = null;
                                oDocuments = (SAPbobsCOM.Documents)GlobalVariable.oCompany.GetBusinessObject(GlobalVariable.oObjectType);

                                oDocuments.CardCode = strCardCode;
                                oDocuments.NumAtCard = strNumAtCard;

                                oDocuments.DocDate = DateTime.Today;
                                oDocuments.DocDueDate = Convert.ToDateTime(oDTHeader.Rows[0]["Receipt Date"].ToString());
                                oDocuments.TaxDate = DateTime.Today;

                                oDocuments.UserFields.Fields.Item("U_PODate").Value = Convert.ToDateTime(GlobalVariable.oDTImpData.Rows[0]["Receipt Date"].ToString());


                                for (int intRowD = 0; intRowD <= oDTSO.Rows.Count - 1; intRowD++)
                                {
                                    strCodeBars = oDTSO.Rows[intRowD]["UPC"].ToString();
                                    strShipTo = oDTSO.Rows[intRowD]["Shipto"].ToString();
                                    dblQuantity = Convert.ToDouble(oDTSO.Rows[intRowD]["BuyQty"].ToString());

                                    strQuery = string.Format("CALL \"FTSI_IMPORT_SO_MULTIDISCOUNT_BASE\" ('{0}', '{1}', '{2}');", strCodeBars, strCardCode, strShipTo);

                                    oRecordset = null;
                                    oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                    oRecordset.DoQuery(strQuery);

                                    if (oRecordset.RecordCount > 0)
                                    {
                                        strWhsCode = oRecordset.Fields.Item("U_Whscode").Value.ToString();
                                        strVatGroup = oRecordset.Fields.Item("ECVatGroup").Value.ToString();

                                        dblPriceAfVat = Convert.ToDouble(oRecordset.Fields.Item("PriceAfVat").Value.ToString());

                                        dblGTotal = dblQuantity * dblPriceAfVat;

                                        if (intRowD > 0)
                                            oDocuments.Lines.Add();

                                        oDocuments.Lines.BarCode = strCodeBars;
                                        oDocuments.Lines.Quantity = dblQuantity;
                                        oDocuments.Lines.WarehouseCode = strWhsCode;
                                        oDocuments.Lines.UnitPrice = dblPriceAfVat;
                                        oDocuments.Lines.VatGroup = strVatGroup;
                                        oDocuments.Lines.UserFields.Fields.Item("U_WhseName").Value = strShipTo;

                                        #region "LEVEL 1"

                                        strQuery = string.Format("CALL \"FTSI_IMPORT_SO_MULTIDISCOUNT_LEVEL\" ('{0}', '{1}', '{2}', to_date('{3}', 'MM/dd/yyyy'))", strCodeBars, strCardCode, "1", DateTime.Today.ToString("MM/dd/yyyy"));

                                        oRSDiscount = null;
                                        oRSDiscount = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        oRSDiscount.DoQuery(strQuery);

                                        if (oRSDiscount.RecordCount > 0)
                                        {
                                            strDiscType = oRSDiscount.Fields.Item("U_DiscType").Value.ToString();
                                            strExpCode = oRSDiscount.Fields.Item("U_FrghtCode").Value.ToString();
                                            dblDiscRate = Convert.ToDouble(oRSDiscount.Fields.Item("U_DiscRate").Value.ToString());

                                            dblDiscPItem1 = dblPriceAfVat * dblDiscRate / 100;
                                            dblTotDisc1 = dblQuantity * dblDiscPItem1;
                                            dblNetDiscAmt1 = dblGTotal - dblTotDisc1;
                                            dblNetDiscAmt = dblNetDiscAmt1;
                                        }
                                        else
                                        {
                                            strDiscType = "";
                                            strExpCode = "";
                                            dblDiscRate = 0;
                                            dblDiscPItem1 = 0;
                                            dblTotDisc1 = 0;
                                            dblNetDiscAmt1 = dblGTotal - dblTotDisc1;
                                            dblNetDiscAmt = dblNetDiscAmt1;
                                        }

                                        if (dblTotDisc1 > 0)
                                        {
                                            //oDocuments.Lines.UserFields.Fields.Item("U_DiscType1").Value = strDiscType;
                                            oDocuments.Lines.UserFields.Fields.Item("U_DiscType1").Value = strDiscType;
                                            oDocuments.Lines.UserFields.Fields.Item("U_DiscRate1").Value = dblDiscRate;
                                            oDocuments.Lines.UserFields.Fields.Item("U_DiscPerItem1").Value = dblDiscPItem1;
                                            oDocuments.Lines.UserFields.Fields.Item("U_DiscAmt1").Value = dblTotDisc1;
                                            oDocuments.Lines.UserFields.Fields.Item("U_NetofDisc1").Value = dblNetDiscAmt1;

                                            oDRFreight = oDTFreight.Select("ExpCode = '" + strExpCode + "' ");

                                            if (oDRFreight.Length > 0)
                                            {
                                                intFRow = Convert.ToInt32(oDRFreight[0]["Row"]);
                                                oDTFreight.Rows[intFRow]["Amount"] = dblTotDisc1 +
                                                                                               Convert.ToDouble(oDTFreight.Rows[intFRow]["Amount"]);
                                            }
                                            else
                                            {
                                                oDTFreight.Rows.Add(intDRow,
                                                                    strExpCode,
                                                                    strVatGroup,
                                                                    dblTotDisc1);

                                                intDRow++;
                                            }
                                        }

                                        #endregion

                                        #region "LEVEL 2"

                                        strQuery = string.Format("CALL \"FTSI_IMPORT_SO_MULTIDISCOUNT_LEVEL\" ('{0}', '{1}', '{2}', to_date('{3}', 'MM/dd/yyyy'))", strCodeBars, strCardCode, "2", DateTime.Today.ToString("MM/dd/yyyy"));

                                        oRSDiscount = null;
                                        oRSDiscount = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        oRSDiscount.DoQuery(strQuery);

                                        if (oRSDiscount.RecordCount > 0)
                                        {
                                            strDiscType = oRSDiscount.Fields.Item("U_DiscType").Value.ToString();
                                            dblDiscRate = Convert.ToDouble(oRSDiscount.Fields.Item("U_DiscRate").Value.ToString());
                                            strExpCode = oRSDiscount.Fields.Item("U_FrghtCode").Value.ToString();
                                            strBasePrice = oRSDiscount.Fields.Item("U_BasePrice").Value.ToString();

                                            if (strBasePrice == "G")
                                                dblDiscPItem2 = dblPriceAfVat * dblDiscRate / 100;
                                            else
                                                dblDiscPItem2 = dblNetDiscAmt / dblQuantity * dblDiscRate / 100;

                                            dblTotDisc2 = dblQuantity * dblDiscPItem2;
                                            dblNetDiscAmt2 = dblNetDiscAmt - dblTotDisc2;
                                            dblNetDiscAmt = dblNetDiscAmt2;


                                        }
                                        else
                                        {
                                            strDiscType = "";
                                            strExpCode = "";
                                            dblDiscRate = 0;
                                            dblDiscPItem2 = 0;
                                            dblTotDisc2 = 0;
                                            dblNetDiscAmt2 = dblNetDiscAmt - dblTotDisc2;
                                            dblNetDiscAmt = dblNetDiscAmt2;
                                        }

                                        if (dblTotDisc2 > 0)
                                        {
                                            oDocuments.Lines.UserFields.Fields.Item("U_DiscType2").Value = strDiscType;
                                            oDocuments.Lines.UserFields.Fields.Item("U_DiscRate2").Value = dblDiscRate;
                                            oDocuments.Lines.UserFields.Fields.Item("U_DiscPerItem2").Value = dblDiscPItem2;
                                            oDocuments.Lines.UserFields.Fields.Item("U_DiscAmt2").Value = dblTotDisc2;
                                            oDocuments.Lines.UserFields.Fields.Item("U_NetofDisc2").Value = dblNetDiscAmt2;



                                            oDRFreight = oDTFreight.Select("ExpCode = '" + strExpCode + "' ");

                                            if (oDRFreight.Length > 0)
                                            {
                                                intFRow = Convert.ToInt32(oDRFreight[0]["Row"]);
                                                oDTFreight.Rows[intFRow]["Amount"] = dblTotDisc2 +
                                                                                               Convert.ToDouble(oDTFreight.Rows[intFRow]["Amount"]);
                                            }
                                            else
                                            {
                                                oDTFreight.Rows.Add(intDRow,
                                                                              strExpCode,
                                                                              strVatGroup,
                                                                              dblTotDisc2);

                                                intDRow++;
                                            }
                                        }

                                        #endregion

                                        #region "LEVEL 3"

                                        strQuery = string.Format("CALL \"FTSI_IMPORT_SO_MULTIDISCOUNT_LEVEL\" ('{0}', '{1}', '{2}', to_date('{3}', 'MM/dd/yyyy'))", strCodeBars, strCardCode, "3", DateTime.Today.ToString("MM/dd/yyyy"));

                                        oRSDiscount = null;
                                        oRSDiscount = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        oRSDiscount.DoQuery(strQuery);
                                        int check = oRSDiscount.RecordCount;
                                        if (oRSDiscount.RecordCount > 0)
                                        {
                                            strDiscType = oRSDiscount.Fields.Item("U_DiscType").Value.ToString();
                                            dblDiscRate = Convert.ToDouble(oRSDiscount.Fields.Item("U_DiscRate").Value.ToString());

                                            strExpCode = oRSDiscount.Fields.Item("U_FrghtCode").Value.ToString();

                                            strBasePrice = oRSDiscount.Fields.Item("U_BasePrice").Value.ToString();

                                            if (strBasePrice == "G")
                                                dblDiscPItem3 = dblPriceAfVat * dblDiscRate / 100;
                                            else
                                                dblDiscPItem3 = dblNetDiscAmt / dblQuantity * dblDiscRate / 100;

                                            dblTotDisc3 = dblQuantity * dblDiscPItem3;
                                            dblNetDiscAmt3 = dblNetDiscAmt - dblTotDisc3;
                                            dblNetDiscAmt = dblNetDiscAmt3;


                                        }
                                        else
                                        {
                                            strDiscType = "";
                                            strExpCode = "";
                                            dblDiscRate = 0;
                                            dblDiscPItem3 = 0;
                                            dblTotDisc3 = 0;
                                            dblNetDiscAmt3 = dblNetDiscAmt - dblTotDisc3;
                                            dblNetDiscAmt = dblNetDiscAmt3;
                                        }

                                        if (dblTotDisc3 > 0)
                                        {
                                            oDocuments.Lines.UserFields.Fields.Item("U_DiscType3").Value = strDiscType;
                                            oDocuments.Lines.UserFields.Fields.Item("U_DiscRate3").Value = dblDiscRate;
                                            oDocuments.Lines.UserFields.Fields.Item("U_DiscPerItem3").Value = dblDiscPItem3;
                                            oDocuments.Lines.UserFields.Fields.Item("U_DiscAmt3").Value = dblTotDisc3;
                                            oDocuments.Lines.UserFields.Fields.Item("U_NetofDisc3").Value = dblNetDiscAmt3;



                                            oDRFreight = oDTFreight.Select("ExpCode = '" + strExpCode + "' ");

                                            if (oDRFreight.Length > 0)
                                            {
                                                intFRow = Convert.ToInt32(oDRFreight[0]["Row"]);
                                                oDTFreight.Rows[intFRow]["Amount"] = dblTotDisc3 +
                                                                                               Convert.ToDouble(oDTFreight.Rows[intFRow]["Amount"]);
                                            }
                                            else
                                            {
                                                oDTFreight.Rows.Add(intDRow,
                                                                              strExpCode,
                                                                              strVatGroup,
                                                                              dblTotDisc3);

                                                intDRow++;
                                            }
                                        }



                                        #endregion

                                        #region "LEVEL 4"

                                        strQuery = string.Format("CALL \"FTSI_IMPORT_SO_MULTIDISCOUNT_LEVEL\" ('{0}', '{1}', '{2}', to_date('{3}', 'MM/dd/yyyy'))", strCodeBars, strCardCode, "4", DateTime.Today.ToString("MM/dd/yyyy"));

                                        oRSDiscount = null;
                                        oRSDiscount = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        oRSDiscount.DoQuery(strQuery);

                                        if (oRSDiscount.RecordCount > 0)
                                        {
                                            strDiscType = oRSDiscount.Fields.Item("U_DiscType").Value.ToString();
                                            dblDiscRate = Convert.ToDouble(oRSDiscount.Fields.Item("U_DiscRate").Value.ToString());

                                            strExpCode = oRSDiscount.Fields.Item("U_FrghtCode").Value.ToString();

                                            strBasePrice = oRSDiscount.Fields.Item("U_BasePrice").Value.ToString();

                                            if (strBasePrice == "G")
                                                dblDiscPItem4 = dblPriceAfVat * dblDiscRate / 100;
                                            else
                                                dblDiscPItem4 = dblNetDiscAmt / dblQuantity * dblDiscRate / 100;

                                            dblTotDisc4 = dblQuantity * dblDiscPItem4;
                                            dblNetDiscAmt4 = dblNetDiscAmt - dblTotDisc4;
                                            dblNetDiscAmt = dblNetDiscAmt4;


                                        }
                                        else
                                        {
                                            strDiscType = "";
                                            strExpCode = "";
                                            dblDiscRate = 0;
                                            dblDiscPItem4 = 0;
                                            dblTotDisc4 = 0;
                                            dblNetDiscAmt4 = dblNetDiscAmt - dblTotDisc4;
                                            dblNetDiscAmt = dblNetDiscAmt4;
                                        }
                                        if (dblTotDisc4 > 0)
                                        {
                                            oDocuments.Lines.UserFields.Fields.Item("U_DiscType4").Value = strDiscType;
                                            oDocuments.Lines.UserFields.Fields.Item("U_DiscRate4").Value = dblDiscRate;
                                            oDocuments.Lines.UserFields.Fields.Item("U_DiscPerItem4").Value = dblDiscPItem4;
                                            oDocuments.Lines.UserFields.Fields.Item("U_DiscAmt4").Value = dblTotDisc4;
                                            oDocuments.Lines.UserFields.Fields.Item("U_NetofDisc4").Value = dblNetDiscAmt4;

                                            oDRFreight = oDTFreight.Select("ExpCode = '" + strExpCode + "' ");

                                            if (oDRFreight.Length > 0)
                                            {
                                                intFRow = Convert.ToInt32(oDRFreight[0]["Row"]);
                                                oDTFreight.Rows[intFRow]["Amount"] = dblTotDisc4 +
                                                                                               Convert.ToDouble(oDTFreight.Rows[intFRow]["Amount"]);
                                            }
                                            else
                                            {
                                                oDTFreight.Rows.Add(intDRow,
                                                                              strExpCode,
                                                                              strVatGroup,
                                                                              dblTotDisc4);

                                                intDRow++;
                                            }
                                        }


                                        #endregion

                                        #region "LEVEL 5"

                                        strQuery = string.Format("CALL \"FTSI_IMPORT_SO_MULTIDISCOUNT_LEVEL\" ('{0}', '{1}', '{2}', to_date('{3}', 'MM/dd/yyyy'))", strCodeBars, strCardCode, "5", DateTime.Today.ToString("MM/dd/yyyy"));

                                        oRSDiscount = null;
                                        oRSDiscount = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        oRSDiscount.DoQuery(strQuery);

                                        if (oRSDiscount.RecordCount > 0)
                                        {
                                            strDiscType = oRSDiscount.Fields.Item("U_DiscType").Value.ToString();
                                            dblDiscRate = Convert.ToDouble(oRSDiscount.Fields.Item("U_DiscRate").Value.ToString());

                                            strExpCode = oRSDiscount.Fields.Item("U_FrghtCode").Value.ToString();

                                            strBasePrice = oRSDiscount.Fields.Item("U_BasePrice").Value.ToString();

                                            if (strBasePrice == "G")
                                                dblDiscPItem5 = dblPriceAfVat * dblDiscRate / 100;
                                            else
                                                dblDiscPItem5 = dblNetDiscAmt / dblQuantity * dblDiscRate / 100;

                                            dblTotDisc5 = dblQuantity * dblDiscPItem5;
                                            dblNetDiscAmt5 = dblNetDiscAmt - dblTotDisc5;
                                            dblNetDiscAmt = dblNetDiscAmt4;


                                        }
                                        else
                                        {
                                            strDiscType = "";
                                            strExpCode = "";
                                            dblDiscRate = 0;
                                            dblDiscPItem5 = 0;
                                            dblTotDisc5 = 0;
                                            dblNetDiscAmt5 = dblNetDiscAmt - dblTotDisc5;
                                            dblNetDiscAmt = dblNetDiscAmt5;
                                        }
                                        if (dblTotDisc5 > 0)
                                        {

                                            oDocuments.Lines.UserFields.Fields.Item("U_DiscType5").Value = strDiscType;
                                            oDocuments.Lines.UserFields.Fields.Item("U_DiscRate5").Value = dblDiscRate;
                                            oDocuments.Lines.UserFields.Fields.Item("U_DiscPerItem5").Value = dblDiscPItem5;
                                            oDocuments.Lines.UserFields.Fields.Item("U_DiscAmt5").Value = dblTotDisc5;
                                            oDocuments.Lines.UserFields.Fields.Item("U_NetofDisc5").Value = dblNetDiscAmt5;

                                            oDRFreight = oDTFreight.Select("ExpCode = '" + strExpCode + "' ");

                                            if (oDRFreight.Length > 0)
                                            {
                                                intFRow = Convert.ToInt32(oDRFreight[0]["Row"]);
                                                oDTFreight.Rows[intFRow]["Amount"] = dblTotDisc5 +
                                                                                               Convert.ToDouble(oDTFreight.Rows[intFRow]["Amount"]);
                                            }
                                            else
                                            {
                                                oDTFreight.Rows.Add(intDRow,
                                                                              strExpCode,
                                                                              strVatGroup,
                                                                              dblTotDisc5);

                                                intDRow++;
                                            }
                                        }

                                        #endregion
                                    }
                                    else
                                    {

                                        strQuery = string.Format("SELECT CRD1.\"Address\", CRD1.\"U_Whscode\", OCRD.\"ECVatGroup\" FROM CRD1 INNER JOIN OCRD ON CRD1.\"CardCode\" =  OCRD.\"CardCode\" WHERE CRD1.\"CardCode\" = '{0}' AND CRD1.\"Address\"= '{1}'", strCardCode, strShipTo);

                                        oRecordset = null;
                                        oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        oRecordset.DoQuery(strQuery);

                                        if (oRecordset.RecordCount == 0)
                                        {
                                            //to do raise error
                                            GlobalVariable.intErrNum = GlobalVariable.oCompany.GetLastErrorCode();
                                            GlobalVariable.strErrMsg = GlobalVariable.oCompany.GetLastErrorDescription();

                                            strStatus = "E";
                                            strMsgBod = string.Format("Error Posting {0}, Invalid ShipTo Address " + strShipTo + " ShipTo Address is Not found in database ", GlobalVariable.strFileName);

                                            SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), strMsgBod);

                                            return false;
                                        }


                                        strQuery = string.Format("SELECT OITM.\"ItemCode\", OITM.\"ItemName\" FROM OITM INNER JOIN OBCD ON OITM.\"ItemCode\" = OBCD.\"ItemCode\" WHERE OBCD.\"BcdCode\" = '{0}'", strCodeBars);

                                        oRecordset = null;
                                        oRecordset = (SAPbobsCOM.Recordset)GlobalVariable.oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                        oRecordset.DoQuery(strQuery);

                                        if (oRecordset.RecordCount == 0)
                                        {
                                            //to do raise error
                                            GlobalVariable.intErrNum = GlobalVariable.oCompany.GetLastErrorCode();
                                            GlobalVariable.strErrMsg = GlobalVariable.oCompany.GetLastErrorDescription();

                                            strStatus = "E";
                                            strMsgBod = string.Format("Error Posting {0}, Invalid Barcode No." + strCodeBars + " Barcode Not found in database ", GlobalVariable.strFileName);

                                            SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), strMsgBod);

                                            return false;
                                        }
                                    }

                                }

                                intPONum++;

                                if (oDTFreight.Rows.Count > 0)
                                {
                                    for (int intRow = 0; intRow < oDTFreight.Rows.Count; intRow++)
                                    {

                                        dblFAmount = Convert.ToDouble(oDTFreight.Rows[intRow]["Amount"]);
                                        if (dblFAmount > 0)
                                        {

                                            oDocuments.Expenses.ExpenseCode = Convert.ToInt32(oDTFreight.Rows[intRow]["ExpCode"].ToString());
                                            oDocuments.Expenses.LineGross = dblFAmount * -1;
                                            oDocuments.Expenses.VatGroup = oDTFreight.Rows[intRow]["VatGroup"].ToString();
                                            oDocuments.Expenses.Add();

                                        }
                                    }

                                }

                                oDocuments.ShipToCode = strShipTo;

                                if (oDocuments.Add() != 0)
                                {
                                    GlobalVariable.intErrNum = GlobalVariable.oCompany.GetLastErrorCode();
                                    GlobalVariable.strErrMsg = GlobalVariable.oCompany.GetLastErrorDescription();

                                    strStatus = "E";
                                    strMsgBod = string.Format("Error Posting {0}.\rError Code: {1}\rDescription: {2} ", GlobalVariable.strFileName, GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                                    SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, "", "", dteStart, "E", GlobalVariable.intErrNum.ToString(), GlobalVariable.strErrMsg);

                                    return false;
                                }
                                else
                                {
                                    GlobalVariable.intErrNum = 0;
                                    GlobalVariable.strErrMsg = GlobalVariable.oCompany.GetLastErrorDescription();

                                    strPostDocNum = GlobalVariable.oCompany.GetNewObjectKey().ToString();

                                    strMsgBod = string.Format("Successfully Posted Sales Order from {0}. Posted Document Number: {1} ", GlobalVariable.strFileName, strPostDocNum);

                                    SystemFunction.transHandler("Import", strTransType, GlobalVariable.intObjType.ToString(), GlobalVariable.strFileName, GlobalVariable.oCompany.GetNewObjectKey(), strPostDocNum, dteStart, "S", "0", strMsgBod);

                                    strStatus = "S";
                                }
                            }

                        }

                        if (GlobalVariable.oCompany.InTransaction)
                            GlobalVariable.oCompany.EndTransaction(BoWfTransOpt.wf_Commit);
                    }
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
        private static bool importDIAPIPostDocumentFExcel(string strFile)
        {
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
        private static bool importXMLData(string strFile)
        {
            string strOrderId, strBarCode, strShipTo = "";

            int intRowX, intRowL;

            double dblQty;

            DateTime dteOrder = Convert.ToDateTime("01/01/1900");

            XmlDocument xmlDoc = new XmlDocument();
            XmlNodeList xmlNodLst;

            DataRow[] oDRXML;

            try
            {
                initXMLDT();

                xmlDoc.Load(strFile);

                strOrderId = xmlDoc.GetElementsByTagName("orderId").ToString();

                xmlNodLst = xmlDoc.SelectNodes("orders/orderLogisticalInformation/orderLogisticalDateGroup/requestedDeliveryDate/latestDate/dateAndTime");
                foreach (XmlNode xmlNod in xmlNodLst)
                {
                    dteOrder = Convert.ToDateTime(xmlNod.SelectSingleNode("date").InnerText);
                }

                xmlNodLst = xmlDoc.SelectNodes("orders/orderLogisticalInformation/shipToLogistics/shipTo");
                foreach (XmlNode xmlNod in xmlNodLst)
                {
                    strShipTo = xmlNod.SelectSingleNode("name").InnerText;
                }

                intRowX = 0;
                xmlNodLst = xmlDoc.SelectNodes("orders/orderLineItem/tradeItemId");
                foreach (XmlNode xmlNod in xmlNodLst)
                {

                    strBarCode = xmlNod.SelectSingleNode("gtin").InnerText;

                    oDTXML.Rows.Add(intRowX, strOrderId, dteOrder, strShipTo, strBarCode, 0.00);
                    intRowX++;
                }

                intRowX = 0;
                xmlNodLst = xmlDoc.SelectNodes("orders/orderLineItem/requestedQuantity");
                foreach (XmlNode xmlNod in xmlNodLst)
                {
                    dblQty = Convert.ToDouble(xmlNod.SelectSingleNode("value").InnerText);

                    oDRXML = oDTXML.Select(string.Format("Row = {0} ", intRowX));
                    if (oDRXML.Length > 0)
                    {
                        intRowL = Convert.ToInt32(oDRXML[0]["Row"]);
                        oDTXML.Rows[intRowL]["Quantity"] = dblQty;
                    }
                    intRowX++;
                }

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

            bool blWithErr = false;

            DataTable oDTDate;
            DataRow[] oDRTrck;

            DateTime dteDoc, dteAct, dteCnt, dteCur;

            SAPbobsCOM.Recordset oRecordset;

            try
            {
                initTrackDT();

                strCurDat = DateTime.Today.ToString("MM/dd/yyyy");

                dteCur = GlobalFunction.getDateTime(strCurDat.Replace("/", ""), "MMDDYYYY", "DD/MM/YYYY");

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
                        
                        dteDoc = GlobalFunction.getDateTime(strDocDate.Replace("/", ""), "MMDDYYYY", "DD/MM/YYYY");

                        if (strStatus != "C" && strCancel == "N")
                        {
                            if (!(string.IsNullOrEmpty(oDTDate.Rows[intRow][1].ToString())))
                            {
                                strActlDte = Convert.ToDateTime(oDTDate.Rows[intRow][1].ToString()).ToString("MM/dd/yyyy");

                                dteAct = GlobalFunction.getDateTime(strActlDte.Replace("/", ""), "MMDDYYYY", "DD/MM/YYYY");

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

                            dteCnt = GlobalFunction.getDateTime(strCntrDte.Replace("/", ""), "MMDDYYYY", "DD/MM/YYYY");

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
        private static void initFreightDT()
        {
            oDTFreight = new DataTable("Freight");
            oDTFreight.Columns.Add("Row", typeof(System.Int32));
            oDTFreight.Columns.Add("ExpCode", typeof(System.String));
            oDTFreight.Columns.Add("VatGroup", typeof(System.String));
            oDTFreight.Columns.Add("Amount", typeof(System.Double));
            oDTFreight.Clear();
        }
        private static void initXMLDT()
        {
            oDTXML = new DataTable("XMLData");
            oDTXML.Columns.Add("Row", typeof(System.Int32));
            oDTXML.Columns.Add("OrderID", typeof(System.String));
            oDTXML.Columns.Add("OrderDate", typeof(System.String));
            oDTXML.Columns.Add("ShipTo", typeof(System.String));
            oDTXML.Columns.Add("BarCode", typeof(System.String));
            oDTXML.Columns.Add("Quantity", typeof(System.Double));
            oDTXML.Clear();

        }
    }
}
