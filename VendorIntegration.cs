using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace Esker_Vendor_Integration
{
    public class VendorIntegration
    {
        public static void Execute(FTS00OVI Form, string SvrType, string Server, string LicSvr, string SQLLogin, string SQLPass, string SAPDBName, string SAPLogin, string SAPPass,
           string InputFilePath, string InputProcessedFilePath, string OutputFilePath)
        {
            string connString = "";
            string query = "";
            string ruid = "";
            int retcode = 0;

            /* Connect Diapi */

            SAPbobsCOM.Company oCom = new SAPbobsCOM.Company();

            if (SvrType == "MSSQL2005")
            {
                oCom.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2005;
            }
            else if (SvrType == "MSSQL2008")
            {
                oCom.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008;
            }
            else if (SvrType == "MSSQL2012")
            {
                oCom.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012;
            }
            else if (SvrType == "MSSQL2014")
            {
                oCom.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014;
            }
            else if (SvrType == "MSSQL2016")
            {
                oCom.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2016;
            }
            else if (SvrType == "MSSQL2017")
            {
                oCom.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2017;
            }
            else if (SvrType == "MSSQL2019")
            {
                oCom.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2019;
            }

            oCom.Server = Server;
            oCom.DbUserName = SQLLogin;
            oCom.DbPassword = SQLPass;
            oCom.language = SAPbobsCOM.BoSuppLangs.ln_English;
            oCom.LicenseServer = LicSvr;
            oCom.CompanyDB = SAPDBName;
            oCom.UserName = SAPLogin;
            oCom.Password = SAPPass;

            if (oCom.Connect() != 0)
            {
                string message = oCom.GetLastErrorDescription().ToString().Replace("'", "");
                Form.Log.AppendText("[Error] " + DateTime.Now.ToString() + " : " + message + Environment.NewLine);
                return;
            }
            else
            {
                Form.Log.AppendText("[Message] " + DateTime.Now.ToString() + " : " + SAPDBName + " Diapi successfully Connected!" + Environment.NewLine);
            }

            /* Process all files in folder */
            if (Directory.Exists(InputFilePath))
            {
                foreach (string file in Directory.EnumerateFiles(@"" + InputFilePath, "*.xml"))
                {
                    query = "SELECT MY_XML.value('(/VendorRegistration/@RUID)[1]', 'VARCHAR(50)') AS [RUID], " +
                        "MY_XML.Vendor.query('Company').value('.', 'VARCHAR(100)') AS [CompnyName], " +
                        "CASE WHEN MY_XML.Vendor.query('Z_VendorCategory').value('.', 'VARCHAR(50)') = 'Trade' THEN 'VT-' ELSE 'VG-' END + " +
                        "LEFT((SELECT " + SAPDBName + ".dbo.FT_fn_GetAlphaFromString (MY_XML.Vendor.query('Company').value('.', 'VARCHAR(100)'))),1) + " +
                        "CAST((SELECT * FROM " + SAPDBName + "..FTS_fn_GetVendorCodeLastNo (MY_XML.Vendor.query('Company').value('.', 'VARCHAR(100)'), MY_XML.Vendor.query('Z_VendorCategory').value('.', 'VARCHAR(50)') ) T0) AS NVARCHAR(30)) AS [CardCode], " +
                        "MY_XML.Vendor.query('Z_PaymentTerms').value('.', 'VARCHAR(50)') AS [PaymentTerms], " +
                        "MY_XML.Vendor.query('VendorNumber').value('.', 'VARCHAR(50)') AS [VendorNumber], " +
                        //"ISNULL(T1.GroupNum, 0) AS [SAPPaymentTerms], " +
                        "MY_XML.Vendor.query('Z_DefaultCurrency').value('.', 'VARCHAR(50)') AS [Currency], " +
                        "MY_XML.Vendor.query('Z_VendorCategory').value('.', 'VARCHAR(50)') AS [Category], " +
                        "MY_XML.Vendor.query('VendorRegistrationDUNSNumber').value('.', 'VARCHAR(50)') AS [CompanyRegNo], " +
                        "MY_XML.Vendor.query('Country').value('.', 'VARCHAR(50)') AS [Country], " +
                        "MY_XML.Vendor.query('MailSub').value('.', 'VARCHAR(100)') AS [Address1], " +
                        "MY_XML.Vendor.query('Street').value('.', 'VARCHAR(100)') AS [Address2], " +
                        "MY_XML.Vendor.query('City').value('.', 'VARCHAR(50)') AS [City]," +
                        "MY_XML.Vendor.query('Mail_State').value('.', 'VARCHAR(50)') AS [State], " +
                        "MY_XML.Vendor.query('Zip_Code').value('.', 'VARCHAR(50)')  AS [ZipCode], " +
                        "MY_XML.Vendor.query('Website').value('.', 'VARCHAR(50)')  AS [Website] " +
                        "FROM " +
                        "( " +
                        "   SELECT CAST(MY_XML AS xml) " +
                        "   FROM OPENROWSET(BULK '" + file + "', SINGLE_BLOB) AS T(MY_XML) " +
                        ") AS T(MY_XML) CROSS APPLY MY_XML.nodes('VendorRegistration') AS MY_XML(Vendor) " +
                       // "LEFT OUTER JOIN " + SAPDBName + "..OCTG T1 ON MY_XML.Vendor.query('Z_PaymentTerms').value('.', 'VARCHAR(50)') = T1.GroupCode " +
                        " " +
                        "SELECT MY_XML.value('(/VendorRegistration/@RUID)[1]', 'VARCHAR(50)') AS [RUID], " +
                        "MY_XML.Contact.query('Email').value('.', 'VARCHAR(100)') AS [Email], " +
                        "MY_XML.Contact.query('FirstName').value('.', 'VARCHAR(50)') AS [FirstName], " +
                        "MY_XML.Contact.query('LastName').value('.', 'VARCHAR(50)') AS [LastName] " +
                        "FROM " +
                        "( " +
                        "   SELECT CAST(MY_XML AS xml) " +
                        "   FROM OPENROWSET(BULK '" + file + "', SINGLE_BLOB) AS T(MY_XML) " +
                        ") AS T(MY_XML) CROSS APPLY MY_XML.nodes('/VendorRegistration/CompanyOfficersTable/item') AS MY_XML(Contact) " + 
                        " " +
                        "SELECT MY_XML.value('(/VendorRegistration/@RUID)[1]', 'VARCHAR(50)') AS [RUID], " +
                        "T1.BankCode,T1.CountryCod AS [BankCountry], " +
                        "MY_XML.Bank.query('BankAccountHolder').value('.', 'VARCHAR(100)') AS [BankAccountHolder], " +
                        "MY_XML.Bank.query('BankAccountNumber').value('.', 'VARCHAR(50)') AS [BankAccountNumber], " +
                        "MY_XML.Bank.query('BankCountry').value('.', 'VARCHAR(50)') AS [BankCountry], " +
                        "MY_XML.Bank.query('SWIFT_BICCode').value('.', 'VARCHAR(50)') AS [SWIFT_BICCode] " +
                        "FROM " +
                        "( " +
                        "   SELECT CAST(MY_XML AS xml) " +
                        "   FROM OPENROWSET(BULK '" + file +"', SINGLE_BLOB) AS T(MY_XML) " +
                        ") AS T(MY_XML) CROSS APPLY MY_XML.nodes('/VendorRegistration/CompanyBankAccountsTable/item') AS MY_XML(Bank) " +
                        "LEFT OUTER JOIN " + SAPDBName + "..ODSC T1 ON MY_XML.Bank.query('SWIFT_BICCode').value('.', 'VARCHAR(50)') = T1.SwiftNum ";

                    connString = @"Persist Security Info=True;Data Source=" + Server + ";Initial Catalog=" + SAPDBName + ";User ID=" + SQLLogin + ";Password=" + SQLPass + ";";

                    using (SqlConnection conn = new SqlConnection(connString))
                    {
                        if (conn.State == ConnectionState.Open)
                            conn.Close();

                        conn.Open();

                        using (SqlDataAdapter da = new SqlDataAdapter("", conn))
                        {
                            da.SelectCommand.CommandText = query;
                            da.SelectCommand.CommandTimeout = 0;
                            DataSet ds = new DataSet();
                            da.Fill(ds);

                            DataTable dtVendor = ds.Tables[0];
                            DataTable dtContact = ds.Tables[1];
                            DataTable dtBank = ds.Tables[2];

                            da.Dispose();

                            if (dtVendor.Rows.Count > 0)
                            {
                                SAPbobsCOM.BusinessPartners oBP = (SAPbobsCOM.BusinessPartners)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
                                try
                                {
                                    foreach (DataRow rowBP in dtVendor.Rows)
                                    {
                                        if (!oCom.InTransaction) oCom.StartTransaction();
                                        if (oBP == null) oBP = (SAPbobsCOM.BusinessPartners)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
                                        if (rowBP["VendorNumber"].ToString() != "")
                                        {
                                            oBP.GetByKey(rowBP["VendorNumber"].ToString());
                                            ruid = rowBP["RUID"].ToString();

                                            oBP.CardName = rowBP["CompnyName"].ToString();
                                            oBP.EmailAddress = rowBP["Website"].ToString();
                                            oBP.CompanyRegistrationNumber = rowBP["CompanyRegNo"].ToString();

                                            oBP.UserFields.Fields.Item("U_EskerNo").Value = ruid;
                                            if (oBP.Addresses.Count == 0)
                                            {
                                                if (rowBP["CompnyName"].ToString().Length > 50)
                                                    oBP.Addresses.AddressName = rowBP["CompnyName"].ToString().Substring(0, 50);
                                                else
                                                    oBP.Addresses.AddressName = rowBP["CompnyName"].ToString();
                                                oBP.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_BillTo;
                                                oBP.Addresses.Street = rowBP["Address1"].ToString();
                                                oBP.Addresses.StreetNo = rowBP["Address2"].ToString();
                                                oBP.Addresses.Block = rowBP["City"].ToString();
                                                oBP.Addresses.Country = rowBP["Country"].ToString();
                                                //oBP.Addresses.City = rowBP["City"].ToString();
                                                oBP.Addresses.ZipCode = rowBP["ZipCode"].ToString();
                                            }
                                            else
                                            {
                                                for (int i =0; i < oBP.Addresses.Count; i++)
                                                {
                                                    oBP.Addresses.SetCurrentLine(i);
                                                    if (oBP.Addresses.AddressType == SAPbobsCOM.BoAddressType.bo_BillTo)
                                                    {
                                                        oBP.Addresses.Street = rowBP["Address1"].ToString();
                                                        oBP.Addresses.StreetNo = rowBP["Address2"].ToString();
                                                        oBP.Addresses.Block = rowBP["City"].ToString();
                                                        oBP.Addresses.Country = rowBP["Country"].ToString();
                                                        //oBP.Addresses.City = rowBP["City"].ToString();
                                                        oBP.Addresses.ZipCode = rowBP["ZipCode"].ToString();
                                                        break;
                                                    }
                                                }
                                            }

                                            DataRow[] drs1 = dtContact.Select("RUID = '" + ruid + "'");
                                            bool exist = false;
                                            if (drs1.Length > 0)
                                            {
                                                foreach (DataRow rowContact in drs1)
                                                {
                                                    for (int i = 0; i < oBP.ContactEmployees.Count; i ++)
                                                    {
                                                        oBP.ContactEmployees.SetCurrentLine(i);
                                                        if (rowContact["FirstName"].ToString() + " " + rowContact["LastName"].ToString() == oBP.ContactEmployees.Name)
                                                        {
                                                            oBP.ContactEmployees.FirstName = rowContact["FirstName"].ToString();
                                                            oBP.ContactEmployees.LastName = rowContact["LastName"].ToString();
                                                            oBP.ContactEmployees.E_Mail = rowContact["Email"].ToString();
                                                            exist = true;
                                                            break;
                                                        }
                                                        else
                                                        {
                                                            exist = false;
                                                        }
                                                    }
                                                    if (exist == false)
                                                    {
                                                        if (oBP.ContactEmployees.Count > 0) 
                                                        {
                                                            oBP.ContactEmployees.Add();
                                                            oBP.ContactEmployees.SetCurrentLine(oBP.ContactEmployees.Count - 1);
                                                        }
                                                        oBP.ContactEmployees.Name = rowContact["FirstName"].ToString() + " " + rowContact["LastName"].ToString();
                                                        oBP.ContactEmployees.FirstName = rowContact["FirstName"].ToString();
                                                        oBP.ContactEmployees.LastName = rowContact["LastName"].ToString();
                                                        oBP.ContactEmployees.E_Mail = rowContact["Email"].ToString();
                                                    }
                                                }
                                            }

                                            DataRow[] drs2 = dtBank.Select("RUID = '" + ruid + "'");
                                            exist = false;
                                            if (drs2.Length > 0)
                                            {
                                                foreach (DataRow rowBank in drs2)
                                                {
                                                    if (rowBank["BankCode"].ToString() != "")
                                                    {
                                                        for (int i = 0; i < oBP.BPBankAccounts.Count; i++)
                                                        {
                                                            oBP.BPBankAccounts.SetCurrentLine(i);
                                                            if (oBP.BPBankAccounts.AccountNo == rowBank["BankAccountNumber"].ToString() && oBP.BPBankAccounts.BankCode == rowBank["BankCode"].ToString())
                                                            {
                                                                oBP.BPBankAccounts.AccountNo = rowBank["BankAccountNumber"].ToString();
                                                                oBP.BPBankAccounts.BICSwiftCode = rowBank["SWIFT_BICCode"].ToString();
                                                                oBP.BPBankAccounts.AccountName = rowBank["BankAccountHolder"].ToString();
                                                                exist = true;
                                                                break;
                                                            }
                                                            else
                                                            {
                                                                exist = false;
                                                            }
                                                        }
                                                        if (exist == false)
                                                        {
                                                            if (oBP.BPBankAccounts.Count > 0)
                                                            {
                                                                oBP.BPBankAccounts.Add();
                                                                oBP.BPBankAccounts.SetCurrentLine(oBP.BPBankAccounts.Count - 1);
                                                            }
                                                            oBP.BPBankAccounts.BankCode = rowBank["BankCode"].ToString();
                                                            oBP.BPBankAccounts.AccountNo = rowBank["BankAccountNumber"].ToString();
                                                            oBP.BPBankAccounts.BICSwiftCode = rowBank["SWIFT_BICCode"].ToString();
                                                            oBP.BPBankAccounts.AccountName = rowBank["BankAccountHolder"].ToString();
                                                        }
                                                    }
                                                }
                                            }
                                            retcode = oBP.Update();
                                        }
                                        else
                                        {
                                            ruid = rowBP["RUID"].ToString();

                                            oBP.CardCode = rowBP["CardCode"].ToString();
                                            oBP.CardName = rowBP["CompnyName"].ToString();
                                            oBP.CardType = SAPbobsCOM.BoCardTypes.cSupplier;
                                            if (rowBP["Category"].ToString() == "Trade")
                                                oBP.GroupCode = 111;
                                            else
                                                oBP.GroupCode = 112;

                                            oBP.PayTermsGrpCode = int.Parse(rowBP["PaymentTerms"].ToString());
                                            oBP.Currency = rowBP["Currency"].ToString();
                                            oBP.EmailAddress = rowBP["Website"].ToString();
                                            oBP.CompanyRegistrationNumber = rowBP["CompanyRegNo"].ToString();

                                            if (rowBP["Category"].ToString() == "Trade")
                                                if (rowBP["Currency"].ToString() == "MYR")
                                                    oBP.DebitorAccount = "30010";
                                                else
                                                    oBP.DebitorAccount = "30020";
                                            else
                                                oBP.DebitorAccount = "30030";

                                            oBP.UserFields.Fields.Item("U_EskerNo").Value = ruid;

                                            if (rowBP["CompnyName"].ToString().Length > 50)
                                                oBP.Addresses.AddressName = rowBP["CompnyName"].ToString().Substring(0, 50);
                                            else
                                                oBP.Addresses.AddressName = rowBP["CompnyName"].ToString();
                                            oBP.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_BillTo;
                                            oBP.Addresses.Street = rowBP["Address1"].ToString();
                                            oBP.Addresses.StreetNo = rowBP["Address2"].ToString();
                                            oBP.Addresses.Block = rowBP["City"].ToString();
                                            oBP.Addresses.Country = rowBP["Country"].ToString();
                                            //oBP.Addresses.City = rowBP["City"].ToString();
                                            oBP.Addresses.ZipCode = rowBP["ZipCode"].ToString();

                                            DataRow[] drs1 = dtContact.Select("RUID = '" + ruid + "'");
                                            if (drs1.Length > 0)
                                            {
                                                foreach (DataRow rowContact in drs1)
                                                {
                                                    oBP.ContactEmployees.Name = rowContact["FirstName"].ToString() + " " + rowContact["LastName"].ToString();
                                                    oBP.ContactEmployees.FirstName = rowContact["FirstName"].ToString();
                                                    oBP.ContactEmployees.LastName = rowContact["LastName"].ToString();
                                                    oBP.ContactEmployees.E_Mail = rowContact["Email"].ToString();

                                                    oBP.ContactEmployees.Add();
                                                }
                                            }

                                            DataRow[] drs2 = dtBank.Select("RUID = '" + ruid + "'");
                                            if (drs2.Length > 0)
                                            {

                                                foreach (DataRow rowBank in drs2)
                                                {
                                                    if (rowBank["BankCode"].ToString() != "")
                                                    {
                                                        oBP.BPBankAccounts.Country = rowBank["BankCountry"].ToString();
                                                        oBP.BPBankAccounts.BankCode = rowBank["BankCode"].ToString();
                                                        oBP.BPBankAccounts.AccountNo = rowBank["BankAccountNumber"].ToString();
                                                        oBP.BPBankAccounts.BICSwiftCode = rowBank["SWIFT_BICCode"].ToString();
                                                        oBP.BPBankAccounts.AccountName = rowBank["BankAccountHolder"].ToString();
                                                        oBP.BPBankAccounts.Add();
                                                    }
                                                }
                                            }

                                            retcode = oBP.Add();
                                        }
                                        if (retcode != 0)
                                        {
                                            int errcode = 0;
                                            string errmsg = null;
                                            oCom.GetLastError(out errcode, out errmsg);
                                            if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                            Form.Log.AppendText("[Error] " + DateTime.Now.ToString() + " : " + ruid + " - " + errmsg + Environment.NewLine);

                                            writetoXML(ruid, "Failed", errmsg, OutputFilePath, Path.GetFileNameWithoutExtension(file));
                                        }
                                        else
                                        {
                                            if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                                            Form.Log.AppendText("[Message] " + DateTime.Now.ToString() + " : " + ruid + " - " + " successfully Added!" + Environment.NewLine);

                                            writetoXML(ruid, "Success", oCom.GetNewObjectKey(), OutputFilePath, Path.GetFileNameWithoutExtension(file));
                                        }
                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oBP);
                                        oBP = null;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                    Form.Log.AppendText("[Error] " + DateTime.Now.ToString() + " : " + ruid + " - " + ex.Message + Environment.NewLine);

                                    writetoXML(ruid, "Failed", ex.Message, OutputFilePath, Path.GetFileNameWithoutExtension(file));
                                }
                                oBP = null;
                            }
                            dtVendor.Dispose();
                            dtContact.Dispose();
                            dtBank.Dispose();
                        }
                        conn.Close();
                    }
                    /* Move File to Processed */
                    if (File.Exists(InputProcessedFilePath + Path.GetFileName(file)))
                    {
                        File.Delete(InputProcessedFilePath + Path.GetFileName(file));
                        File.Move(file, InputProcessedFilePath + Path.GetFileName(file));
                    }
                    else
                        File.Move(file, InputProcessedFilePath + Path.GetFileName(file));
                }
            }
        }

        public static void writetoXML(string ruid, string result, string resultMsg, string OutputFilePath, string fileName)
        {
            XmlTextWriter xmlWriter = new XmlTextWriter(@"" + OutputFilePath + fileName + "_" + DateTime.Now.ToString("yyyyMMddhhmmss") + ".xml", Encoding.UTF8);

            xmlWriter.Formatting = Formatting.Indented;

            xmlWriter.WriteStartDocument();

            xmlWriter.WriteStartElement("ERPAck");

            xmlWriter.WriteElementString("EskerVendorRegistrationID", ruid);
            if (result == "Success")
                xmlWriter.WriteElementString("ERPVendorNumber", resultMsg);
            else
                xmlWriter.WriteElementString("ERPPostingError", resultMsg);

            xmlWriter.WriteEndElement();
            xmlWriter.WriteEndDocument();
            xmlWriter.Flush();
            xmlWriter.Close();
        }
    }
}
