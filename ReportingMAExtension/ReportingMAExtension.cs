
using System;
using Microsoft.MetadirectoryServices;
using System.Text;

namespace Mms_ManagementAgent_ReportingMAExtension
{
    /// <summary>
    /// Summary description for MAExtensionObject.
    /// </summary>
    public class MAExtensionObject : IMASynchronization
    {
        public MAExtensionObject()
        {
            //
            // TODO: Add constructor logic here
            //
        }
        void IMASynchronization.Initialize()
        {
            //
            // TODO: write initialization code
            //
        }

        void IMASynchronization.Terminate()
        {
            //
            // TODO: write termination code
            //
        }

        bool IMASynchronization.ShouldProjectToMV(CSEntry csentry, out string MVObjectType)
        {
            //
            // TODO: Remove this throw statement if you implement this method
            //
            throw new EntryPointNotImplementedException();
        }

        DeprovisionAction IMASynchronization.Deprovision(CSEntry csentry)
        {
            //
            // TODO: Remove this throw statement if you implement this method
            //
            throw new EntryPointNotImplementedException();
        }

        bool IMASynchronization.FilterForDisconnection(CSEntry csentry)
        {
            //
            // TODO: write connector filter code
            //
            throw new EntryPointNotImplementedException();
        }

        void IMASynchronization.MapAttributesForJoin(string FlowRuleName, CSEntry csentry, ref ValueCollection values)
        {
            //
            // TODO: write join mapping code
            //
            throw new EntryPointNotImplementedException();
        }

        bool IMASynchronization.ResolveJoinSearch(string joinCriteriaName, CSEntry csentry, MVEntry[] rgmventry, out int imventry, ref string MVObjectType)
        {
            //
            // TODO: write join resolution code
            //
            throw new EntryPointNotImplementedException();
        }

        void IMASynchronization.MapAttributesForImport(string FlowRuleName, CSEntry csentry, MVEntry mventry)
        {
            //
            // TODO: write your import attribute flow code
            //
            throw new EntryPointNotImplementedException();
        }

        void IMASynchronization.MapAttributesForExport(string FlowRuleName, MVEntry mventry, CSEntry csentry)
        {
            //
            // export attribute flow code
            //

            //RF Connector added to retrieve UM Reference variables (Exchange)
            ConnectedMA RFADMA;
            int RFconnectors = 0;
            string strmsExchUMRecipientDialPlanLink, strmsExchUMTemplateLink;
            CSEntry csentryRF;

            //Code added to retrieve UM attributes
            RFADMA = mventry.ConnectedMAs["Resource Forest AD MA"];
            RFconnectors = RFADMA.Connectors.Count;

            switch (FlowRuleName)
            {
                case "cd.person:Lillynet_Enabled<-mv.person:Lillynet_Enabled,sAMAccountName":
                    // TODO: remove the following statement and add your scripted export attribute flow here
                    if (mventry["Lillynet_Enabled"].IsPresent && mventry["Lillynet_Enabled"].Value.ToUpper() == "Y")
                    {
                        csentry["Lillynet_Enabled"].BooleanValue = true;
                    }
                    else
                    {
                        csentry["Lillynet_Enabled"].BooleanValue = false;
                    }
                    break;

				////HCM Comments- QuickConnect MA retired. commenting below code.
				/*
                case "cd.person:QuickConnect_Enabled<-mv.person:quickConnectFlag,sAMAccountName":
                    // TODO: remove the following statement and add your scripted export attribute flow here
                    // TODO: remove the following statement and add your scripted export attribute flow here
                    if (mventry.ConnectedMAs["Quick Connect MA"].Connectors.Count > 0)
                    {
                        csentry["QuickConnect_Enabled"].BooleanValue = true;
                    }

                    else
                    {
                        csentry["QuickConnect_Enabled"].BooleanValue = false;
                    }
                    break;
				*/

                case "cd.person:AD_LAN_Enabled<-mv.person:afcreate_dt,sAMAccountName":
                    // TODO: remove the following statement and add your scripted export attribute flow here
                    // TODO: remove the following statement and add your scripted export attribute flow here
                    if (mventry["afcreate_dt"].IsPresent)
                    {
                        csentry["AD_LAN_Enabled"].BooleanValue = true;
                    }
                    else
                    {
                        csentry["AD_LAN_Enabled"].BooleanValue = false;
                    }
                    break;


                case "cd.person:Exchange_Enabled<-mv.person:msExchHideFromAddressList,msExchMailboxGuid,sAMAccountName":
                    // TODO: remove the following statement and add your scripted export attribute flow here
                    if (mventry["msExchMailboxGuid"].IsPresent)
                    {
                        if (mventry["msExchHideFromAddressList"].IsPresent && mventry["msExchHideFromAddressList"].BooleanValue == true)
                        {
                            csentry["Exchange_Enabled"].BooleanValue = false;
                        }
                        else
                        {
                            csentry["Exchange_Enabled"].BooleanValue = true;
                        }

                    }
                    else
                    {
                        csentry["Exchange_Enabled"].BooleanValue = false;
                    }
                    break;



                case "cd.person:Lync_Enabled<-mv.person:msRTCSIP-PrimaryUserAddress,msRTCSIP-UserEnabled,sAMAccountName":
                    // TODO: remove the following statement and add your scripted export attribute flow here

                    if (mventry["msRTCSIP-PrimaryUserAddress"].IsPresent && mventry["msRTCSIP-UserEnabled"].IsPresent && mventry["msRTCSIP-UserEnabled"].BooleanValue == true)
                    {
                        csentry["Lync_Enabled"].BooleanValue = true;
                    }
                    else
                    {
                        csentry["Lync_Enabled"].BooleanValue = false;
                    }
                    break;


                case "cd.person:Pwd_Sent<-mv.person:pwd_email_sentDate,sAMAccountName":
                    // TODO: remove the following statement and add your scripted export attribute flow here
                    if (mventry["pwd_email_sentDate"].IsPresent)
                    {
                        csentry["Pwd_Sent"].BooleanValue = true;
                    }
                    else
                    {
                        csentry["Pwd_Sent"].BooleanValue = false;
                    }
                    break;



                case "cd.person:ACC_DSBLD_FLG<-mv.person:sAMAccountName,userAccountControl":
                    // TODO: remove the following statement and add your scripted export attribute flow here
                    // TODO: remove the following statement and add your scripted export attribute flow here
                    if (mventry["userAccountControl"].IsPresent && mventry["userAccountControl"].IntegerValue == 514)
                    {

                        csentry["ACC_DSBLD_FLG"].BooleanValue = true;
                    }
                    else
                    {
                        csentry["ACC_DSBLD_FLG"].BooleanValue = false;
                    }
                    break;


                case "cd.person:EXCHNGE_CRT_DT<-mv.person:Exchnge_crt_dt,sAMAccountName":
                    // TODO: remove the following statement and add your scripted export attribute flow here

                    if (mventry["Exchnge_crt_dt"].IsPresent)
                    {
                        csentry["EXCHNGE_CRT_DT"].Value = Convert.ToDateTime(mventry["Exchnge_crt_dt"].Value).ToString("yyyy-MM-dd HH:mm:ss");

                    }
                    break;

                case "cd.person:LNC_CRT_DT<-mv.person:LyncCreateDate,sAMAccountName":
                    if (mventry["LyncCreateDate"].IsPresent)
                    {
                        csentry["LNC_CRT_DT"].Value = Convert.ToDateTime(mventry["LyncCreateDate"].Value).ToString("yyyy-MM-dd HH:mm:ss");

                    }
                    break;

                case "cd.person:DSBLD_DT<-mv.person:Dsbld_dt,sAMAccountName":

                    if (!mventry["Dsbld_dt"].IsPresent || mventry["Dsbld_dt"].Value == string.Empty)
                    {
                        csentry["DSBLD_DT"].Delete();
                    }
                    else
                    {

                        csentry["DSBLD_DT"].Value = mventry["Dsbld_dt"].Value;
                    }

                    break;

                //FIMReportingTool Archice Users New Flow - July 2020
                case "cd.person:OBJECT_SID<-mv.person:objectSid,sAMAccountName":                    

                    if (!mventry["objectSid"].IsPresent || mventry["objectSid"].Value == string.Empty)
                    {
                        csentry["OBJECT_SID"].Delete();
                    }
                    else
                    {
                        csentry["OBJECT_SID"].Value = ConvertByteToStringSid(mventry["objectSid"].IsPresent == false ? null : mventry["objectSid"].BinaryValue);
                    }

                    break;

                //FIMReportingTool Archice Users New Flow - July 2020
                case "cd.person:HOME_DIRECTORY<-mv.person:HomeDirectory,sAMAccountName":

                    if (!mventry["HomeDirectory"].IsPresent || mventry["HomeDirectory"].Value == string.Empty)
                    {
                        csentry["HOME_DIRECTORY"].Delete();
                    }
                    else
                    {
                        csentry["HOME_DIRECTORY"].Value = mventry["HomeDirectory"].Value.ToString();
                    }

                    break;

                //FIMReportingTool Archice Users New Flow - July 2020
                case "cd.person:DN<-mv.person:msDS-SourceObjectDN,sAMAccountName":

                    if (!mventry["msDS-SourceObjectDN"].IsPresent || mventry["msDS-SourceObjectDN"].Value == string.Empty)
                    {
                        csentry["DN"].Delete();
                    }
                    else
                    {
                        csentry["DN"].Value = mventry["msDS-SourceObjectDN"].Value.ToString();
                    }

                    break;

                //FIMReportingTool Archice Users New Flow - July 2020
                case "cd.person:SPRVSR_PRSNL_NBR<-mv.person:ManagerPrsnlNbr,sAMAccountName":

                    if (!mventry["ManagerPrsnlNbr"].IsPresent || mventry["ManagerPrsnlNbr"].Value == string.Empty)
                    {
                        csentry["SPRVSR_PRSNL_NBR"].Delete();
                    }
                    else
                    {
                        csentry["SPRVSR_PRSNL_NBR"].Value = mventry["ManagerPrsnlNbr"].Value.ToString();
                    }

                    break;

                //FIMReportingTool Archice Users New Flow - July 2020
                case "cd.person:MSEXCH_MAILBOX_GUID<-mv.person:msExchMailboxGuid,sAMAccountName":
                   
                    if (!mventry["msExchMailboxGuid"].IsPresent || mventry["msExchMailboxGuid"].Value == string.Empty)
                    {
                        csentry["MSEXCH_MAILBOX_GUID"].Delete();
                    }
                    else
                    {
                        csentry["MSEXCH_MAILBOX_GUID"].Value = ConvertByteToStringGUID(mventry["msExchMailboxGuid"].IsPresent == false ? null : mventry["msExchMailboxGuid"].BinaryValue);
                    }

                    break;

                //FIMReportingTool Archice Users New Flow - July 2020
                case "cd.person:MSEXCH_HOME_SRVR_NM<-mv.person:msExchHomeServerName,sAMAccountName":

                    if (!mventry["msExchHomeServerName"].IsPresent || mventry["msExchHomeServerName"].Value == string.Empty)
                    {
                        csentry["MSEXCH_HOME_SRVR_NM"].Delete();
                    }
                    else
                    {
                        csentry["MSEXCH_HOME_SRVR_NM"].Value = mventry["msExchHomeServerName"].Value.ToString();
                    }

                    break;

                //FIMReportingTool Archice Users New Flow - July 2020
                case "cd.person:MSEXCH_HOMEMDB<-mv.person:homeMDB,sAMAccountName":

                    if (!mventry["HomeMDB"].IsPresent || mventry["homeMDB"].Value == string.Empty)
                    {
                        csentry["MSEXCH_HOMEMDB"].Delete();
                    }
                    else
                    {
                        csentry["MSEXCH_HOMEMDB"].Value = mventry["homeMDB"].Value.ToString();
                    }

                    break;

                //FIMReportingTool Archice Users New Flow - July 2020
                case "cd.person:LEGACY_EXCH_DN<-mv.person:legacyExchangeDn,sAMAccountName":

                    if (!mventry["legacyExchangeDn"].IsPresent || mventry["legacyExchangeDn"].Value == string.Empty)
                    {
                        csentry["LEGACY_EXCH_DN"].Delete();
                    }
                    else
                    {
                        csentry["LEGACY_EXCH_DN"].Value = mventry["legacyExchangeDn"].Value.ToString();
                    }

                    break;


                //FIMReportingTool Archice Users New Flow - July 2020
                case "cd.person:PROXY_ADDRS<-mv.person:proxyAddresses,sAMAccountName":

                    if (!mventry["proxyAddresses"].IsPresent)
                    {
                        csentry["PROXY_ADDRS"].Delete();
                    }
                    else
                    {
                        csentry["PROXY_ADDRS"].Value = getProxyAddresses(mventry["proxyAddresses"].Values);
                    }

                    break; 
               

                //FIMReportingTool Archice Users New Flow - July 2020 - TBD
                case "cd.person:MSEXCH_UM_ENBLD_FLAGS<-mv.person:msExchUMEnabledFlags,sAMAccountName":

                    if (!mventry["msExchUMEnabledFlags"].IsPresent || mventry["msExchUMEnabledFlags"].Value == string.Empty)
                    {
                        csentry["MSEXCH_UM_ENBLD_FLAGS"].Delete();
                    }
                    else
                    {
                        csentry["MSEXCH_UM_ENBLD_FLAGS"].Value = mventry["msExchUMEnabledFlags"].Value.ToString();
                    }

                    break;


                default:
                    // TODO: remove the following statement and add your default script here
                    throw new EntryPointNotImplementedException();
            }
        }

        private string getProxyAddresses(ValueCollection objValCollection)
        {
            string strFinalString = "";
            int intTempCount = 1;
            foreach (Value addrElement in objValCollection)
            {
                if (intTempCount < objValCollection.Count)
                {
                    strFinalString = strFinalString + addrElement.ToString() + ",";
                }
                else
                {
                    strFinalString = strFinalString + addrElement.ToString();
                }
                intTempCount++;
            }
            return strFinalString;


        }


        /// <summary>
        /// takes the Byte array and returns the msexchangemailguid in LDP format to be updated in deprovision file
        /// </summary>
        /// <param name="strguid"></param>
        /// <returns></returns>
        /// 
        private string ConvertByteToStringGUID(Byte[] GUIDBytes)
        {
            StringBuilder strguid;
            try
            {
                strguid = new StringBuilder();
                int guidcount = GUIDBytes.Length;
                for (int i = 3; i >= 0; i--)
                {
                    string strtempguid = Convert.ToString(Convert.ToUInt32(GUIDBytes[i].ToString(), 10), 16);
                    if (strtempguid.Length < 2)
                        strtempguid = "0" + strtempguid;
                    strguid.Append(strtempguid);
                }

                strguid.Append("-");

                for (int i = 5; i >= 4; i--)
                {
                    string strtempguid = Convert.ToString(Convert.ToUInt32(GUIDBytes[i].ToString(), 10), 16);
                    if (strtempguid.Length < 2)
                        strtempguid = "0" + strtempguid;
                    strguid.Append(strtempguid);
                }

                strguid.Append("-");

                for (int i = 7; i >= 6; i--)
                {
                    string strtempguid = Convert.ToString(Convert.ToUInt32(GUIDBytes[i].ToString(), 10), 16);
                    if (strtempguid.Length < 2)
                        strtempguid = "0" + strtempguid;
                    strguid.Append(strtempguid);
                }

                strguid.Append("-");

                for (int i = 8; i <= 9; i++)
                {
                    string strtempguid = Convert.ToString(Convert.ToUInt32(GUIDBytes[i].ToString(), 10), 16);
                    if (strtempguid.Length < 2)
                        strtempguid = "0" + strtempguid;
                    strguid.Append(strtempguid);
                }

                strguid.Append("-");

                for (int i = 10; i <= 15; i++)
                {
                    string strtempguid = Convert.ToString(Convert.ToUInt32(GUIDBytes[i].ToString(), 10), 16);
                    if (strtempguid.Length < 2)
                        strtempguid = "0" + strtempguid;
                    strguid.Append(strtempguid);
                }

                return strguid.ToString();
            }
            catch
            {
                strguid = null;
                return "";
            }
            finally
            {
                strguid = null;
            }
        }

        /// <summary>
        /// takes the Byte array and returns integer equivalent in perticular format
        /// </summary>
        /// <param name="strSid"></param>
        /// <returns></returns>
        private string ConvertByteToStringSid(Byte[] sidBytes)
        {
            short sSubAuthorityCount = 0;
            StringBuilder strSid;

            try
            {
                // Add SID revision.
                strSid = new StringBuilder();
                strSid.Append("S-");
                strSid.Append(sidBytes[0].ToString());
                sSubAuthorityCount = Convert.ToInt16(sidBytes[1]);

                // Next six bytes are SID authority value.
                if (sidBytes[2] != 0 || sidBytes[3] != 0)
                {
                    string strAuth = String.Format("0x{0:2x}{1:2x}{2:2x}{3:2x}{4:2x}{5:2x}",
                                (Int16)sidBytes[2],
                                (Int16)sidBytes[3],
                                (Int16)sidBytes[4],
                                (Int16)sidBytes[5],
                                (Int16)sidBytes[6],
                                (Int16)sidBytes[7]);
                    strSid.Append("-");
                    strSid.Append(strAuth);
                }
                else
                {
                    Int64 iVal = (Int32)(sidBytes[7]) +
                            (Int32)(sidBytes[6] << 8) +
                            (Int32)(sidBytes[5] << 16) +
                            (Int32)(sidBytes[4] << 24);
                    strSid.Append("-");
                    strSid.Append(iVal.ToString());
                }
                // Get sub authority count...
                int idxAuth = 0;
                int intCount = 0;
                for (int i = 0; i < sSubAuthorityCount; i++)
                {
                    idxAuth = 8 + i * 4;
                    intCount = intCount + 1;
                    UInt32 iSubAuth = BitConverter.ToUInt32(sidBytes, idxAuth);
                    strSid.Append("-");
                    strSid.Append(iSubAuth.ToString());
                }
                return strSid.ToString();
            }
            catch
            {
                strSid = null;
                return "";
            }
            finally
            {
                strSid = null;
            }
        }
    }
}
