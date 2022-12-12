
using System;
using Microsoft.MetadirectoryServices;
using System.Collections;
using System.Xml;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace Mms_Metaverse
{
    /// <summary>
    /// Summary description for MVExtensionObject.
    /// </summary>
    public class MVExtensionObject : IMVSynchronization
    {
        XmlNode rnode;
        XmlNode node;
        string version, etypeNodeValue, timetolive, Logslocation, LogsSwitch;
        StreamWriter objStreamWriter;

        public MVExtensionObject()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        void IMVSynchronization.Initialize()
        {
            //
            // Initialize config details from xml file
            //
            try
            {
                const string XML_CONFIG_FILE = @"\rules-config.xml";
                XmlDocument config = new XmlDocument();
                string dir = Utils.ExtensionsDirectory;
                config.Load(dir + XML_CONFIG_FILE);

                rnode = config.SelectSingleNode("rules-extension-properties");
                node = rnode.SelectSingleNode("environment");
                string env = node.InnerText;
                rnode = config.SelectSingleNode
                    ("rules-extension-properties/management-agents/" + env + "/ad-ma");
                XmlNode confignode = rnode.SelectSingleNode("siteconfigfile");
                //Provisioning Version
                node = rnode.SelectSingleNode("version");
                version = node.InnerText;
                //Get the timetolive(ttyl) value from config file
                node = rnode.SelectSingleNode("ttl");
                timetolive = node.InnerText;
                //get the Users to be provisioned from config                       
                node = rnode.SelectSingleNode("provision");
                etypeNodeValue = node.InnerText.ToUpper();
                //get the location for saving the log files
                node = rnode.SelectSingleNode("AFLogsloc");
                Logslocation = node.InnerText;
                //get the option for logging the file. When Y then logging functionality enabled.When N logging is disabled.
                node = rnode.SelectSingleNode("AFLogsSwitch");
                LogsSwitch = node.InnerText;
            }
            catch (NullReferenceException nre)
            {
                // If a tag does not exist in the xml, then the stopped-extension-dll 
                // error will be thrown.
                throw nre;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        void IMVSynchronization.Terminate()
        {
            //
            // TODO: Add termination logic here
            //
        }

        void IMVSynchronization.Provision(MVEntry mventry)
        {
            ///
            /// Provisoning for Deprovisioned MA
            ///

            ConnectedMA deprovisionedMA;
            ConnectedMA sadMA;
            CSEntry csentry;
            string cSEntryType = mventry.ObjectType;
            int connectors = 0;
            int sadconnectors = 0;
            string formattedDate;

            #region FIM Reporting Tool Code Retire
            /*
            //RF Connector added to retrieve UM Reference variables (Exchange)
            ConnectedMA RFADMA;
            int RFconnectors = 0;
            string strmsExchUMRecipientDialPlanLink,strmsExchUMTemplateLink ;
            CSEntry csentryRF;
            */
            #endregion
            //HCM Comments - Commenting the below code Since it is related to reitred AF Privileged account MAs
            #region HCM Code Retire
            //Connectors for privilege accounts
            //ConnectedMA AFPrivAccMADS;
            //int AFPrivAccDSconnectors = 0;
            // ConnectedMA AFPrivAccMANO;
            //int AFPrivAccNOconnectors = 0;
            // ConnectedMA AFPrivAccMADB;
            // int AFPrivAccDBconnectors = 0;
            #endregion

            #region FIM Reporting Tool Code Retire
            // strmsExchUMRecipientDialPlanLink = string.Empty;
            // strmsExchUMTemplateLink = string.Empty;
            #endregion

            try
            {
                switch (cSEntryType)
                {
                    case "person":
                        {
                            //Initializing MA objects and connector counts
                            deprovisionedMA = mventry.ConnectedMAs["Deprovisioned MA"];
                            connectors = deprovisionedMA.Connectors.Count;
                            sadMA = mventry.ConnectedMAs["Staging Area Database MA"];
                            sadconnectors = sadMA.Connectors.Count;

                            #region FIM Reporting Tool Code Retire
                            /*
                            //Code added to retrieve UM attributes
                            RFADMA = mventry.ConnectedMAs["Resource Forest AD MA"];
                            RFconnectors = RFADMA.Connectors.Count;
                            */
                            #endregion

                            //HCM Comments - Commenting the below code Since it is related to reitred AF Privileged account MAs
                            #region HCM Code Retire
                            //BOC
                            //Code added to retrieve Privileged account MA connector
                            //AFPrivAccMADS = mventry.ConnectedMAs["AF Privileged Account - DS MA"];
                            // AFPrivAccDSconnectors = AFPrivAccMADS.Connectors.Count;

                            //Code added to retrieve Privileged account MA connector
                            // AFPrivAccMANO = mventry.ConnectedMAs["AF Privileged Account - NO MA"];
                            // AFPrivAccNOconnectors = AFPrivAccMANO.Connectors.Count;

                            //Code added to retrieve Privileged account MA connector
                            // AFPrivAccMADB = mventry.ConnectedMAs["AF Privileged Account - DB MA"];
                            // AFPrivAccDBconnectors = AFPrivAccMADB.Connectors.Count;

                            // EOC

                            #endregion

                            //Deprovisioned MA has no connector
                            if (connectors == 0)
                            {
                                //Also SAD MA has no connector, then a SAD object is deleted
                                //Proceed to create a Deprovisioned MA connector
                                if (sadconnectors == 0)
                                {
                                    csentry = deprovisionedMA.Connectors.StartNewConnector("Person");
                                    csentry["EMPLOYEEID"].Value = mventry["employeeID"].Value;
                                    //Set the inital date in Format
                                    formattedDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                    csentry["DEPROVISIONEDDATE"].Value = formattedDate;
                                    csentry.CommitNewConnector();
                                }
                            }
                            //Deprovisioned MA has one connector
                            else if (connectors == 1)
                            {
                                //Get the connector object for deprovisioning
                                csentry = deprovisionedMA.Connectors.ByIndex[0];

                                //SAD MA is not there, then check for Time to Live
                                if (sadconnectors == 0)
                                {
                                    //TTL expired then call deprovision on all connected MA's
                                    if (AccountTTLExpired(csentry["DEPROVISIONEDDATE"].Value, timetolive))
                                    {
                                        if (LogsSwitch.ToString()  == "Y")
                                        {
                                            CheckAndCreateFile();
                                            string strtempobjestsid, strtempobjestguid, strtempmsexchmailboxguid; //exchange
                                            objStreamWriter = new StreamWriter(BuildFileName(), true);

                                            if (mventry["objectSid"].IsPresent)
                                            {
                                                strtempobjestsid = ConvertByteToStringSid(mventry["objectSid"].IsPresent == false ? null : mventry["objectSid"].BinaryValue);
                                            }
                                            else
                                            {
                                                strtempobjestsid = string.Empty;
                                            }

                                            if (mventry["msExchMailboxGuid"].IsPresent)
                                            {
                                                strtempmsexchmailboxguid = ConvertByteToStringGUID(mventry["msExchMailboxGuid"].IsPresent == false ? null : mventry["msExchMailboxGuid"].BinaryValue);
                                            }
                                            else
                                            {
                                                strtempmsexchmailboxguid = string.Empty;
                                            }
                                          
                                            string strtemp = "-------------------------------------------------------" + "\r\n";

                                            #region FIM Reporting Tool Code Retire
                                            //RF Connector added to retrieve UM Reference variables (Exchange)
                                            /* FIM Reporting Tool Updates - Retiring flow for below attributes to logs
                                            if (!(RFconnectors == 0))
                                            {
                                                csentryRF = RFADMA.Connectors.ByIndex[0];
                                                if ((csentryRF["msExchUMTemplateLink"].IsPresent))
                                                {
                                                    strmsExchUMTemplateLink = csentryRF["msExchUMTemplateLink"].Value;
                                                }
                                                if ((csentryRF["msExchUMRecipientDialPlanLink"].IsPresent))
                                                {
                                                    strmsExchUMRecipientDialPlanLink = csentryRF["msExchUMRecipientDialPlanLink"].Value;
                                                }
                                            }
                                            */
                                            #endregion

                                            //HCM Comments - replacing the EDS_Supervisor_Prsnl_Nbr with ManagerPrsnlNbr, since EDS MA is retiring.
                                            strtemp = strtemp
                                            + " SystemID:        " + (mventry["sAMAccountName"].IsPresent == false ? string.Empty : mventry["sAMAccountName"].Value)
                                            + "\r\n EmployeeID:      " + (mventry["employeeID"].IsPresent == false ? string.Empty : mventry["employeeID"].Value)
                                            + "\r\n DisplayName:     " + (mventry["displayName"].IsPresent == false ? string.Empty : mventry["displayName"].Value)
                                            + "\r\n EmployeeType:    " + (mventry["employeeType"].IsPresent == false ? string.Empty : mventry["employeeType"].Value)
                                            + "\r\n C:               " + (mventry["c"].IsPresent == false ? string.Empty : mventry["c"].Value)
                                            + "\r\n Personnel Area Code:  " + (mventry["personnel_area_cd"].IsPresent == false ? string.Empty : mventry["personnel_area_cd"].Value)
                                            + "\r\n Object SID:      " + strtempobjestsid
                                            + "\r\n Home Directory:  " + (mventry["HomeDirectory"].IsPresent == false ? string.Empty : mventry["HomeDirectory"].Value)
                                            + "\r\n DN:              " + (mventry["msDS-SourceObjectDN"].IsPresent == false ? string.Empty : mventry["msDS-SourceObjectDN"].Value)
                                            + "\r\n Supervisor Personal Number:  " + (mventry["ManagerPrsnlNbr"].IsPresent == false ? string.Empty : mventry["ManagerPrsnlNbr"].Value)//HCM Comments - replacing the EDS_Supervisor_Prsnl_Nbr with ManagerPrsnlNbr, since EDS MA is retiring.
                                            + "\r\n MailNickname:     " + (mventry["mailNickname"].IsPresent == false ? string.Empty : mventry["mailNickname"].Value)   //exchange
                                            + "\r\n msExchMailboxGuid:   " + strtempmsexchmailboxguid  //exchange
                                            + "\r\n msExchHomeServerName:     " + (mventry["msExchHomeServerName"].IsPresent == false ? string.Empty : mventry["msExchHomeServerName"].Value)   //exchange
                                            + "\r\n msExchHomeMDB:      " + (mventry["HomeMDB"].IsPresent == false ? string.Empty : mventry["HomeMDB"].Value)   //exchange
                                            + "\r\n legacyExchangeDn:    " + (mventry["legacyExchangeDn"].IsPresent == false ? string.Empty : mventry["legacyExchangeDn"].Value) //exchange
                                            + "\r\n proxyAddresses:     " + (mventry["proxyAddresses"].IsPresent == false ? string.Empty : getProxyAddresses(mventry["proxyAddresses"].Values)) //exchange
                                           // + "\r\n msExchUMTemplateLink:      " + (!(strmsExchUMTemplateLink == null) == false ? string.Empty : strmsExchUMTemplateLink)   //exchange --FIM Reporting Tool Code Retired
                                           // + "\r\n msExchUMRecipientDialPlanLink:      " + (!(strmsExchUMRecipientDialPlanLink == null) == false ? string.Empty : strmsExchUMRecipientDialPlanLink)   //exchange --FIM Reporting Tool Code Retired
                                            + "\r\n msExchUMEnabledFlags:      " + (mventry["msExchUMEnabledFlags"].IsPresent == false ? string.Empty : mventry["msExchUMEnabledFlags"].Value)   //exchange
                                            + "\r\n";
                                            strtemp = strtemp + "\r\n";
                                            //objStreamWriter.WriteLine(mventry["sAMAccountName"].Value, ",", mventry["employeeID"], mventry["displayName"],mventry["objectSid"], mventry ["employeeType"], mventry["c"], mventry["personnel_area_cd"], mventry["manager"]  );
                                            objStreamWriter.WriteLine(strtemp);
                                            objStreamWriter.Close();


                                            //HCM Comments - commneting below code for privileged acocunts , since we are retiring AF Privileged MAs
                                            #region HCM Retirement
                                            /*
                                            string strPrivAcc = string.Empty;
                                            CheckAndCreateFile();
                                            if (AFPrivAccDSconnectors > 0)
                                            {
                                                strPrivAcc = strPrivAcc
                                            + " SystemID:        " + (mventry["sAMAccountName"].IsPresent == false ? string.Empty : mventry["sAMAccountName"].Value.ToString() + "-DS")
                                            + "\r\n EmployeeID:      " + (mventry["employeeID"].IsPresent == false ? string.Empty : mventry["employeeID"].Value)
                                            + "\r\n DisplayName:     " + (mventry["displayName"].IsPresent == false ? string.Empty : mventry["displayName"].Value.ToString() + "-DS")
                                            + "\r\n EmployeeType:    " + (mventry["employeeType"].IsPresent == false ? string.Empty : mventry["employeeType"].Value)
                                             + "\r\n";
                                            }
                                            if (AFPrivAccNOconnectors > 0)
                                            {
                                                strPrivAcc = strPrivAcc
                                           + " SystemID:        " + (mventry["sAMAccountName"].IsPresent == false ? string.Empty : mventry["sAMAccountName"].Value.ToString() + "-NO")
                                           + "\r\n EmployeeID:      " + (mventry["employeeID"].IsPresent == false ? string.Empty : mventry["employeeID"].Value)
                                           + "\r\n DisplayName:     " + (mventry["displayName"].IsPresent == false ? string.Empty : mventry["displayName"].Value.ToString() + "-NO")
                                           + "\r\n EmployeeType:    " + (mventry["employeeType"].IsPresent == false ? string.Empty : mventry["employeeType"].Value)
                                            + "\r\n";
                                            }
                                            if (AFPrivAccDBconnectors > 0)
                                            {
                                                strPrivAcc = strPrivAcc
                                           + " SystemID:        " + (mventry["sAMAccountName"].IsPresent == false ? string.Empty : mventry["sAMAccountName"].Value.ToString() + "-DB")
                                           + "\r\n EmployeeID:      " + (mventry["employeeID"].IsPresent == false ? string.Empty : mventry["employeeID"].Value)
                                           + "\r\n DisplayName:     " + (mventry["displayName"].IsPresent == false ? string.Empty : mventry["displayName"].Value.ToString() + "-DB")
                                           + "\r\n EmployeeType:    " + (mventry["employeeType"].IsPresent == false ? string.Empty : mventry["employeeType"].Value)
                                            + "\r\n";
                                            }
                                            if (AFPrivAccDSconnectors > 0 || AFPrivAccNOconnectors > 0 || AFPrivAccDBconnectors > 0)
                                            {
                                                objStreamWriter = new StreamWriter(BuildDeprovPrivAccountFileName(), true);
                                                objStreamWriter.WriteLine(strPrivAcc);
                                                objStreamWriter.Close();
                                            }
                                            */
                                            #endregion
                                        }
                                        mventry.ConnectedMAs.DeprovisionAll();
                                    }
                                }
                                else
                                {
                                    //SAD MA Object connected - SAD Subscription added back the deleted row 
                                    //Deprovision just Deprovision MA object only
                                    csentry.Deprovision();
                                }
                            }
                            else
                            {
                                throw (new UnexpectedDataException("Multiple connectors in MA:" + deprovisionedMA.Connectors.Count.ToString()));
                            }
                        }
                        break;
                }

            }
            // Handle any exceptions
            catch (ObjectAlreadyExistsException)
            {
                // Ignore if the object already exists, join rules will join the existing object later
            }
            catch (AttributeNotPresentException)
            {
                // Ignore if the attribute on the mventry object is not available at this time
                // For example if employee type isn't present then for those users will be ignored, 
                // This exception is used instead of using isPresent() method.
            }
            //catch (NoSuchAttributeException)
            //{
            //    // Ignore if the attribute is not available at this time - used during development
            //}
            catch (Exception ex)
            {
                // All other exceptions re-throw to rollback the object synchronization transaction and 
                // report the error to the run history
                throw ex;
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
                    strFinalString = strFinalString + addrElement.ToString() ;
                }
                intTempCount++;
            }
            return strFinalString;
        
        
        }

        bool IMVSynchronization.ShouldDeleteFromMV(CSEntry csentry, MVEntry mventry)
        {
            //
            // TODO: Add MV deletion logic here
            //
            throw new EntryPointNotImplementedException();
        }
        public bool AccountTTLExpired(string TerminatedDate, string TTL)
        {
            // If the TerminatedDate and TimeToLive attributes contain values, then
            // add the attributes and compare to the current date. If the current date
            // is more than or equal to the TerminatesDate and ToLiveTime value,
            // the function returns true.
            if (TerminatedDate.Equals(""))
            {
                return (false);
            }
            if (TTL.Equals(""))
            {
                return (false);
            }
            try
            {
                //Convert TerminatedDate to a DateTime object
                DateTime StartTTLDate;
                StartTTLDate = Convert.ToDateTime(TerminatedDate);
                //Convert the TTL string to a double
                double DaysToTTL = Convert.ToDouble(TTL); //TTL
                DateTime TimeToLiveDate = new DateTime();                
                //Add DaysToTTL to StartTTLDate to get TimeToLiveDate
                TimeToLiveDate = StartTTLDate.AddDays(DaysToTTL);
                //Round the DateTime to starting of the Day
                DateTime CurrentDateTimeRounded = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
                DateTime TimeToLiveDateRounded = new DateTime(TimeToLiveDate.Year, TimeToLiveDate.Month, TimeToLiveDate.Day);
                if (CurrentDateTimeRounded > TimeToLiveDateRounded)
                {
                    return (true);
                }
            }
            catch
            {
                // Handle exceptions here.
            }
            return false;
        }

        public string BuildFileName()
        {
            string strFileName = (DateTime.Now).Year.ToString() + "-" + (((DateTime.Now).Month <= 9) ? "0" + (DateTime.Now).Month.ToString() : (DateTime.Now).Month.ToString()) + "-" + ((DateTime.Now).Day.ToString().Length == 1 ? ("0" + (DateTime.Now).Day.ToString()) : (DateTime.Now).Day.ToString()) + "_DEPROVISION_INFO.txt";
            return Logslocation + strFileName;
        }

        //HCM Comments - commneting below code for privileged acocunts , since we are retiring AF Privileged MAs
        #region HCM Code Retire
           /*
        public string BuildDeprovPrivAccountFileName()
        {
            string strFileName = (DateTime.Now).Year.ToString() + "-" + (((DateTime.Now).Month <= 9) ? "0" + (DateTime.Now).Month.ToString() : (DateTime.Now).Month.ToString()) + "-" + ((DateTime.Now).Day.ToString().Length == 1 ? ("0" + (DateTime.Now).Day.ToString()) : (DateTime.Now).Day.ToString()) + "_DEPROVISION_PRIVILEGEDACCOUNT_INFO.txt";
            return Logslocation + strFileName;
        }
        */
        #endregion

        public void CheckAndCreateFile()
        {
            string strFileName = BuildFileName();
            FileInfo objFileInfo;
            try
            {
                if (!(File.Exists(strFileName)))   // If file does not exists then create
                {
                    objFileInfo = new FileInfo(strFileName);
                }
            }
            catch
            {
                objFileInfo = null;
            }
            finally
            {
                objFileInfo = null;
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

                for (int i = 5; i >=4 ; i--)
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
    }
}