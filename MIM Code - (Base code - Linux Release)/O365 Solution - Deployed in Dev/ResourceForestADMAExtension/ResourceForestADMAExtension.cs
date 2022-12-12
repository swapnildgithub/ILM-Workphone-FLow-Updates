using System;
using System.Xml;
using Microsoft.MetadirectoryServices;
using System.Data;
using System.Collections;
using System.Diagnostics;
using System.DirectoryServices;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using CommonLayer_NameSpace;

namespace Mms_ManagementAgent_ResourceForestADMAExtension
{
    /// <summary>
    /// Extension for Resource Forest AD MA. It also maps attributes.
    /// </summary>
    public class MAExtensionObject : IMASynchronization
    {
        XmlNode rnode;
        XmlNode node;
        //NLW mini Release change
        string version, lcsflagoff, etypeNodeValue, publicOu, domain, upn, sipHomeServer, env, strexchangemailprefix, strlegacyExchangeDN, strAcceptedDomains, strAcceptedDomainsCN, strRBACPolicyLink, strbogusdomain, strOVAFlagSwitch;

        string archivingenabled, userlocationprofile, userinternetaccessenabled, userfederationenabled, useroptionflags, strUMMailboxPolicy, strO365List;
        string[] arrSIPDomainList;
        Hashtable afvalhash;
        Hashtable valhashERMADatabase;
        ArrayList arrproxyadd, userpolicy, impstrArrServerValues;
        string[] strArrServerValues; //Exchange Random
        StreamReader objStreamReader;
        string strskipmail;
        //Used for list of domains
        string strZone1Domains, strZone2Domains, strZone3Domains;
        string strLillyDomainSuffix, strElancoDomainSuffix, strAgspanDomainSuffix, strContractorEmailDomainPrefix, strRFDomain, strAllowableUPNDomains;//Start - CHG-CHG1178839 

        string strValidateExpression = string.Empty;

        //ERMA code BOC
        StreamReader objStreamReaderERMA;

        string strERMASystemIDs;
        string[] arrERMASystemIDs;


        string[] arrAcceptedDomain;
        ValueCollection vcAcceptedDomain = Utils.ValueCollection("initialvalue");

        ValueCollection vcRecipientPolicy = Utils.ValueCollection("initialvalue");

        ValueCollection vcSecureDomainWildCard = Utils.ValueCollection("initialvalue");
        ValueCollection vcSecureDomainWithoutWildCard = Utils.ValueCollection("initialvalue");

        //Used for list of secure email domains
        string strSecureEmailDomain;

        //Secure email array and valuecollection
        string[] arrSecureDomain;


        //Steam reader for Recipient policy
        StreamReader objStreamReaderRecPolicy;

        // Mini NLW Release
        string strExchangeLivedate = string.Empty;


        public MAExtensionObject()
        {
            //
            // TODO: Add constructor logic here
            //
        }
        void IMASynchronization.Initialize()
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
                afvalhash = LoadAFConfig(confignode.InnerText);

                //ERMA change
                valhashERMADatabase = LoadERMAConfig();
                //

                //Provisioning Version
                node = rnode.SelectSingleNode("version");
                version = node.InnerText;
                //get the Users to be set with LCS flag false
                node = rnode.SelectSingleNode("lcsflagoff");
                lcsflagoff = node.InnerText.ToUpper();
                //get the Users to be provisioned from config                       
                node = rnode.SelectSingleNode("provision");
                etypeNodeValue = node.InnerText.ToUpper();
                
                //Bogus target
                node = rnode.SelectSingleNode("bogusdomain");
                strbogusdomain = node.InnerText;
                //OVA Switch
                node = rnode.SelectSingleNode("OVAEnableFlag");
                strOVAFlagSwitch = node.InnerText;

                //Provisioning Version

                //R5
                arrproxyadd = LoadPAConfig(dir + XML_CONFIG_FILE);

                //Exchange email prefix (important for non-prd env)
                rnode = config.SelectSingleNode
                    ("rules-extension-properties/management-agents/" + env + "/Exchange");
                node = rnode.SelectSingleNode("exchangemailprefix");
                strexchangemailprefix = node.InnerText;

                rnode = config.SelectSingleNode
                    ("rules-extension-properties/management-agents/" + env + "/Exchange");
                node = rnode.SelectSingleNode("exchangemailprefix");
                strLillyDomainSuffix = node.InnerText;

                rnode = config.SelectSingleNode
                   ("rules-extension-properties/management-agents/" + env + "/Exchange");
                node = rnode.SelectSingleNode("domainPrefix");
                strContractorEmailDomainPrefix = node.InnerText;//Network as contractor email domain prefix




                rnode = config.SelectSingleNode
                    ("rules-extension-properties/management-agents/" + env + "/Exchange");
                node = rnode.SelectSingleNode("ElancoEmailSuffix");
                strElancoDomainSuffix = node.InnerText;

                rnode = config.SelectSingleNode
                    ("rules-extension-properties/management-agents/" + env + "/Exchange");
                node = rnode.SelectSingleNode("AgspanEmailSuffix");
                strAgspanDomainSuffix = node.InnerText;

                //Start - CHG-CHG1178839, added below lines to read AllowableUPNDomains Email suffix list dynamically from Rules-config file.
                rnode = config.SelectSingleNode
                    ("rules-extension-properties/management-agents/" + env + "/Exchange");
                node = rnode.SelectSingleNode("AllowableUPNDomains");
                strAllowableUPNDomains = node.InnerText;
                // END - CHG - CHG1178839

                rnode = config.SelectSingleNode
                   ("rules-extension-properties/management-agents/" + env + "/Exchange");
                node = rnode.SelectSingleNode("RFDomainUPN");
                strRFDomain = node.InnerText;

                node = rnode.SelectSingleNode("legacyExchangeDN");
                strlegacyExchangeDN = node.InnerText;

                node = rnode.SelectSingleNode("RBAC");
                strRBACPolicyLink = node.InnerText;

                //List of available domains in environment
                rnode = config.SelectSingleNode
                 ("rules-extension-properties/management-agents/" + env + "/AvailableDomains");
                //Zone 1
                node = rnode.SelectSingleNode("Zone1");
                strZone1Domains = node.InnerText;
                //Zone 2
                node = rnode.SelectSingleNode("Zone2");
                strZone2Domains = node.InnerText;
                //Zone 3
                node = rnode.SelectSingleNode("Zone3");
                strZone3Domains = node.InnerText;

                //BOC NLW mini Release
                rnode = config.SelectSingleNode
                                   ("rules-extension-properties/management-agents/" + env + "/Exchange");
                node = rnode.SelectSingleNode("ExchangeLiveDate");
                strExchangeLivedate = node.InnerText;
                //EOC NLW mini Release


                //string for regular expression
                strValidateExpression = "^smtp:.*@exchange[123]\\." + strexchangemailprefix.Replace(".", "\\.") + "$";

                //Read Recipient policies
                objStreamReaderRecPolicy = File.OpenText("C:\\Program Files\\Microsoft Forefront Identity Manager\\2010\\Synchronization Service\\Extensions\\RecipientPolicies.txt");
                string inputRecipientPolicy;
                string strRecipientPolicy = string.Empty;
                string[] arrRecipientPolicy;


                while ((inputRecipientPolicy = objStreamReaderRecPolicy.ReadLine()) != null)
                {
                    strRecipientPolicy = strRecipientPolicy + inputRecipientPolicy.Trim();
                }
                if ((strRecipientPolicy == "") || (strRecipientPolicy == null))
                    throw new Exception("Exchange Accepted Domain is empty");
                else
                {
                    arrRecipientPolicy = strRecipientPolicy.Split(';');
                }

                vcRecipientPolicy.Remove("initialvalue");

                for (int i = 0; i < arrRecipientPolicy.Length; i++)
                {
                    vcRecipientPolicy.Add(arrRecipientPolicy[i].ToString().ToLower());
                }

                StreamReader objStreamReader;
                string input;

                objStreamReader = File.OpenText("C:\\Program Files\\Microsoft Forefront Identity Manager\\2010\\Synchronization Service\\Extensions\\allowedDomainsCN.txt");

                while ((input = objStreamReader.ReadLine()) != null)
                {
                    strAcceptedDomains = strAcceptedDomains + input.Trim();
                }

                if ((strAcceptedDomains == "") || (strAcceptedDomains == null))
                    throw new Exception("Exchange Accepted Domain is empty");
                else
                {
                    arrAcceptedDomain = strAcceptedDomains.Split(';');
                }

                vcAcceptedDomain.Remove("initialvalue");

                for (int i = 0; i < arrAcceptedDomain.Length; i++)
                {
                    vcAcceptedDomain.Add(arrAcceptedDomain[i].ToString().ToLower());
                }

                //Secure email load
                //List of secure email domains
                StreamReader objStreamReaderSecureDomain;
                string inputSecureDomain;

                objStreamReaderSecureDomain = File.OpenText("C:\\Program Files\\Microsoft Forefront Identity Manager\\2010\\Synchronization Service\\Extensions\\SecureDomains.txt");

                while ((inputSecureDomain = objStreamReaderSecureDomain.ReadLine()) != null)
                {
                    strSecureEmailDomain = strSecureEmailDomain + inputSecureDomain;
                }

                vcSecureDomainWildCard.Remove("initialvalue");
                vcSecureDomainWithoutWildCard.Remove("initialvalue");

                if ((strSecureEmailDomain == "") || (strSecureEmailDomain == null))
                {
                    throw new Exception("Secure Email Domain is empty");
                }
                else
                {
                    arrSecureDomain = strSecureEmailDomain.Split(';');
                }
                //Add secure domains to the value collection
                for (int i = 0; i < arrSecureDomain.Length; i++)
                {
                    if (arrSecureDomain[i].ToString().TrimStart().StartsWith("."))
                    {
                        vcSecureDomainWildCard.Add(arrSecureDomain[i].ToString());
                    }
                    else
                    {
                        vcSecureDomainWithoutWildCard.Add(arrSecureDomain[i].ToString());
                    }
                }
                //

                StreamReader objStreamReaderSIPDomain;
                string inputSIPDomainList, strSIPDomainList;
                strSIPDomainList = string.Empty;

                objStreamReaderSIPDomain = File.OpenText("C:\\Program Files\\Microsoft Forefront Identity Manager\\2010\\Synchronization Service\\Extensions\\AllowedSIPDomains.txt");

                while ((inputSIPDomainList = objStreamReaderSIPDomain.ReadLine()) != null)
                {
                    strSIPDomainList = strSIPDomainList + ";" + inputSIPDomainList.TrimStart().TrimEnd();
                }
                if ((strSIPDomainList == string.Empty) || (strSIPDomainList == null))
                    //throw new Exception("Exchange Accepted Domain is empty");
                    arrSIPDomainList = null;
                else
                {
                    arrSIPDomainList = strSIPDomainList.Split(';');
                }


                //Members of PVP group
                string inputSystemID;
                objStreamReaderERMA = File.OpenText("C:\\Program Files\\Microsoft Forefront Identity Manager\\2010\\Synchronization Service\\Extensions\\ERMASystemIDs.txt");

                while ((inputSystemID = objStreamReaderERMA.ReadLine()) != null)
                {
                    strERMASystemIDs = strERMASystemIDs + inputSystemID.ToUpper();
                }
                //ERMA code EOC

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
            // TODO: write Deprovision code
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
            int sadconnectors = 0;
            ConnectedMA sadMA = mventry.ConnectedMAs["Staging Area Database MA"];
            sadconnectors = sadMA.Connectors.Count;

            if (FlowRuleName.StartsWith("ImportRule:"))
            {
                string[] strFlow = FlowRuleName.Split(':');
                if (csentry[strFlow[1]].IsPresent)
                    mventry[strFlow[2]].Value = csentry[strFlow[1]].Value;
                else
                    mventry[strFlow[2]].Delete();
            }
            else
            {
                switch (FlowRuleName)
                {
                    
                    // O365 bug fix - setting targetAddress in import flow for email address of inactive Users.
                    #region targetAddress
                    //TargetAddress
                    case "cd.user:description->mv.person:TargetAddress":                       

                        setUserVariables(mventry["personnel_area_cd"].Value, mventry["employeeType"].Value);
                        switch (CommonLayer.GetUserType(mventry, csentry, "RF", "no"))
                        {

                            case "MADS":
                                if (!CommonLayer.IsMigrationOverride(mventry, csentry, "RF"))
                                {
                                    mventry["TargetAddress"].Value = "SMTP:" + mventry["Internet_Style_Adrs"].Value.ToLower();
                                }
                                break;
                            case "LocalMailbox":
                                mventry["TargetAddress"].Delete();
                                break;
                            case "MailContact":
                               
                                //BOC NLW mini Release
                                mventry["TargetAddress"].Value = "SMTP:" + mventry["External_Email_Address"].Value.ToString().ToLower();
                                if (sadconnectors == 0)
                                {
                                    string strBogusemailAdd = mventry["sAMAccountName"].Value.ToUpper() + "@" + strbogusdomain;
                                    mventry["TargetAddress"].Value = "SMTP:" + strBogusemailAdd.ToLower();
                                }
                                break;
                            //EOC NLW mini Release

                            case "None":
                                mventry["TargetAddress"].Delete();
                                break;
                            case "RemoteMailbox":
                                //do nothing
                                break;
                        }
                        break;
                    #endregion

                    #region telephoneNumber
                    case "cd.user:<dn>,telephoneNumber->mv.person:telephoneNumber":
                        if (csentry["telephoneNumber"].IsPresent)
                            mventry["telephoneNumber"].Value = csentry["telephoneNumber"].Value;
                        else
                            mventry["telephoneNumber"].Delete();
                        break;
                    #endregion

                    #region IsContractorEmailCreated
                    case "cd.user:employeeType,msExchMailboxGuid,msExchRecipientTypeDetails,sAMAccountName->mv.person:IsContractorEmailCreated":
                        if(
                            csentry["employeeType"].Value.ToUpper() == "D" &&
                            csentry["msExchRecipientTypeDetails"].IsPresent &&
                            (csentry["msExchRecipientTypeDetails"].Value.ToString() == "2" || csentry["msExchRecipientTypeDetails"].Value.ToString() == "2147483648") &&
                            csentry["msExchMailboxGuid"].IsPresent
                            )
                        {
                            mventry["IsContractorEmailCreated"].Value = "Y";
                        }
                        break;
                    #endregion

                    #region homeMTA
                    case "cd.user:homeMTA->mv.person:homeMTA":
                        if (csentry["homeMTA"].IsPresent)
                        {
                            mventry["homeMTA"].Value = csentry["homeMTA"].Value;
                        }
                        break;
                    #endregion

                    #region EA14
                    //BOC Mini NLW Release
                    case "cd.user:<dn>,extensionAttribute14->mv.person:EA14":
                        if(csentry["extensionAttribute14"].IsPresent)
                        mventry["EA14"].Value = csentry["extensionAttribute14"].Value;
                        break;
                    #endregion

                    #region tnmsVerificationFlags
                    case "cd.user:msExchUMEnabledFlags,msExchUMRecipientDialPlanLink,msExchUMTemplateLink,msRTCSIP-OptionFlags,msRTCSIP-UserPolicies,msRTCSIP-UserPolicy,telephoneNumber->mv.person:tnmsVerificationFlags":
                        string tnmsVerificationFlags = "";
                        if (csentry["msExchUMRecipientDialPlanLink"].IsPresent)
                            tnmsVerificationFlags = "DialPlan=" + csentry["msExchUMRecipientDialPlanLink"] + "&";

                        if (csentry["msExchUMTemplateLink"].IsPresent)
                            tnmsVerificationFlags += "MailboxPolicy=" + csentry["msExchUMTemplateLink"].Value + "&";

                        if (csentry["msRTCSIP-UserPolicies"].IsPresent)
                        {
                            string policies = "";
                            foreach (Value policy in csentry["msRTCSIP-UserPolicies"].Values)
                            {
                                policies = policies + policy.ToString() + ";";
                            }
                            tnmsVerificationFlags += "UserPolicy=" + policies +"&";
                        }
                        else if (csentry["msRTCSIP-UserPolicy"].IsPresent)
                            tnmsVerificationFlags += "UserPolicy=" + csentry["msRTCSIP-UserPolicy"].Value +"&";

                        if (csentry["telephoneNumber"].IsPresent)
                            tnmsVerificationFlags += "LineURI=" + csentry["telephoneNumber"].Value + "&";

                        if (csentry["msRTCSIP-OptionFlags"].IsPresent)
                        tnmsVerificationFlags += "EVFlag=" + csentry["msRTCSIP-OptionFlags"].Value + "&";
                        mventry["tnmsVerificationFlags"].Value=tnmsVerificationFlags;
                        break;
                    #endregion

                    #region Exchnge_crt_dt
                    case "cd.user:homeMTA,msExchRecipientTypeDetails->mv.person:Exchnge_crt_dt":

                        if (csentry["homeMTA"].IsPresent && csentry["msExchRecipientTypeDetails"].IsPresent
                           && csentry["msExchRecipientTypeDetails"].Value == "2")// check for mailboxguid is present and msExchRecipientTypeDetails for mailbox.
                        {

                            if (mventry["rfcreate_dt"].IsPresent)
                            {
                                string strrfdate = mventry["rfcreate_dt"].Value;
                                int intyr = Convert.ToInt32(strrfdate.Substring(0, 4));
                                int intmnth = Convert.ToInt32(strrfdate.Substring(5, 2));
                                int intday = Convert.ToInt32(strrfdate.Substring(8, 2));
                                DateTime dtrfcreate_dt = new DateTime(intyr, intmnth, intday);

                                strrfdate = strExchangeLivedate;
                                intyr = Convert.ToInt32(strrfdate.Substring(0, 4));
                                intmnth = Convert.ToInt32(strrfdate.Substring(5, 2));
                                intday = Convert.ToInt32(strrfdate.Substring(8, 2));
                                DateTime dtExchangeLive_dt = new DateTime(intyr, intmnth, intday);

                                //check for mventry["Exchnge_crt_dt"] is not null, then set current date to mventry["Exchnge_crt_dt"]
                                int compareFlag = DateTime.Compare(dtrfcreate_dt, dtExchangeLive_dt);
                                if (compareFlag <= 0)
                                {
                                    mventry["Exchnge_crt_dt"].Value = dtExchangeLive_dt.ToString("yyyy-MM-dd HH:mm:ss");
                                }
                                else
                                {
                                    mventry["Exchnge_crt_dt"].Value = dtrfcreate_dt.ToString("yyyy-MM-dd HH:mm:ss");
                                }
                            }
                        }
                        break;
                    #endregion

                    #region LyncCreateDate
                    case "cd.user:msRTCSIP-PrimaryUserAddress,msRTCSIP-UserEnabled->mv.person:LyncCreateDate":
                        if (!mventry["LyncCreateDate"].IsPresent)
                        {
                            if (csentry["msRTCSIP-PrimaryUserAddress"].IsPresent && csentry["msRTCSIP-UserEnabled"].IsPresent
                                && csentry["msRTCSIP-UserEnabled"].BooleanValue == true)
                            {
                                mventry["LyncCreateDate"].Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                            }
                        }
                        break;
                    //EOC Mini NLW Release
                    #endregion

                    #region rfcreate_dt
                    case "cd.user:whenCreated->mv.person:rfcreate_dt":
                        string strDatetime = csentry["whenCreated"].Value.ToString();
                        int year = Convert.ToInt32(strDatetime.Substring(0, 4));
                        int month = Convert.ToInt32(strDatetime.Substring(4, 2));
                        int day = Convert.ToInt32(strDatetime.Substring(6, 2));
                        int hh = Convert.ToInt32(strDatetime.Substring(8, 2));
                        int mm = Convert.ToInt32(strDatetime.Substring(10, 2));
                        int dd = Convert.ToInt32(strDatetime.Substring(12, 2));
                        DateTime dt = new DateTime(year, month, day, hh, mm, dd, 123);
                        mventry["rfcreate_dt"].Value = dt.ToString("yyyy-MM-dd HH:mm:ss");
                        break;
                    #endregion

                    #region homeMDB
                    case "cd.user:homeMDB->mv.person:homeMDB":
                        string sName = csentry["homeMDB"].Value.ToString();
                        mventry["homeMDB"].Value = sName;
                        break;
                    #endregion

                    #region msExchHomeServerName
                    case "cd.user:destinationIndicator,msExchHomeServerName->mv.person:msExchHomeServerName":

                        if (csentry["msExchHomeServerName"].IsPresent)
                        {
                            mventry["msExchHomeServerName"].Value = csentry["msExchHomeServerName"].Value;

                            setUserVariables(mventry["personnel_area_cd"].Value, mventry["employeeType"].Value);
                            int verifyserver = 0;

                            for (int i = 0; i <= impstrArrServerValues.Count - 1; i++)
                            {
                                if (csentry["msExchHomeServerName"].Value.ToLower().Equals(impstrArrServerValues[i].ToString().ToLower()))
                                {
                                    verifyserver = 1;
                                    break;
                                }
                            }

                            if (verifyserver == 0)
                            {
                                string sSource, sLog, sEvent;

                                sSource = "MIIS RF MA";
                                sLog = "Application";
                                sEvent = mventry["sAMAccountName"].Value + " - msExchHomeServer value doesn't match PAC & employeeType combination specified within Config file";

                                if (!EventLog.SourceExists(sSource))
                                    EventLog.CreateEventSource(sSource, sLog);

                                EventLog.WriteEntry(sSource, sEvent, EventLogEntryType.Error, 8161);

                            }
                        }
                        else
                        {
                            mventry["msExchHomeServerName"].Delete();
                        }
                        break;
                    #endregion

                    #region proxyAddresses
                    case "cd.user:mail,mailNickname,msExchRecipientTypeDetails,proxyAddresses,sAMAccountName,targetAddress->mv.person:proxyAddresses":

                        string strUserType = CommonLayer.GetUserType(mventry, csentry, "RF", "no");
                        bool HasSipAddress = false;//Need to add sip address (default value)
                        string strSipAddress = string.Empty;
                        mventry["proxyAddresses"].Values = csentry["proxyAddresses"].Values;
                        ValueCollection vcproxyAddress = Utils.ValueCollection("initialValue");
                        vcproxyAddress.Remove("initialValue");

                        //Logic starts
                        if (strUserType == "None")
                        {
                            if (mventry["Exch_Trans_Cd"].Value.ToString() == "9")
                                mventry["proxyAddresses"].Values.Clear();
                        }
                        //Only add these values into proxyaddress when EAP and FIM updates in same cycle 
                        //("Export change is not re-imported" error)
                        else
                        {
                            bool IsExchangeCoexAddress = false;
                            #region Iterate proxyaddress
                            foreach (Value proxyElement in csentry["proxyAddresses"].Values)
                            {
                                IsExchangeCoexAddress = Regex.IsMatch(proxyElement.ToString(), strValidateExpression, RegexOptions.IgnoreCase);

                                //sip check
                                if (proxyElement.ToString().StartsWith("sip:"))
                                {
                                    if (strUserType == "LocalMailbox" || strUserType == "RemoteMailbox")
                                    {
                                        HasSipAddress = true;//Do nothing as it is already present
                                        vcproxyAddress.Add(proxyElement);
                                        strSipAddress = proxyElement.ToString();
                                    }
                                }
                                else if (
                                          !(proxyElement.ToString().StartsWith("smtp:")) ||
                                          (csentry["targetAddress"].IsPresent && proxyElement.ToString().ToLower() == csentry["targetAddress"].Value.ToString().ToLower()) ||
                                          (IsAcceptedDomain(proxyElement.ToString(), vcAcceptedDomain) == true && (strUserType == "LocalMailbox" || strUserType == "RemoteMailbox" || IsExchangeCoexAddress == false))
                                        )
                                {
                                    vcproxyAddress.Add(proxyElement);
                                }
                            }
                            #endregion

                            if (!ContainsIgnoreCase(vcproxyAddress, "smtp:" + mventry["sAMAccountName"].Value.ToString() + "@" + strexchangemailprefix))
                            {
                                vcproxyAddress.Add("smtp:" + mventry["sAMAccountName"].Value.ToString() + "@" + strexchangemailprefix);
                            }
                            ValueCollection vc = Utils.ValueCollection("initialValue");
                            vc.Remove("initialValue");
                            vc.Add(vcproxyAddress);

                            if ((strUserType == "LocalMailbox" || strUserType == "RemoteMailbox") && csentry["mail"].IsPresent)
                            {
                                if (HasSipAddress)
                                {
                                    //Remove SIP value from proxy Address
                                    vcproxyAddress.Remove(strSipAddress);
                                }
                                vcproxyAddress.Add("sip:" + csentry["mail"].Value.ToString());
                            }
                        }

                        if (csentry["mail"].IsPresent)
                        {
                            vcproxyAddress.Add("SMTP:" + csentry["mail"].Value.ToString());
                        }

                        //Clear the existing proxyaddress values 
                        mventry["proxyAddresses"].Values.Clear();
                        //Put the proxy address values with accepted domains
                        mventry["proxyAddresses"].Values.Add(vcproxyAddress);
                        break;
                    #endregion


                    #region mail
                    // O365 bug fix - added Description as a depenency to fire the below code on Delta Sync as well
                    case "cd.user:<dn>,description,mail,mailNickname,msExchPoliciesIncluded,proxyAddresses,sAMAccountName->mv.person:mail":
                        if (!csentry["mail"].IsPresent)
                        {//To clear mail if mail is not present in RF
                            mventry["mail"].Delete();
                        }
                        else
                        {
                            setUserVariables(mventry["personnel_area_cd"].Value, mventry["employeeType"].Value);
                            switch (CommonLayer.GetUserType(mventry, csentry, "RF", "no"))
                            {
                                case "LocalMailbox":
                                case "RemoteMailbox":
                                case "MADS":
                                    {
                                        mventry["mail"].Value = csentry["mail"].Value;
                                    }
                                    break;
                                case "None":
                                    mventry["mail"].Delete();
                                    break;
                                case "MailContact":

                                    if (csentry["mailnickname"].IsPresent)
                                    {
                                        bool blnmatchflag = false;
                                        string strExpression = "^" + csentry["mailnickname"].Value.ToString() + "[0-9]*$";
                                        //Iterate
                                        foreach (Value ExchPolicyElement in csentry["msExchPoliciesIncluded"].Values)//User Policy
                                        {
                                            foreach (Value addrRecipientPolicyElement in vcRecipientPolicy)//System Policy
                                            {
                                                if (addrRecipientPolicyElement.ToString().Contains(ExchPolicyElement.ToString().ToLower()))
                                                {
                                                    string[] EmailDomainstr;
                                                    EmailDomainstr = addrRecipientPolicyElement.ToString().Split('@');

                                                    foreach (Value prxyElement in csentry["proxyaddresses"].Values)//proxyaddress
                                                    {
                                                        if (prxyElement.ToString().ToLower().EndsWith("@" + EmailDomainstr[1]))
                                                        {
                                                            string[] strprxymail;
                                                            string[] strtempmail;

                                                            strprxymail = prxyElement.ToString().Split(':');

                                                            strtempmail = strprxymail[1].ToLower().Split('@');

                                                            if (!(strtempmail[0] == "") && (csentry["mailnickname"].IsPresent))
                                                            {
                                                                if (strtempmail[0].Equals(csentry["mailnickname"].Value.ToLower()))
                                                                {
                                                                    mventry["mail"].Value = strprxymail[1];
                                                                    blnmatchflag = true;
                                                                    break;
                                                                }
                                                                else if (
                                                                        strtempmail[0].StartsWith(csentry["mailnickname"].Value.ToLower()) &&
                                                                        Regex.IsMatch(csentry["mailnickname"].Value.ToString(), strExpression, RegexOptions.IgnoreCase)
                                                                        )
                                                                {
                                                                    mventry["mail"].Value = strprxymail[1];
                                                                    blnmatchflag = true;
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                                if (blnmatchflag)
                                                    break;
                                            }
                                            if (blnmatchflag)
                                            {
                                                blnmatchflag = false;
                                                break;
                                            }
                                        }
                                    }
                                    else
                                        mventry["mail"].Delete();

                                    break;
                            }
                        }

                        break;
                    #endregion

                    #region IsMigrationAccessGranted
                    case "cd.user:extensionAttribute14,sAMAccountName->mv.person:IsMigrationAccessGranted":

                        if (csentry["extensionAttribute14"].IsPresent)
                        {
                            mventry["IsMigrationAccessGranted"].BooleanValue = true;
                        }

                        break;
                    # endregion
                    

                    default:
                        throw new EntryPointNotImplementedException();
                }
            }
        }

        void IMASynchronization.MapAttributesForExport(string FlowRuleName, MVEntry mventry, CSEntry csentry)
        {
            ///
            /// Export attribute flow 
            ///
            int sadconnectors = 0;
            ConnectedMA sadMA = mventry.ConnectedMAs["Staging Area Database MA"];
            sadconnectors = sadMA.Connectors.Count;
            switch (FlowRuleName)
            {

                #region userPrincipalName


                // case "cd.user:userPrincipalName<-mv.person:Anglczd_first_nm,Anglczd_last_nm,Anglczd_middle_nm,Business_Area_CD,employeeID,employeeType,Exch_Trans_CD,Ext_Auth_Flag,External_Email_Address,homeMTA,Internet_Style_Adrs,mail,msDS-SourceObjectDN,msExchRecipientTypeDetails,org_unit_cd,personnel_area_cd,sAMAccountName,UC_Flag":
                case "cd.user:userPrincipalName<-mv.person:Anglczd_first_nm,Anglczd_last_nm,Anglczd_middle_nm,Business_Area_CD,employeeID,employeeType,Exch_Trans_CD,Ext_Auth_Flag,External_Email_Address,homeMTA,Internet_Style_Adrs,mail,mailNickname,msDS-SourceObjectDN,msExchRecipientTypeDetails,org_unit_cd,personnel_area_cd,sAMAccountName,UC_Flag":
                    //UPN set for the user translated out of XML file
                    //Deleted and throw error for UPN being null
                    //Filter for "Z" types, i.e build UPN only for "Z" types
                    if (mventry["sAMAccountName"].IsPresent
                        && mventry["personnel_area_cd"].IsPresent
                        && mventry["employeeType"].IsPresent
                        && etypeNodeValue.Contains(mventry["employeeType"].Value.ToUpper()))
                    {
                        //XML Method to get the values                        
                        setUserVariables(mventry["personnel_area_cd"].Value, mventry["employeeType"].Value);
                        //string tempupn = upn;
                        //BOC NLW mini Release
                        string tempupn = string.Empty;
                        string[] arrDcComponents = null;
                        string getUserTypeValue = CommonLayer.GetUserType(mventry, csentry, "RF", "no");
                        if (mventry["msDS-SourceObjectDN"].IsPresent)
                        {
                            int intDCCount = GetDCComponents(ref arrDcComponents, mventry["msDS-SourceObjectDN"].Value.ToString());
                            if (arrDcComponents != null)
                            {
                                if (intDCCount == 3)
                                {
                                    tempupn = arrDcComponents[0] + "." + arrDcComponents[1] + "." + arrDcComponents[2];
                                }
                                else
                                {
                                    tempupn = arrDcComponents[0] + "." + arrDcComponents[1] + "." + arrDcComponents[2] + "." + arrDcComponents[3];
                                }
                            }
                        }
                        //EOC NLW mini Release

                        //Start - CHG-CHG1178839, added below lines to read AllowableUPNDomains Email suffix list dynamically from Rules-config file.
                        //#CHG1324643 - Deploying below changes for AllowableUPNDomains
                        if (mventry["mail"].IsPresent)
                        {
                            string[] mailSuffix = mventry["mail"].Value.ToLower().Split('@');
                            if (strAllowableUPNDomains.Contains(mailSuffix[1]))
                            {
                                csentry["userPrincipalName"].Value = mventry["mail"].Value.ToLower();
                            }
                        }  //End - CHG-CHG1178839
                        else if (mventry["Exch_Trans_CD"].IsPresent && (mventry["Exch_Trans_CD"].Value == "1" || mventry["Exch_Trans_CD"].Value == "4") && getUserTypeValue != "MADS")
                        {
                            //get mailnickname

                            //HCM - Modifying logic to read MailNickName from MV (created in SAD)
                            //get mailnickname
                            //string defaultmNickname = CommonLayer.BuildMailNickName(mventry).ToLower();
                            //if (defaultmNickname != string.Empty)
                            // defaultmNickname = CommonLayer.GetUniqueMailNickname(defaultmNickname.ToLower(), mventry["sAMAccountName"].Value.ToLower());

                            string defaultmNickname = "";

                            if (mventry["mailNickname"].IsPresent)
                            {
                                //HCM - Setting Mailnickname from MV(calculated in SAD)

                                defaultmNickname = mventry["mailNickname"].Value;                               
                            }                            
							
							if(defaultmNickname != string.Empty)
							{
								if (mventry["employeeType"].Value.ToUpper() == "D")
								{
									if (mventry["Business_Area_CD"].Value.ToUpper() == "K" && (mventry["org_unit_cd"].IsPresent && mventry["org_unit_cd"].Value == "00115327"))
										csentry["userPrincipalName"].Value = defaultmNickname + "@" + strContractorEmailDomainPrefix + strAgspanDomainSuffix;
									else if (mventry["Business_Area_CD"].Value.ToUpper() == "K" && (mventry["org_unit_cd"].IsPresent && mventry["org_unit_cd"].Value != "00115327"))
										csentry["userPrincipalName"].Value = defaultmNickname + "@" + strContractorEmailDomainPrefix + strElancoDomainSuffix;
									else
										csentry["userPrincipalName"].Value = defaultmNickname + "@" + strContractorEmailDomainPrefix + strLillyDomainSuffix;
								}
								else
								{
									if (mventry["Business_Area_CD"].Value.ToUpper() == "K" && (mventry["org_unit_cd"].IsPresent && mventry["org_unit_cd"].Value == "00115327"))
										csentry["userPrincipalName"].Value = defaultmNickname + "@" + strAgspanDomainSuffix;
									else if (mventry["Business_Area_CD"].Value.ToUpper() == "K" && (mventry["org_unit_cd"].IsPresent && mventry["org_unit_cd"].Value != "00115327"))
										csentry["userPrincipalName"].Value = defaultmNickname + "@" + strElancoDomainSuffix;
									else
										csentry["userPrincipalName"].Value = defaultmNickname + "@" + strLillyDomainSuffix;
								}
							}
							else
							{
								//If email is not present set UPN with System ID
								if (!string.IsNullOrEmpty(tempupn))
									csentry["userPrincipalName"].Value = mventry["sAMAccountName"].Value + "@" + tempupn;
								else
									csentry["userPrincipalName"].Value = mventry["sAMAccountName"].Value + "@" + strRFDomain;
							}
                        }
                        else
                        {
                            //If email is not present set UPN with System ID
                            if (!string.IsNullOrEmpty(tempupn))
                                csentry["userPrincipalName"].Value = mventry["sAMAccountName"].Value + "@" + tempupn;
                            else
                                csentry["userPrincipalName"].Value = mventry["sAMAccountName"].Value + "@" + strRFDomain;
                        }
                    }
                    break;
                #endregion


                #region extensionAttribute14
                //BOC NLW mini release
                case "cd.user:extensionAttribute14<-mv.person:employeeType,EV_Flag,Exch_Trans_CD,personnel_area_cd,sAMAccountName,tnmsChangedFlags,tnmsClientPolicy,tnmsDialPlan,tnmsMailboxPolicy,tnmsRegistrarServer,tnmsTelephoneNumber,tnmsTNType,TNMSTransCode,tnmsVoicePolicy,UC_Flag":

                    string strExtensionAttribute = string.Empty;
                    string tnmsUCFlag = string.Empty;
                    string strUMFlag = "0";
                    string tnmsDialPlan, tnmsVoicePolicy, tnmsRegistrarServer, tnmsTelephoneNumber, tnmsClientPolicy, tnmsTNType, tnmsMailboxPolicy;
                    if (mventry["UC_Flag"].IsPresent && mventry["UC_Flag"].Value.ToUpper() == "Y")
                        tnmsUCFlag = "Y";
                    else if (mventry["UC_Flag"].IsPresent && mventry["UC_Flag"].Value.ToUpper() == "N")
                        tnmsUCFlag = "N";

                    if (mventry["tnmsDialPlan"].IsPresent)
                        tnmsDialPlan = mventry["tnmsDialPlan"].Value;
                    else
                        tnmsDialPlan = null;
                    if (mventry["tnmsVoicePolicy"].IsPresent)
                        tnmsVoicePolicy = mventry["tnmsVoicePolicy"].Value;
                    else
                        tnmsVoicePolicy = null;
                    if (mventry["tnmsRegistrarServer"].IsPresent)
                        tnmsRegistrarServer = mventry["tnmsRegistrarServer"].Value;
                    else
                        tnmsRegistrarServer = null;
                    if (mventry["tnmsTelephoneNumber"].IsPresent)
                        tnmsTelephoneNumber = mventry["tnmsTelephoneNumber"].Value;
                    else
                        tnmsTelephoneNumber = null;
                    if (mventry["tnmsClientPolicy"].IsPresent)
                        tnmsClientPolicy = mventry["tnmsClientPolicy"].Value;
                    else
                        tnmsClientPolicy = null;
                    if (mventry["tnmsTNType"].IsPresent)
                        tnmsTNType = mventry["tnmsTNType"].Value;
                    else
                        tnmsTNType = null;

                    if (!mventry["tnmsMailboxPolicy"].IsPresent)
                    {
                        setUserVariables(mventry["personnel_area_cd"].Value, mventry["employeeType"].Value);
                        tnmsMailboxPolicy = strUMMailboxPolicy;
                    }
                    else// if (mventry["tnmsMailboxPolicy"].IsPresent)
                    {
                        tnmsMailboxPolicy = mventry["tnmsMailboxPolicy"].Value;
                    }

                    if (mventry["Exch_Trans_Cd"].IsPresent && (mventry["Exch_Trans_Cd"].Value == "1" || mventry["Exch_Trans_Cd"].Value == "2") && (strOVAFlagSwitch == "ON" || mventry.ConnectedMAs["TNMS Data MA"].Connectors.Count > 0))
                    {
                        strUMFlag = "1";
                    }

                    if (!string.IsNullOrEmpty(tnmsMailboxPolicy))//mventry["tnmsDialPlan"].IsPresent && mventry["tnmsVoicePolicy"].IsPresent && mventry["tnmsRegistrarServer"].IsPresent && mventry["tnmsTelephoneNumber"].IsPresent && mventry["tnmsClientPolicy"].IsPresent && mventry["tnmsTNType"].IsPresent && mventry["tnmsMailboxPolicy"].IsPresent && mventry["Exch_Trans_Cd"].IsPresent)
                    {
                        if (mventry["EV_Flag"].IsPresent && mventry["EV_Flag"].BooleanValue)
                            strExtensionAttribute = "1:" + tnmsDialPlan + ":" + tnmsVoicePolicy + ":" + tnmsRegistrarServer + ":" + tnmsTelephoneNumber + ":" + tnmsClientPolicy + ":" + tnmsTNType + ":" + tnmsMailboxPolicy + ":" + tnmsUCFlag;// + ":" + mventry["Exch_Trans_Cd"].Value.ToUpper() + ":" + strUMFlag;
                        else                           
                            strExtensionAttribute = "0:" + tnmsDialPlan + ":" + tnmsVoicePolicy + ":" + tnmsRegistrarServer + ":" + tnmsTelephoneNumber + ":" + tnmsClientPolicy + ":" + tnmsTNType + ":" + tnmsMailboxPolicy + ":" + tnmsUCFlag;// + ":" + mventry["Exch_Trans_Cd"].Value.ToUpper() + ":" + strUMFlag;
                    }


                    if (mventry["tnmsChangedFlags"].IsPresent)
                    {
                        csentry["extensionAttribute14"].Value = mventry["tnmsChangedFlags"].Value + "," + strExtensionAttribute;
                    }//when the phone number is un-assigned:
                    else
                    {
                        csentry["extensionAttribute14"].Value = mventry["TNMSTransCode"].Value + " ," + strExtensionAttribute;
                    }

                    break;

                //EOC NLW mini release
                #endregion

                #region extensionAttribute8
                case "cd.user:extensionAttribute8<-mv.person:homeMDB,sAMAccountName":

                    if (csentry["homeMDB"].IsPresent && strERMASystemIDs.Contains(mventry["sAMAccountName"].Value.ToUpper()))
                    {
                        string strpartialhomeMDB = Convert.ToString(csentry["homeMDB"].Value.ToString().Split(',')[0].ToString());
                        int inthomeMDBLength = strpartialhomeMDB.Length;
                        string strhomeMDB = strpartialhomeMDB.Substring(1, inthomeMDBLength - 1).ToString();

                        csentry["extensionAttribute8"].Value = setERMAVariables(strhomeMDB);
                    }

                    break;
                #endregion
                //25-June-2021 - CHG1764452 - Workday phone number Flow Update to remove TNMS number precedence and create direct flow from SAD to RF
                /*
                #region telephoneNumber
                case "cd.user:telephoneNumber<-mv.person:Full_Work_Phone_Nbr,human_readable,sAMAccountName":

                    if (mventry["human_readable"].IsPresent)
                        csentry["telephoneNumber"].Value = mventry["human_readable"].Value.ToString();
                    else if (mventry["Full_Work_Phone_Nbr"].IsPresent)
                        csentry["telephoneNumber"].Value = mventry["Full_Work_Phone_Nbr"].Value.ToString();
                    else
                        csentry["telephoneNumber"].Delete();
                    break;
                #endregion
                */

                #region extensionAttribute15
                case "cd.user:extensionAttribute15<-mv.person:MigrationStatus,sAMAccountName":

                    if (!csentry["extensionAttribute15"].IsPresent)
                    {
                        if (mventry["MigrationStatus"].IsPresent && mventry["MigrationStatus"].Value.ToString() == "Pending FIM processing")
                        {
                            csentry["extensionAttribute15"].Value = "MIGRATE";
                        }
                    }
                    break;
                #endregion

                #region description
                case "cd.user:description<-mv.person:deprovisionedDate,employeeID":
                    //deprovisionedDate,employeeID,employeeStatus is added to trigger the code
                    //Used SAD connector count logic to enable/disble account
                    if (sadconnectors == 0
                        && mventry["deprovisionedDate"].IsPresent)
                    {
                        //Convert TerminatedDate to a DateTime object
                        DateTime TerminatedDate = Convert.ToDateTime(mventry["deprovisionedDate"].Value);
                        //Set the date in Format                   
                        string deprovisionedDate = TerminatedDate.ToString("dd-MMM-yyyy").ToUpper();
                        //Delete the value before setting with new value
                        csentry["description"].Delete();
                        csentry["description"].Value = "DDDD " + deprovisionedDate + " OTHR ILM";
                    }
                    else
                    {   //Delete the value to clean
                        csentry["description"].Delete();
                    }
                    break;
                #endregion description

                #region mailNickname

                // HCM - modified the flow to add mailnickname from MV
                //case "cd.user:mailNickname<-mv.person:Anglczd_first_nm,Anglczd_last_nm,Anglczd_middle_nm,Exch_Trans_CD,Ext_Auth_Flag,mailNicknameFromInitialLoad,sAMAccountName":
                case "cd.user:mailNickname<-mv.person:Anglczd_first_nm,Anglczd_last_nm,Anglczd_middle_nm,Exch_Trans_CD,Ext_Auth_Flag,mailNicknameFromInitialLoad,mailNickname,sAMAccountName":
                    //Set and forget mailNickName

                    if (mventry["Exch_Trans_CD"].IsPresent)
                    {
                        switch ((mventry["Exch_Trans_CD"].Value))
                        {
                            case "1": //These are all Mailbox types
                            case "2":
                            case "3":
                            case "4": // These are all contacts
                            case "5":
                            case "8":
                            case "6":
                            case "7":
                            case "10":
                                {
                                    if (!csentry["mailNickname"].IsPresent)
									{
										//HCM - Setting Mailnickname from MV(calculated in SAD)
											
										csentry["mailNickname"].Value = mventry["mailNickname"].Value;
											
										//string defaultmNickname = CommonLayer.BuildMailNickName(mventry);
										//if (defaultmNickname != string.Empty)
											// csentry["mailNickname"].Value = CommonLayer.GetUniqueMailNickname(defaultmNickname.ToLower(), mventry["sAMAccountName"].Value.ToLower());
									}
                                }
								
                                break;
                            case "9":
                            case "0":
                                csentry["mailNickname"].Delete();
                                break;
                        }
                    }

                    break;
                #endregion mailNickname

                #region msExchMasterAccountSid
                case "cd.user:msExchMasterAccountSid<-mv.person:Exch_Trans_CD,objectSid,sAMAccountName,WhereToProvision":
                    {
                        if (mventry["Exch_Trans_CD"].IsPresent)
                        {

                            switch ((mventry["Exch_Trans_CD"].Value))
                            {//These are all Mailbox types
                                case "1":
                                case "2":
                                case "3":
                                case "4":
                                case "5":
                                case "6":
                                case "7":
                                case "8":
                                case "10":
                                    {
                                        if (mventry["objectSid"].IsPresent)
                                        {
                                            csentry["msExchMasterAccountSid"].Value = mventry["objectSid"].Value;
                                        }
                                    }
                                    break;
                                case "9":
                                    {
                                        csentry["msExchMasterAccountSid"].Delete();
                                    }
                                    break;
                            }
                        }
                    }
                    break;
                #endregion

                #region extensionAttribute10
                //msExchTransitionDate
                //This function is to determine msExchtransitionDate (extensionAttribute10)
                case "cd.user:extensionAttribute10<-mv.person:Exch_Trans_CD,msExchRecipientTypeDetails,sAMAccountName":

                    if (mventry["Exch_Trans_CD"].IsPresent)
                    {
                        switch ((mventry["Exch_Trans_CD"].Value))
                        {
                            case "1":
                            case "3":
                            case "4":
                            case "5":
                                csentry["extensionAttribute10"].Delete();
                                break;
                            case "8":
                                break;
                            case "6":
                            case "9":
                                csentry["extensionAttribute10"].Delete();
                                break;
                            case "7":
                            case "10":
                                csentry["extensionAttribute10"].Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                break;
                            case "2":
                                if (csentry["targetAddress"].IsPresent && mventry["msExchRecipientTypeDetails"].IsPresent && mventry["msExchRecipientTypeDetails"].Value.ToString() != CommonLayer.REMOTE_MAILBOX)
                                    csentry["extensionAttribute10"].Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                else
                                    csentry["extensionAttribute10"].Delete();
                                break;
                        }
                    }
                    break;
                #endregion

                #region msRTCSIP-PrimaryUserAddress
                case "cd.user:msRTCSIP-PrimaryUserAddress<-mv.person:Exch_Trans_CD,msExchRecipientTypeDetails,msRTCSIP-UserEnabled,sAMAccountName,TargetAddress":
                    if ((!mventry["msRTCSIP-UserEnabled"].IsPresent) || (mventry["msRTCSIP-UserEnabled"].IsPresent && (!mventry["msRTCSIP-UserEnabled"].BooleanValue)))
                    {
                        if (mventry["msExchRecipientTypeDetails"].IsPresent && (mventry["msExchRecipientTypeDetails"].Value.ToString() == CommonLayer.MAIL_USER || mventry["msExchRecipientTypeDetails"].Value.ToString() == CommonLayer.MAIL_CONTACT) && csentry["targetAddress"].IsPresent)
                        {
                            if (!csentry["msRTCSIP-PrimaryUserAddress"].IsPresent)
                                for (int i = 0; i < arrSIPDomainList.Length; i++)
                                {
                                    if (csentry["targetAddress"].Value.ToString().Split('@')[1].ToString().ToLower() == arrSIPDomainList[i].ToString().ToLower())
                                    {
                                        csentry["msRTCSIP-PrimaryUserAddress"].Value = "sip:" + csentry["targetAddress"].Value.ToString().Split(':')[1];
                                        break;
                                    }

                                }
                            else
                            {
                                int deleteSIPAddress = 1;
                                for (int i = 0; i < arrSIPDomainList.Length; i++)
                                {
                                    if (csentry["targetAddress"].Value.ToString().Split('@')[1].ToString().ToLower() == arrSIPDomainList[i].ToString().ToLower())
                                    {
                                        deleteSIPAddress = 0;
                                        break;
                                    }
                                }
                                if (deleteSIPAddress == 1)
                                {
                                    csentry["msRTCSIP-PrimaryUserAddress"].Delete();
                                }
                            }
                        }
                    }
                    break;
                #endregion

                #region employeeType

                case "cd.user:employeeType<-mv.person:employeeType,Exch_Trans_CD,sAMAccountName":
                    if (mventry["Exch_Trans_CD"].IsPresent && (!(mventry["Exch_Trans_CD"].Value == "1" ||
                                           mventry["Exch_Trans_CD"].Value == "2"))
                        )
                    {
                        csentry["employeeType"].Value = mventry["employeeType"].Value;
                    }
                    break;
                #endregion
               

                default:
                    throw new EntryPointNotImplementedException();
            }
        }

        public Hashtable LoadAFConfig(string filename)
        {
            //
            // Initializing AD config file, any new PArent Tag in the config means code change / addition to this section 
            //

            XmlTextReader xmlReader = new XmlTextReader("C:\\Program Files\\Microsoft Identity Integration Server\\Extensions\\" + filename);
            // Read the line of the xml file
            Hashtable primaryhash = new Hashtable();

            while (xmlReader.ReadToFollowing("PARENT"))
            {
                Hashtable valhash = new Hashtable();
                if (xmlReader.HasAttributes)
                {
                    for (int i = 0; i < xmlReader.AttributeCount; i++)
                    {
                        xmlReader.MoveToAttribute(i);
                        primaryhash.Add(xmlReader.Value, valhash);
                    }
                }
                if (xmlReader.ReadToFollowing("PATH"))
                {
                    xmlReader.Read();
                    string path = xmlReader.Value;
                    valhash.Add("PATH", path);
                }

                if (xmlReader.ReadToFollowing("UPN"))
                {
                    xmlReader.Read();
                    string upn = xmlReader.Value;
                    valhash.Add("UPN", upn);
                }
                if (xmlReader.ReadToFollowing("DOMAIN"))
                {
                    xmlReader.Read();
                    string domain = xmlReader.Value;
                    valhash.Add("DOMAIN", domain);
                }

                //change
                //Release 5 Lilac
                if (xmlReader.ReadToFollowing("LOGINSCRIPT"))//
                {
                    xmlReader.Read();
                    string loginscript = xmlReader.Value;
                    valhash.Add("LOGINSCRIPT", loginscript);
                }
                if (xmlReader.ReadToFollowing("EMAILTEMPLATE"))//
                {
                    xmlReader.Read();
                    string emailtemplate = xmlReader.Value;
                    valhash.Add("EMAILTEMPLATE", emailtemplate);
                }
                //Release 5 Lilac
                //change

                if (xmlReader.ReadToFollowing("SIPHOMESERVER"))
                {
                    xmlReader.Read();
                    string domain = xmlReader.Value;
                    valhash.Add("SIPHOMESERVER", domain);
                }

                if (xmlReader.ReadToFollowing("ARCHIVINGENABLED"))
                {
                    xmlReader.Read();
                    string archivingenabled = xmlReader.Value;
                    valhash.Add("ARCHIVINGENABLED", archivingenabled);
                }
                if (xmlReader.ReadToFollowing("USERLOCATIONPROFILE"))
                {
                    xmlReader.Read();
                    string userlocationprofile = xmlReader.Value;
                    valhash.Add("USERLOCATIONPROFILE", userlocationprofile);
                }
                if (xmlReader.ReadToFollowing("USERPOLICY"))
                {
                    ArrayList userpolicyArray = new ArrayList();
                    XmlReader tempxmlReader = (XmlReader)xmlReader.ReadSubtree();
                    while (tempxmlReader.ReadToFollowing("VALUE"))
                    {
                        tempxmlReader.Read();
                        string strvalue = tempxmlReader.Value;
                        //userpolicyArray.Add(strvalue.ToUpper());
                        userpolicyArray.Add(strvalue);
                    }
                    valhash.Add("USERPOLICY", userpolicyArray);
                }

                if (xmlReader.ReadToFollowing("MAILBOXPOLICY"))
                {
                    xmlReader.Read();
                    string UMMailboxPolicy = xmlReader.Value;
                    valhash.Add("MAILBOXPOLICY", UMMailboxPolicy);
                }

                if (xmlReader.ReadToFollowing("OPTIONFLAGS"))
                {
                    xmlReader.Read();
                    string useroptionflags = xmlReader.Value;
                    valhash.Add("OPTIONFLAGS", useroptionflags);
                }
                if (xmlReader.ReadToFollowing("FEDERATIONENABLED"))
                {
                    xmlReader.Read();
                    string userfederationenabled = xmlReader.Value;
                    valhash.Add("FEDERATIONENABLED", userfederationenabled);
                }

                if (xmlReader.ReadToFollowing("skipmail"))
                {
                    xmlReader.Read();
                    string strskipmail = xmlReader.Value;
                    valhash.Add("skipmail", strskipmail.ToLower());
                }

                if (xmlReader.ReadToFollowing("PERSONNELAREA"))
                {
                    ArrayList parray = new ArrayList();
                    XmlReader tempxmlReader = (XmlReader)xmlReader.ReadSubtree();
                    while (tempxmlReader.ReadToFollowing("CODE"))
                    {
                        tempxmlReader.Read();
                        string pcode = tempxmlReader.Value;
                        parray.Add(pcode);
                    }
                    valhash.Add("PERSONNELAREA", parray);
                }

                if (xmlReader.ReadToFollowing("CONSTITUENT"))
                {
                    ArrayList etypearray = new ArrayList();
                    XmlReader tempxmlReader = (XmlReader)xmlReader.ReadSubtree();
                    while (tempxmlReader.ReadToFollowing("EMPLY_GRP"))
                    {
                        tempxmlReader.Read();
                        string etype = tempxmlReader.Value;
                        etypearray.Add(etype.ToUpper());
                    }
                    valhash.Add("CONSTITUENT", etypearray);
                }



                string strNameExchHomeServer = "";
                string[,] strArrRandomized = new string[250, 2];
                int j = 0;
                if (xmlReader.ReadToFollowing("REGION"))
                {
                    ArrayList ExchHomeServer = new ArrayList();
                    XmlReader exchxmlReader = (XmlReader)xmlReader.ReadSubtree();
                    while (exchxmlReader.ReadToFollowing("ExchHomeServer"))
                    {
                        if (exchxmlReader.AttributeCount > 0)
                        {
                            strNameExchHomeServer = exchxmlReader.GetAttribute("name");
                        }
                        XmlReader mdbxmlReader = (XmlReader)xmlReader.ReadSubtree();
                        while (mdbxmlReader.ReadToFollowing("homeMDB"))
                        {
                            mdbxmlReader.Read();
                            strArrRandomized[j, 0] = (strNameExchHomeServer);
                            strArrRandomized[j, 1] = (mdbxmlReader.Value);
                            j = j + 1;
                        }
                        ExchHomeServer.Add(strNameExchHomeServer);

                    }
                    valhash.Add("MSEXCHHOMESERVER", ExchHomeServer);
                    valhash.Add("RANDOMIZEDARRAY", strArrRandomized);
                }
                //change
                if (xmlReader.ReadToFollowing("WhereToProvision"))//
                {
                    xmlReader.Read();
                    string WhereToProv = xmlReader.Value;
                    valhash.Add("WhereToProvision", WhereToProv);
                }
            }
            return primaryhash;
        }

        public void setUserVariables(string paccode, string employeeType)
        {
            //
            // Initializing the variable necessary to build a AD account from the already initialized AD config file in a Hashtable
            //

            foreach (DictionaryEntry de in afvalhash)
            {
                Hashtable afconfig = (Hashtable)de.Value;
                ArrayList listarea = (ArrayList)afconfig["PERSONNELAREA"];
                ArrayList listtype = (ArrayList)afconfig["CONSTITUENT"];
                publicOu = null;
                domain = null;
                upn = null;
                sipHomeServer = null;

                if (listarea.Contains(paccode) && listtype.Contains(employeeType.ToUpper()))
                {
                    publicOu = afconfig["PATH"].ToString();
                    domain = afconfig["DOMAIN"].ToString().ToUpper();
                    upn = afconfig["UPN"].ToString();
                    sipHomeServer = afconfig["SIPHOMESERVER"].ToString();
                    archivingenabled = afconfig["ARCHIVINGENABLED"].ToString();        //Exchange OCS code
                    userlocationprofile = afconfig["USERLOCATIONPROFILE"].ToString();  //Exchange OCS code
                    userpolicy = (ArrayList)afconfig["USERPOLICY"];     //Exchange OCS code
                    //userinternetaccessenabled = afconfig["INTERNETACCESSENABLED"].ToString();        //Exchange OCS code
                    userfederationenabled = afconfig["FEDERATIONENABLED"].ToString();        //CleanUp Release Code modified
                    useroptionflags = afconfig["OPTIONFLAGS"].ToString();
                    strUMMailboxPolicy = afconfig["MAILBOXPOLICY"].ToString();//UM Mailbox Policy
                    strskipmail = afconfig["skipmail"].ToString(); //Exchange

                    strArrServerValues = getRandomizedValues((string[,])afconfig["RANDOMIZEDARRAY"]);
                    impstrArrServerValues = (ArrayList)afconfig["MSEXCHHOMESERVER"];

                    break;
                }
            }
        }

        public string[] getRandomizedValues(string[,] strArrInput)
        {
            Random intRandomNumber = new Random();
            string[] strArrOut = new string[2];
            int intRandom = 0;
            int i = 0;

            while (strArrInput[i, 0] != null)
            {
                i++;
            }

            intRandom = intRandomNumber.Next(0, i - 1);
            strArrOut[0] = strArrInput[intRandom, 0];
            strArrOut[1] = strArrInput[intRandom, 1];
            return strArrOut;
        }

        public ArrayList LoadPAConfig(string XML_CONFIG_FILE)
        {
            ArrayList arrList = new ArrayList();
            XmlTextReader xmlReader = new XmlTextReader(XML_CONFIG_FILE);

            if (xmlReader.ReadToFollowing("proxyaddresses"))
            {
                XmlReader tempxmlReader = (XmlReader)xmlReader.ReadSubtree();
                while (tempxmlReader.ReadToFollowing("value"))
                {
                    tempxmlReader.Read();
                    string paddess = tempxmlReader.Value;
                    arrList.Add(paddess);
                }
            }
            return arrList;
        }

        public static string StripName(string strName)
        {
            string strtemp = "";
            string[] tmpdmn = strName.Split(' ');
            for (int i = 0; i < tmpdmn.Length; i++)
            {
                strtemp = strtemp + tmpdmn[i].ToString();
                if ((i == 1) && (tmpdmn[i].ToString() == string.Empty))
                {
                    strtemp = strtemp + "_";
                }
                if ((i > 1) && (tmpdmn[i].ToString() == string.Empty))
                {
                    strtemp = strtemp;
                }
            }
            return strtemp;
        }




        private bool IsAcceptedDomain(string proxyElement, ValueCollection vcAcceptedDomain)
        {
            //Accepted domain check
            foreach (Value acceptedDomainElement in vcAcceptedDomain)
            {
                if (proxyElement.ToString().Split('@')[1].ToLower() == acceptedDomainElement.ToString())
                {
                    return true;
                }
            }
            return false;

        }

        private bool ContainsIgnoreCase(ValueCollection vcCollection, string strCompareValue)
        {
            foreach (Value CollectionElement in vcCollection)
            {
                if (CollectionElement.ToString().ToLower() == strCompareValue.ToLower())
                {
                    return true;
                }
            }
            return false;
        }

        private bool ContainsCaseComparison(ValueCollection vcCollection, string strCompareValue)
        {
            foreach (Value CollectionElement in vcCollection)
            {
                if (CollectionElement.ToString() == strCompareValue.ToString())
                {
                    return true;
                }
            }
            return false;
        }

        public Hashtable LoadERMAConfig()
        {
            Hashtable primaryhashERMA = new Hashtable();
            try
            {
                XmlTextReader xmlReader = new XmlTextReader("C:\\Program Files\\Microsoft Forefront Identity Manager\\2010\\Synchronization Service\\Extensions\\" + "ERMA.xml");

                // Read the line of the xml file

                while (xmlReader.ReadToFollowing("Database"))
                {
                    Hashtable valhashERMA = new Hashtable();
                    if (xmlReader.HasAttributes)
                    {
                        for (int i = 0; i < xmlReader.AttributeCount; i++)
                        {
                            xmlReader.MoveToAttribute(i);
                            primaryhashERMA.Add(xmlReader.Value, valhashERMA);
                        }
                    }
                    if (xmlReader.ReadToFollowing("Entry"))
                    {
                        xmlReader.Read();
                        string entry = xmlReader.Value;
                        valhashERMA.Add("ENTRY", entry);
                    }
                    if (xmlReader.ReadToFollowing("CustomAttribute"))
                    {
                        xmlReader.Read();
                        string customAttribute = xmlReader.Value;
                        valhashERMA.Add("ATTRIBUTE", customAttribute);
                    }
                }
                return primaryhashERMA;
            }
            catch (Exception ex)
            {
                return primaryhashERMA;
            }
        }

        private int GetDCComponents(ref string[] strDCComponents, string strCsentry)
        {
            string strCsEntryDnString = strCsentry.ToString();
            string[] arrstrDCComponents = new string[4];
            int intDCindex = strCsEntryDnString.IndexOf("DC=");
            int intDCCount = 0;

            while (intDCindex != -1)
            {
                intDCCount++;

                if (intDCindex + 2 <= strCsEntryDnString.Length)
                {
                    strCsEntryDnString = strCsEntryDnString.Substring(intDCindex + 2);

                    int intcommaindex = strCsEntryDnString.IndexOf(",");
                    if (intcommaindex != -1)
                    {
                        arrstrDCComponents[intDCCount - 1] = strCsEntryDnString.Substring(1, intcommaindex - 1);
                    }
                    else
                    {
                        arrstrDCComponents[intDCCount - 1] = strCsEntryDnString.Substring(1, strCsEntryDnString.Length - 1);
                    }


                    intDCindex = strCsEntryDnString.IndexOf("DC=");

                }
                else
                {
                    intDCindex = -1;
                }


            }
            strDCComponents = arrstrDCComponents;
            return intDCCount;
        }

        public string setERMAVariables(string strMailbox)
        {
            string strERMAAttribute = string.Empty;
            foreach (DictionaryEntry de in valhashERMADatabase)
            {
                Hashtable ermaconfig = (Hashtable)de.Value;
                string strERMADatabase = ermaconfig["ENTRY"].ToString();
                if (strMailbox.ToUpper().EndsWith(strERMADatabase.ToUpper()))
                {
                    strERMAAttribute = ermaconfig["ATTRIBUTE"].ToString();
                    break;
                }
            }
            return strERMAAttribute;
        }

    }
}
