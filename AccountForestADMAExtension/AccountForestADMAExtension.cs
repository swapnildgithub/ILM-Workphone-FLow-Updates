using System;
using System.Collections;
using Microsoft.MetadirectoryServices;
using System.Xml;
using System.Diagnostics;
using CommonLayer_NameSpace;

namespace Mms_ManagementAgent_AccountForestADMAExtension
{
    /// <summary>
    /// Extension for Account Forest AD MA's. It also maps attributes, read groups and set the condition of the users for LCS and Exchange account provisioning.
    /// </summary>
    public class MAExtensionObject : IMASynchronization
    {
        XmlNode rnode;
        XmlNode node;
        string version, etypeNodeValue, publicOu, domain, upn, sipHomeServer, adminoverridetext, strexchangemailprefix;  //Release 5 Lilac
       // string strLillyDomainSuffix, strElancoDomainSuffix, strAgspanDomainSuffix, strContractorEmailDomainPrefix,strZone1Domain;
       string strLillyDomainSuffix, strElancoDomainSuffix, strAgspanDomainSuffix, strContractorEmailDomainPrefix,strZone1Domain, strAllowableUPNDomains;//Start - CHG-CHG1178839
        Hashtable afvalhash;

        //CHG1471559 - Initial Password Update
        int MinInitPwdLength, MaxInitPwdLength;

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
                //Provisioning Version
                node = rnode.SelectSingleNode("version");
                version = node.InnerText;
                //get the Users to be provisioned from config                       
                node = rnode.SelectSingleNode("provision");
                etypeNodeValue = node.InnerText.ToUpper();
                //Release 5 Lilac

                
                //AM domain for UPN of Unix service account
                node = rnode.SelectSingleNode("AMDomain");
                strZone1Domain = node.InnerText.ToUpper();

                node = rnode.SelectSingleNode("AdminOverrideText");
                adminoverridetext = node.InnerText.ToUpper();
                //Release 5 Lilac

                //CHG1471559 - Initial Password Update
                node = rnode.SelectSingleNode("MinInitPwdLength");
                MinInitPwdLength = Int32.Parse(node.InnerText);

                //CHG1471559 - Initial Password Update
                node = rnode.SelectSingleNode("MaxInitPwdLength");
                MaxInitPwdLength = Int32.Parse(node.InnerText);

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
            if (
                (csentry.DN.ToString().ToLower().Contains("privileged") && csentry["sAMAccountName"].Value.ToLower().EndsWith("-ds"))
                ||
                (csentry.DN.ToString().ToLower().Contains("cloud privileged"))
                )
            {
                MVObjectType = "PrivilegedAccount";
                return true;
            }
            else if (csentry.DN.ToString().ToLower().Contains("cloud service"))
            {
                MVObjectType = "ServiceAccounts";
                return true;
            }
            else
            {
                MVObjectType = "PrivilegedAccount";
                return false;
            }
        }

        DeprovisionAction IMASynchronization.Deprovision(CSEntry csentry)
        {
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

            // Mapping attribute flow for Import.
            //
            switch (FlowRuleName)
            {
                case "cd.user:sAMAccountName->mv.PrivilegedAccount:ms-DS-ConsistencyGuidPrivAcct":
                    //CHG1581173 - generate new GUID to set in Consistency Guid for cloud and Privileged accounts.
                    // if (!mventry["ms-DS-ConsistencyGuidPrivAcct"].IsPresent && csentry["sAMAccountName"].Value.ToLower().EndsWith("-ca"))
                    if (!mventry["ms-DS-ConsistencyGuidPrivAcct"].IsPresent && (csentry["sAMAccountName"].Value.ToLower().EndsWith("-ca") || csentry["sAMAccountName"].Value.ToLower().EndsWith("-ds")))
                        mventry["ms-DS-ConsistencyGuidPrivAcct"].BinaryValue = Guid.NewGuid().ToByteArray();
                    break;

                //CHG1471559 - Initial Password Update Modifying initial password char limits for new AD password Policies
                case "cd.user:sAMAccountName->mv.PrivilegedAccount:initpwdPrivlgdAcct":
                    if (!mventry["initpwdPrivlgdAcct"].IsPresent)
                        mventry["initpwdPrivlgdAcct"].Value = RandomPassword.Generate(MinInitPwdLength, MaxInitPwdLength);
                    break;

                // CHG1471559 - Initial Password UpdateModifying initial password char limits for new AD password Policies
                case "cd.user:<dn>,sAMAccountName->mv.ServiceAccounts:initpwdCloudServiceAcct":
                    if (!mventry["initpwdCloudServiceAcct"].IsPresent && csentry.DN.ToString().ToLower().Contains("ou=cloud"))
                        mventry["initpwdCloudServiceAcct"].Value = RandomPassword.Generate(MinInitPwdLength, MaxInitPwdLength);
                    break;


                case "cd.user:<dn>,sAMAccountName->mv.ServiceAccounts:ms-DS-ConsistencyGuidServiceAcct":
                    //generate new GUID to set in Consistency Guid for cloud
                    if (!mventry["ms-DS-ConsistencyGuidServiceAcct"].IsPresent && csentry.DN.ToString().ToLower().Contains("ou=cloud"))
                        mventry["ms-DS-ConsistencyGuidServiceAcct"].BinaryValue = Guid.NewGuid().ToByteArray();

                    break;

                
               case "cd.user:<dn>,userAccountControl->mv.person:userAccountControl":

                    if (csentry["userAccountControl"].IsPresent)
                    {
                        mventry["userAccountControl"].Value = csentry["userAccountControl"].Value;
                    }
                    break;


                //BOC Mini NLW Release
                case "cd.user:sAMAccountName,userAccountControl->mv.person:Dsbld_dt":

                    if (csentry["userAccountControl"].IsPresent)
                    {
                        if (csentry["userAccountControl"].IntegerValue == 514)
                        {
                            mventry["Dsbld_dt"].Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                        }
                        else
                        {
                            mventry["Dsbld_dt"].Delete();
                        }

                    }

                    break;

                //EOC Mini NLW Release
                case "cd.user:<dn>,cn->mv.person:msDS-SourceObjectDN":
                    //Convert the users DN reference to a string
                    string sName = csentry.DN.ToString();
                    mventry["msDS-SourceObjectDN"].Value = sName;
                    break;

                case "cd.user:pwdLastSet->mv.person:initPassword":

                    if (csentry["pwdlastset"].Value != "0")
                    {
                        mventry["initpassword"].Value = "CHANGED";
                    }
                    break;
                //Release 5 Lilac
                case "cd.user:description,employeeID->mv.person:admin_override_flag":
                    if ((csentry["description"].IsPresent) && (csentry["description"].Value.Contains(adminoverridetext)))
                    {
                        mventry["admin_override_flag"].Value = "True";
                    }
                    else
                    {
                        mventry["admin_override_flag"].Value = "False";
                    }
                    break;

                case "cd.user:whenCreated->mv.person:afcreate_dt":
                    string strDatetime = csentry["whenCreated"].Value.ToString();
                    int year = Convert.ToInt32(strDatetime.Substring(0, 4));
                    int month = Convert.ToInt32(strDatetime.Substring(4, 2));
                    int day = Convert.ToInt32(strDatetime.Substring(6, 2));
                    int hh = Convert.ToInt32(strDatetime.Substring(8, 2));
                    int mm = Convert.ToInt32(strDatetime.Substring(10, 2));
                    int dd = Convert.ToInt32(strDatetime.Substring(12, 2));
                    DateTime dt = new DateTime(year, month, day, hh, mm, dd, 123);
                    mventry["afcreate_dt"].Value = dt.ToString("yyyy-MM-dd HH:mm:ss");

                    break;
                //Release 5 Lilac
                case "cd.user:<dn>,sAMAccountName,userPrincipalName->mv.person:UPN":
                    if (csentry["userPrincipalName"].IsPresent)
                    {
                        mventry["UPN"].Value = csentry["userPrincipalName"].Value;
                    }
                    break;
                case "cd.user:<dn>,sAMAccountName,userPrincipalName->mv.person:userPrincipalName":
                    if (csentry["userPrincipalName"].IsPresent)
                    {
                        mventry["userPrincipalName"].Value = csentry["userPrincipalName"].Value;
                    }
                    break;
                // O365 bug fix - Set the pwdLastSet to 0 when a User is deprovisioned
                case "cd.user:<dn>,pwdLastSet->mv.person:pwd_last_set":
                    if (csentry["pwdLastSet"].IsPresent)
                    {
                        mventry["pwd_last_set"].Value = csentry["pwdLastSet"].Value;
                    }
                    break;

                default:
                    throw new EntryPointNotImplementedException();
            }

        }

        void IMASynchronization.MapAttributesForExport(string FlowRuleName, MVEntry mventry, CSEntry csentry)
        {
            ///
            /// Export attribute flow 
            ///
            const long ADS_UF_NORMAL_ACCOUNT = 0x0200;
            const long ADS_UF_ACCOUNTDISABLE = 0x0002;
            const string USER_ACCOUNT_CONTROL_PROP = "userAccountControl";

            int sadconnectors = 0;
            ConnectedMA sadMA = mventry.ConnectedMAs["Staging Area Database MA"];
            sadconnectors = sadMA.Connectors.Count;

            switch (FlowRuleName)
            {                
                case "cd.user:userAccountControl<-mv.person:admin_override_flag,deprovisionedDate,employeeID,employeeType,emply_sub_grp_cd,System_Access_Flag":
                //Release 5 - Lilac. Process All employee types mentioned in Rules.config
                //The account if disabled manually for Admin purpose then it should not be reenabled by MIIS. 
                //So while disabling the account manually ,perticular Admin Override Text should be present in desciption field e.g. "ADMIN" this Text configurable from rules.config file.
                //If description is present and it is "ADMIN" (whatever text mentioned in rules.config) then make the admin_override_flag as True. If True then skip processing.
                if (mventry["admin_override_flag"].Value == "False")
                {
                    long currentValue = ADS_UF_NORMAL_ACCOUNT;
                    if (csentry[USER_ACCOUNT_CONTROL_PROP].IsPresent)
                    {
                        currentValue = csentry[USER_ACCOUNT_CONTROL_PROP].IntegerValue;
                    }

                    //Account changed to enabled status by adding enable value tot he current value
                    if (sadconnectors == 1)
                    {
                        csentry[USER_ACCOUNT_CONTROL_PROP].IntegerValue = (currentValue | ADS_UF_NORMAL_ACCOUNT)
                                                                          & ~ADS_UF_ACCOUNTDISABLE;
                    }
                    //Account changed to disabled status by adding disable value tot he current value
                    else
                    {
                        csentry[USER_ACCOUNT_CONTROL_PROP].IntegerValue = currentValue
                                                                  | ADS_UF_ACCOUNTDISABLE;
                    }

                }
                break;

            //case "cd.user:description<-mv.person:deprovisionedDate,employeeID,employeeType":  //Release 5 Lilac
            case "cd.user:description<-mv.person:admin_override_flag,deprovisionedDate,employeeID,employeeType":
                //Release 5 - Lilac.  Process All employee types mentioned in Rules.config
                //The account if disabled manually for Admin purpose then it should not be reenabled by MIIS. 
                //So while disabling the account manually ,perticular Admin Override Text should be present in desciption field e.g. "ADMIN" this Text configurable from rules.config file.
                //If description is present and it is "ADMIN" (whatever text mentioned in rules.config) then make the admin_override_flag as True. If True then skip processing.
                if (mventry["admin_override_flag"].Value == "False")
                {
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
                }
                break;
            //O365 bug fix - Set the pwdLastSet of deprovisioned Users to 0 - this will force Change Password At Next Logon.
            case "cd.user:pwdLastSet<-mv.person:deprovisionedDate,employeeID,employeeType":
                //The day User account is deprovisioned, set the pwdLastSet to 0
                if (sadconnectors == 0 && mventry["deprovisionedDate"].IsPresent)
                {
                    csentry["pwdLastSet"].IntegerValue = 0;
                }

                break;
           //HCM- modified rules extension to flow mailnickname from MV
            //case "cd.user:userPrincipalName<-mv.person:Anglczd_first_nm,Anglczd_last_nm,Anglczd_middle_nm,Business_Area_CD,employeeID,employeeType,Exch_Trans_CD,Ext_Auth_Flag,External_Email_Address,homeMTA,Internet_Style_Adrs,mail,msExchRecipientTypeDetails,org_unit_cd,personnel_area_cd,sAMAccountName,UC_Flag":
            case "cd.user:userPrincipalName<-mv.person:Anglczd_first_nm,Anglczd_last_nm,Anglczd_middle_nm,Business_Area_CD,employeeID,employeeType,Exch_Trans_CD,Ext_Auth_Flag,External_Email_Address,homeMTA,Internet_Style_Adrs,mail,mailNickname,msExchRecipientTypeDetails,org_unit_cd,personnel_area_cd,sAMAccountName,UC_Flag":
                //UPN set for the user translated out of XML file
                //Deleted and throw error for UPN being null
                //Filter for "Z" types, i.e build UPN only for "Z" types
                if (mventry["sAMAccountName"].IsPresent
                    && mventry["personnel_area_cd"].IsPresent
                    && mventry["employeeType"].IsPresent
                    && etypeNodeValue.Contains(mventry["employeeType"].Value.ToUpper()))
                {
                    string sSource, sLog, sEvent;
                    //XML Method to get the values                        
                    setUserVariables(mventry["personnel_area_cd"].Value, mventry["employeeType"].Value);
                    //string tempupn = upn;
                    //BOC NLW mini Release
                    string tempupn = string.Empty;
                    string[] arrDcComponents = null;
                    string getUserTypeValue = CommonLayer.GetUserType(mventry, csentry, "AF", "no");
                    int intDCCount = GetDCComponents(ref arrDcComponents, csentry.ToString());
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

                    //EOC NLW mini Release

                    if (tempupn != null)
                    {//O365 release. Set UPN with email
                     //Start - CHG-CHG1178839, added below lines to read AllowableUPNDomains Email suffix list dynamically from Rules-config file.
                     //#CHG1324643 - Deploying below changes for AllowableUPNDomains
                        if (mventry["mail"].IsPresent)
                        {
                            string[] mailSuffix = mventry["mail"].Value.ToLower().Split('@');
                            if (strAllowableUPNDomains.Contains(mailSuffix[1]))
                            {
                                csentry["userPrincipalName"].Value = mventry["mail"].Value.ToLower();
                            }
                        }//End - CHG-CHG1178839
                        else if (mventry["Exch_Trans_CD"].IsPresent && (mventry["Exch_Trans_CD"].Value == "1" || mventry["Exch_Trans_CD"].Value == "4") && getUserTypeValue != "MADS")
                        {
                            //HCM - Modifying logic to read MailNickName from MV (created in SAD)
                            //get mailnickname
                            /*
                            string defaultmNickname = CommonLayer.BuildMailNickName(mventry).ToLower();
                            if (defaultmNickname != string.Empty)
                                defaultmNickname = CommonLayer.GetUniqueMailNickname(defaultmNickname.ToLower(), mventry["sAMAccountName"].Value.ToLower());
                            */

                                string defaultmNickname = "";

                                if (mventry["mailNickname"].IsPresent)
                                {
                                    //HCM - Setting Mailnickname from MV(calculated in SAD)

                                    defaultmNickname = mventry["mailNickname"].Value;
                                }

                                if (defaultmNickname != string.Empty)
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
								else// mailnickname is null or empty in MV, set the default upn
								{
									csentry["userPrincipalName"].Value = mventry["sAMAccountName"].Value + "@" + tempupn;
								}
                            }
                            else//If email is not present set UPN with System ID
                                csentry["userPrincipalName"].Value = mventry["sAMAccountName"].Value + "@" + tempupn;
                        }
                        else
                        {
                            csentry["userPrincipalName"].Delete();
                            //Write to the Event log in case these necessary data isn't available
                            if (tempupn == null)
                            {
                                string ExceptionMessage = mventry["employeeID"].Value
                                    + " - PACCODE and EMPLOYEETYPE combination cannot retrieve userPrincipalName data from XML file";
                                sSource = "MIIS RF-AD MA Export";
                                sLog = "Application";
                                sEvent = ExceptionMessage;

                                if (!EventLog.SourceExists(sSource))
                                    EventLog.CreateEventSource(sSource, sLog);

                                EventLog.WriteEntry(sSource, sEvent, EventLogEntryType.Error, 8104);
                            }
                        }
                    }
                    break;

                case "cd.user:sAMAccountName<-mv.person:employeeID,employeeType,sAMAccountName":
                    //Filter the value for "Z" type only
                    //employeeID,employeeType,sAMAccountName is added to trigger the code
                    //Not Deleted in case sAMAccountName is empty in mv since the join is using sAMAccountName
                    if (mventry["sAMAccountName"].IsPresent
                        && mventry["employeeType"].IsPresent
                        && etypeNodeValue.Contains(mventry["employeeType"].Value.ToUpper()))
                    {
                        csentry["sAMAccountName"].Value = mventry["sAMAccountName"].Value;
                    }
                    break;
                //25-June-2021 - CHG1764452 - Workday phone number Flow Update to remove TNMS number precedence and create direct flow from SAD to RF
                /*
                case "cd.user:telephoneNumber<-mv.person:Full_Work_Phone_Nbr,human_readable,sAMAccountName":

                    if (mventry["human_readable"].IsPresent)
                        csentry["telephoneNumber"].Value = mventry["human_readable"].Value.ToString();
                    else if (mventry["Full_Work_Phone_Nbr"].IsPresent)
                        csentry["telephoneNumber"].Value = mventry["Full_Work_Phone_Nbr"].Value.ToString();
                    else
                        csentry["telephoneNumber"].Delete();
                    break;
                   */

                case "cd.user:userPrincipalName<-mv.ServiceAccounts:initpwdCloudServiceAcct,uid":
                    if (!(mventry["initpwdCloudServiceAcct"].IsPresent) && mventry["uid"].IsPresent)
                        csentry["userPrincipalName"].Value = mventry["uid"].Value.ToString() + "@" + strZone1Domain.ToLower();
                    break;

                case"cd.user:displayName<-mv.ServiceAccounts:initpwdCloudServiceAcct,serviceAccountName":
                    if (!(mventry["initpwdCloudServiceAcct"].IsPresent) && mventry["serviceAccountName"].IsPresent)
                        csentry["displayName"].Value = mventry["serviceAccountName"].Value.ToString();
                    break;

                default:
                    throw new EntryPointNotImplementedException();
            }
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
                if (xmlReader.ReadToFollowing("SIPHOMESERVER"))
                {
                    xmlReader.Read();
                    string domain = xmlReader.Value;
                    valhash.Add("SIPHOMESERVER", domain);
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
                    break;
                }
            }
        }

        //BOC NLW mini Release 
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
        //EOC NLW mini Release 

    }
}
