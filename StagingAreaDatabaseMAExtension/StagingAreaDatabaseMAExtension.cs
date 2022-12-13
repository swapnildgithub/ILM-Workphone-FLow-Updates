using System;
using System.DirectoryServices;
using Microsoft.MetadirectoryServices;
using System.Reflection;
using System.Xml;
using System.Diagnostics;
using System.Collections;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using CommonLayer_NameSpace;
using System.Globalization;
using System.Threading;
//using Microsoft.IdentityManagement.Logging;
//using Microsoft.MetadirectoryServices.Logging;

namespace Mms_ManagementAgent_StagingAreaDatabaseMAExtension
{

    /// <summary>
    /// Extension for SAD MA. Trims the all teh values of SAD data and maps for flow.
    /// </summary>
    public class MAExtensionObject : IMASynchronization
    {

       

        string strMembers;
        IMVSynchronization[] myMVDlls;
        XmlNode rnode;
        XmlNode node;
        string publicOu, domain, upn, etypeNodeValue, sipHomeServer, toggledisplayname, emailtemplate, loginscript;
        //Release 5 -  added 2 new variables emailtemplate, loginscript
        string archivingenabled, userlocationprofile, notesdomain, notesou, env, toggledtag, strPACWhereToprovision;
        string strexchangemailprefix, strskipmail, strmsExchmailprovisiontype, strexchtransitionlogs, strexchtransitiontime, lstrexchtransitionlogs, strexchtransitionlogspath, useroptionflags, userfederationenabled; //-----CleanUp Release Code modifieds
        Hashtable afvalhash;
        ArrayList userpolicy;
        string workerOu = "";
        string strGlobalPAC = string.Empty;
        string[] strArrServerValues;
        ArrayList impstrArrServerValues;

		// HCM Comments  - lotus Notes variables commented.										 
        //Begin- Lotus notes variables
        //string dnCertifier, mailSystem, availableForDirSynch, messageStorage;
        //string checkPassword, passwordChangeInterval, passwordGracePeriod, securePassword;
        //string type, form;
        //string strCommonName = string.Empty;//Common name without NONLILLY at the end
        //End- Lotus notes requirements

        //Used for list of domains
        string strZone1Domains, strZone2Domains, strZone3Domains;

        //Used for list of secure email domains
        string strSecureEmailDomain;

        //Secure email array and valuecollection
        string[] arrSecureDomain;
        ValueCollection vcSecureDomainWildCard = Utils.ValueCollection("initialvalue");
        ValueCollection vcSecureDomainWithoutWildCard = Utils.ValueCollection("initialvalue");

        StreamWriter objStreamWriter;

        //Variable declaration for reading and maintaining PAC list
        StreamReader objStreamReaderRecPolicy;
        string inputPACList = string.Empty;
        string[] arrPACList;
        string strPACList = string.Empty;

        string[] arrNovartisList;

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
				 #region Read Config File and set variables										  
                const string XML_CONFIG_FILE = @"\rules-config.xml";
                XmlDocument config = new XmlDocument();
                string dir = Utils.ExtensionsDirectory;
                config.Load(dir + XML_CONFIG_FILE);

                rnode = config.SelectSingleNode("rules-extension-properties");
                node = rnode.SelectSingleNode("environment");
                env = node.InnerText;
                rnode = config.SelectSingleNode
                    ("rules-extension-properties/management-agents/" + env + "/ad-ma");
                XmlNode confignode = rnode.SelectSingleNode("siteconfigfile");
                afvalhash = LoadAFConfig(confignode.InnerText);
                //get the Users to be considered from config                       
                node = rnode.SelectSingleNode("provision");
                etypeNodeValue = node.InnerText.ToUpper();
                //get the Users to be excluded from setting LillyCollab tag in displayname
                node = rnode.SelectSingleNode("toggledisplayname");
                toggledisplayname = node.InnerText.ToUpper();

                //Exchange
                node = rnode.SelectSingleNode("toggledtag");
                toggledtag = node.InnerText;

                //CHG1471559 - Initial Password Update
                node = rnode.SelectSingleNode("MinInitPwdLength");
                MinInitPwdLength = Int32.Parse(node.InnerText);

                //CHG1471559 - Initial Password Update
                node = rnode.SelectSingleNode("MaxInitPwdLength");
                MaxInitPwdLength = Int32.Parse(node.InnerText);

				// **************HCM Comments- Retiring EDS MA .Commented below code
				#region EDS Retirement
				/*
                rnode = config.SelectSingleNode
                    ("rules-extension-properties/management-agents/" + env + "/eds-ma");
                node = rnode.SelectSingleNode("/eds-ma");
                //node = rnode.SelectSingleNode("accountou");
                //accountOu = node.InnerText;
                node = rnode.SelectSingleNode("workerou");
                workerOu = node.InnerText;
				*/
				# endregion

                //Exchange email prefix (important for non-prd env)
                rnode = config.SelectSingleNode
                    ("rules-extension-properties/management-agents/" + env + "/Exchange");
                node = rnode.SelectSingleNode("exchangemailprefix");
                strexchangemailprefix = node.InnerText;

                node = rnode.SelectSingleNode("exchtransitiontime");
                strexchtransitiontime = node.InnerText;

                node = rnode.SelectSingleNode("Transitionlogsloc");
                strexchtransitionlogs = node.InnerText;

                node = rnode.SelectSingleNode("WhereToProvision");
                strGlobalPAC = node.InnerText;

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
             

				#endregion 	   
               //Reading List of secure email domains required to validate mail contact users for valid forwarders
                StreamReader objStreamReader;
                string input;


				 // SecureDomain file is populated by separate service that pulls data from TPET.
                // (HCM Comments) - Need to read filename from config													 
                objStreamReader = File.OpenText("C:\\Program Files\\Microsoft Forefront Identity Manager\\2010\\Synchronization Service\\Extensions\\SecureDomains.txt");

                while ((input = objStreamReader.ReadLine()) != null)
                {
                    strSecureEmailDomain = strSecureEmailDomain + input;
                }

                //
				#region Novartis Logic
                //***************************************************************************
                // (HCM Comments) Novartis Migration Logic - Can be removed														   
                StreamReader objStreamReaderNovartisGroup;
                string inputNovartisGroup;
                string strNovartisMigratedUser = string.Empty;

				// (HCM Comments) Novartis Migration Logic - Need to read filename from config																	  
                objStreamReaderNovartisGroup = File.OpenText("C:\\Program Files\\Microsoft Forefront Identity Manager\\2010\\Synchronization Service\\Extensions\\NovartisMigratedUser.txt");

                while ((inputNovartisGroup = objStreamReaderNovartisGroup.ReadLine()) != null)
                {
                    strNovartisMigratedUser = strNovartisMigratedUser + ";" + inputNovartisGroup.TrimStart().TrimEnd();
                }
                if ((strNovartisMigratedUser == string.Empty) || (strNovartisMigratedUser == null))
                    //throw new Exception("Exchange Accepted Domain is empty");
                    arrNovartisList = null;
                else
                {
                    arrNovartisList = strNovartisMigratedUser.Split(';');
                }

				 #endregion	  

                
				#region read secure domains into array

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
				#endregion

                #region Read PAC codes
                //Read PAC List 
				// (HCM Comments) - Need to read filename from config												 
                objStreamReaderRecPolicy = File.OpenText("C:\\Program Files\\Microsoft Forefront Identity Manager\\2010\\Synchronization Service\\Extensions\\PACList.txt");

                while ((inputPACList = objStreamReaderRecPolicy.ReadLine()) != null)
                {
                    strPACList = strPACList + ";" + inputPACList.TrimStart().TrimEnd();
                }
                if ((strPACList == string.Empty) || (strPACList == null))
                    //throw new Exception("Exchange Accepted Domain is empty");
                    arrPACList = null;
                else
                {
                    arrPACList = strPACList.Split(';');
                }
				#endregion		  
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
            //BOC NLW mini Realease
            MVObjectType = "Person";
            if (csentry["SYSTEM_ACCESS_FLG"].IsPresent && csentry["SYSTEM_ACCESS_FLG"].Value.ToUpper() == "N"
                && csentry["UNIFIED_COMM_FLG"].IsPresent && csentry["UNIFIED_COMM_FLG"].Value.ToUpper() == "Y")
            {
                string sSource, sLog, sEvent;

                sSource = "MIIS SAD MA";
                sLog = "Application";
                sEvent = csentry["SYSTEM_ID"].Value + " Povisioning is not to be done for SA Flag=N and UC Flag=Y";

                if (!EventLog.SourceExists(sSource))
                    EventLog.CreateEventSource(sSource, sLog);

                EventLog.WriteEntry(sSource, sEvent, EventLogEntryType.Error, 8161);

                return false;
            }
            else
                return true;
            //EOC NLW mini Realease

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
            ///
            ///Import attribute flow.
            ///

            string sSource, sLog, sEvent;

            switch (FlowRuleName)
            {//List of Import attributes is in alphabetical order

                // New Business Area Code Logic added for Elanco Release
                // Modifyng Business Area CODE VALUE BAESD ON Elanco Flag
                #region Business_Area_CD

                //#CHG1324643 - Added below code to update Buisness Area Code Based on Elanco Flag
                //HCM - Phase1 - Removing BUSINESS_AREA_CD from the flow for IDM SAD mapping updates
                //HCM - TBD - remove BUSINESS_AREA_CD from flow
                //case "cd.person:BUSINESS_AREA_CD,ELANCO_FLAG,SYSTEM_ID->mv.person:Business_Area_CD":
                case "cd.person:ELANCO_FLAG,SYSTEM_ID->mv.person:Business_Area_CD":
                    // if (csentry["BUSINESS_AREA_CD"].IsPresent)
                    mventry["Business_Area_CD"].Value = GetBusinessAreaCode(csentry);
                    break;

                #endregion

                #region AssistantPrsnlNbr
				 //HCM - Phase1 - Removing ADMIN_ASST_PRSNL_NBR from the flow for IDM SAD mapping updates
                //HCM - remove ADMIN_ASST_PRSNL_NBR from Sync flow
				/*
                case "cd.person:ADMIN_ASST_PRSNL_NBR->mv.person:AssistantPrsnlNbr":
                    if (csentry["ADMIN_ASST_PRSNL_NBR"].IsPresent)
                    {
                        mventry["AssistantPrsnlNbr"].Value = csentry["ADMIN_ASST_PRSNL_NBR"].Value.ToString();
                    }
                    break;
				*/
                #endregion

                #region cn
                case "cd.person:ANGLCZD_FIRST_NM,ANGLCZD_LAST_NM,ANGLCZD_MDL_NM,FRST_NM,LST_NM,MDL_NM,PRSNL_NBR->mv.person:cn":
                    //Construct Unique CN and change it to upper case before setting to MV
                    //Constructing Unique CN 
                    //Delete if CN null - But CN cannot be null
                    //and Write to Event log for CN null is in the ConstructUniqueCommonName() method
                    string uCN = ConstructUniqueCommonName(csentry, mventry);
                    if (!uCN.Equals(""))
                    {
                        mventry["cn"].Value = uCN.ToUpper().Trim();
                    }
                    else
                    {
                        mventry["cn"].Delete();
                    }
                    break;
                #endregion

                #region countryCode
                case "cd.person:CNTRY_CD_NBR->mv.person:countryCode":
                    //Code to convert the string to integer, 
                    //Not deleted - SAD SLA will gaurentee that the Values will there either 0 or the country code
                    mventry["countryCode"].IntegerValue = Convert.ToInt32(csentry["CNTRY_CD_NBR"].Value);
                    break;
                #endregion

                #region displayName
                case "cd.person:EMPLY_GRP_CD,FRST_NM,LST_NM,MDL_NM,PRSNL_NBR->mv.person:displayName":
                    //Displayname constructed and appeneded with " - LillyCollab" if the users are collab
                    //R3 - Use complete list of exclusion from config file, also manages null & appends LillyCollab for null emply_grp_cd
                    //Not deleted, if value is emplty - instead error thrown

                   // CultureInfo cultureInfo = Thread.CurrentThread.CurrentCulture;
                   // TextInfo textInfo = cultureInfo.TextInfo;                    

					 //HCM Comment - If Legal Name is not present use First, Middle and last name from SAD to create the display name																												
                    if ((!mventry["LegalNameFlag"].IsPresent) || (mventry["LegalNameFlag"].IsPresent && mventry["LegalNameFlag"].Value.ToString() == "0"))
                    {
                        string displayname = "";
                        if (csentry["FRST_NM"].IsPresent)
                        {
                            displayname = csentry["FRST_NM"].Value;
                           // displayname = textInfo.ToTitleCase(csentry["FRST_NM"].Value.ToString().ToLower());
                        }
                        if (csentry["MDL_NM"].IsPresent)
                        {
                            displayname = displayname.Trim() + " " + csentry["MDL_NM"].Value;
                            //displayname = displayname.Trim() + " " + textInfo.ToTitleCase(csentry["MDL_NM"].Value.ToString().ToLower());                            
                        }
                        if (csentry["LST_NM"].IsPresent)
                        {
                            displayname = displayname.Trim() + " " + csentry["LST_NM"].Value;
                            //displayname = displayname.Trim() + " " + textInfo.ToTitleCase(csentry["LST_NM"].Value.ToString().ToLower());
                        }
                        if (csentry["EMPLY_GRP_CD"].IsPresent
                            && !toggledisplayname.Contains(csentry["EMPLY_GRP_CD"].Value.ToUpper()))//If user is not in Employee group, append -network
                        {
                            displayname = displayname.Trim() + toggledtag;
                        }
                        else if (!csentry["EMPLY_GRP_CD"].IsPresent)//If Employee group code is not present, append -network
                        {
                            displayname = displayname.Trim() + toggledtag;
                        }

                        if (!displayname.Trim().Equals(""))//If display name generated is not empty then set it in MV
                        {
                            mventry["displayName"].Value = displayname.Trim();
                        }
                        else
                        {
                            mventry["displayName"].Delete();
                            //Write to the Event log for displayname is empty
                            sSource = "SAD MA";
                            sLog = "Application";
                            sEvent = csentry["PRSNL_NBR"].Value + " - Displayname is null";

                            if (!EventLog.SourceExists(sSource))
                                EventLog.CreateEventSource(sSource, sLog);

                            EventLog.WriteEntry(sSource, sEvent, EventLogEntryType.Error, 8102);
                        }
                    }
                    break;
                #endregion

                #region domain
                case "cd.person:EMPLY_GRP_CD,PRSNL_AREA_CD,PRSNL_NBR->mv.person:domain":
                    //domain info set for the user translated out of XML file
                    //Delete if value is null
                    //PRSNL_NBR is added to trigger the code in case the original attribute value is null
                    if (csentry["PRSNL_AREA_CD"].IsPresent && csentry["EMPLY_GRP_CD"].IsPresent)
                    {
                        //XML Method to get the values from hastable (filled from config xml)
                        setUserVariables(csentry["PRSNL_AREA_CD"].Value, csentry["EMPLY_GRP_CD"].Value);
						//Set Temp domain value = default domain set in above function from config																		  
                        string tempdomain = domain;

                        if ((tempdomain == null) || (tempdomain.ToString().Trim().Length == 0)) //TT09431320 - When domain is empty string it should be deleted from Metaverse
                        {
                            //Don't do anything just delete metaverse entry 
                            //Enter the log in event viewer (moved up - TT10007383)
                            string ExceptionMessage = csentry["PRSNL_NBR"].Value
                                        + " - PACCODE and EMPLOYEETYPE combination cannot retrieve data from XML file";
                            sSource = "SAD MA";
                            sLog = "Application";
                            sEvent = ExceptionMessage;

                            if (!EventLog.SourceExists(sSource))
                                EventLog.CreateEventSource(sSource, sLog);

                            EventLog.WriteEntry(sSource, sEvent, EventLogEntryType.Error, 8031);
                            // If a tag does not exist in the xml, then the stopped-extension-dll 
                            // error will be thrown.
                            mventry["domain"].Delete();
                            throw new ArgumentException("Invalid PAC"); //Exception thrown so that the user is not created in MV
                            //break; // Commented since it would never be executed
                        }
                        else if (tempdomain != null)
                        {
                            mventry["domain"].Value = tempdomain;
                        }
                    }
                    break;
                    #endregion

                  //  HCM Comments-Retiring EDS MA .commented below code. TBD  - Delete MV Flow.											  
                #region EDSWorkerDN - Retired MA
				/*
                case "cd.person:PRSNL_NBR->mv.person:EDSWorkerDN":
                    //Build EDSWorkerDN
                    mventry["EDSWorkerDN"].Value = "employeenumber=" + csentry["PRSNL_NBR"].Value.Trim() + "," + workerOu;
                    break;
				*/
                #endregion

                //HCM Comments- Retiring EDS MA .commented below code. TBD - Delete MV Flow.									  
                #region EDScn - Retired MA
				/*
                case "cd.person:FRST_NM,LST_NM,MDL_NM,PRSNL_NBR->mv.person:EDScn":
                    //Build EDScn
                    string cn = "";
                    if (csentry["FRST_NM"].IsPresent)
                        cn = csentry["FRST_NM"].Value;
                    if (csentry["MDL_NM"].IsPresent)
                        cn = cn.Trim() + " " + csentry["MDL_NM"].Value;
                    if (csentry["LST_NM"].IsPresent)
                        cn = cn.Trim() + " " + csentry["LST_NM"].Value;

                    if (!cn.Trim().Equals(""))
                    {
                        mventry["EDScn"].Value = cn.Trim();
                    }
                    break;
				*/
                #endregion

                //HCM Comments- Retiring EDS MA .commented below code	. TBD - Delete MV Flow.	  
                #region EDS_Supervisor_flag - Retired MA
				/*
                case "cd.person:SUPERVISOR_FLG->mv.person:EDS_Supervisor_flag":
                    if (csentry["SUPERVISOR_FLG"].Value == "0")
                        mventry["EDS_Supervisor_flag"].Value = "N";
                    else if (csentry["SUPERVISOR_FLG"].Value == "1")
                        mventry["EDS_Supervisor_flag"].Value = "Y";
                    break;
                 */
                #endregion

                #region emailTemplate
                //This attribute flow is for Password dissimination functionality , emailTemplate from Environment config file  is passed to Notification table     
                case "cd.person:PRSNL_NBR->mv.person:emailTemplate":
                    //if (emailtemplate.ToString().Trim().Length > 0) - Code changed as this would throw error in case of empty string
                    if (!((emailtemplate == null) || (emailtemplate.ToString().Trim().Length == 0)))
                        mventry["emailTemplate"].Value = emailtemplate;
                    break;
                #endregion

                #region Exch_Trans_CD
                //#region Exch_Trans_CD
                //Code to determine exchange code 
                //Exch_Trans_CD values 
                //0 - None (Do nothing)
                //1 - Create Mailbox (after tipping point) 
                //2 - Maintain Mailbox
                //3 - Unhide Mailbox
                //4 - Create Contact (after tipping point) - SAD Auth and Ext Contact
                //5 - Maintain Contact 
                //6 - Clear Mailbox properties
                //7 - Mailbox to Contact (day 0)
                //8 - Maintain Contact - after transition from mailbox (within 7 days)
                //9 - Clear all mail properties
                //10 - Mailbox to none (day 0) - None
                case "cd.person:BUSINESS_EMAIL_TXT,EMPLY_GRP_CD,EXT_AUTH_FLG,INTERNET_STYLE_ADRS,LAST_UPDT_DT,NOVARTIS_FLG,PRSNL_AREA_CD,PRSNL_NBR,SYSTEM_ACCESS_FLG,SYSTEM_ID,UNIFIED_COMM_FLG->mv.person:Exch_Trans_CD":

                    //User Type logic
                    if (csentry["PRSNL_AREA_CD"].IsPresent && csentry["EMPLY_GRP_CD"].IsPresent)
                    {
                        setUserVariables(csentry["PRSNL_AREA_CD"].Value, csentry["EMPLY_GRP_CD"].Value);
                        long intmsExchRTD = 0;
                        if (mventry["msExchRecipientTypeDetails"].IsPresent)
                        {
                            intmsExchRTD = Convert.ToInt64(mventry["msExchRecipientTypeDetails"].Value);
                        }

                        switch (CommonLayer.GetUserType(mventry, csentry, "SAD", "no"))
                        {
                            case CommonLayer.GUT_MAIL_CONTACT:
                            case CommonLayer.GUT_MADS:
                                if (intmsExchRTD == 0)
                                    mventry["Exch_Trans_CD"].Value = "4";//Create Contact
                                else if (intmsExchRTD == Convert.ToInt32(CommonLayer.MAIL_USER) || intmsExchRTD == Convert.ToInt32(CommonLayer.MAIL_CONTACT))
                                    mventry["Exch_Trans_CD"].Value = "5"; //Maintain Contact
                                else if (intmsExchRTD == Convert.ToInt32(CommonLayer.LINKED_MAILBOX) || intmsExchRTD == Convert.ToInt64(CommonLayer.REMOTE_MAILBOX))
                                {
                                    #region targetaddress is present
                                    if (mventry["targetaddress"].IsPresent)
                                    {
                                        if (mventry["msExchtransitionDate"].IsPresent)
                                        {
                                            if (AccountTTLExpired(mventry["msExchtransitionDate"].Value, strexchtransitiontime))
                                            {
                                                if (!(mventry["msExchUMEnabledFlags"].IsPresent))
                                                {
                                                    mventry["Exch_Trans_CD"].Value = "6";
                                                    LogExchangeAttributes(csentry, mventry);
                                                }
                                                else
                                                {
                                                    //Log an event in eventviewer

                                                    sSource = "SAD MA";
                                                    sLog = "Application";
                                                    sEvent = mventry["sAMAccountName"].Value + " - Mailbox can't be deprovisioned because UM was not successfully disabled for the Account";

                                                    if (!EventLog.SourceExists(sSource))
                                                        EventLog.CreateEventSource(sSource, sLog);

                                                    EventLog.WriteEntry(sSource, sEvent, EventLogEntryType.Error, 8191);
                                                }
                                            }
                                            else
                                            {
                                                mventry["Exch_Trans_CD"].Value = "8";
                                            }
                                        }
                                    }
                                    #endregion
                                    #region targetaddress is not present
                                    else
                                    {
                                        mventry["Exch_Trans_CD"].Value = "7";

                                    }
                                    #endregion
                                }
                                break;

                            case CommonLayer.GUT_LOCAL_MBX:
                            case CommonLayer.GUT_REMOTE_MBX:
                                if (csentry["SYSTEM_ACCESS_FLG"].IsPresent && csentry["SYSTEM_ACCESS_FLG"].Value.ToUpper() == "Y")
                                {
                                    if (!mventry["msExchMailboxGuid"].IsPresent)
                                    {
                                        mventry["Exch_Trans_CD"].Value = "1";//Create Mailbox
                                    }
                                    else
                                    {
                                        if (mventry["msExchHideFromAddressList"].IsPresent && mventry["msExchHideFromAddressList"].BooleanValue == true)
                                        {
                                            mventry["Exch_Trans_CD"].Value = "3";
                                        }
                                        else
                                        {
                                            mventry["Exch_Trans_CD"].Value = "2";//Maintain mailbox
                                        }
                                    }
                                }
                                else
                                {
                                    if (mventry["msExchMailboxGuid"].IsPresent)
                                        mventry["Exch_Trans_CD"].Value = "10"; //Mailbox to none
                                    else
                                        mventry["Exch_Trans_CD"].Value = "0"; //None
                                }
                                break;

                            case CommonLayer.GUT_None:
                                if (intmsExchRTD == 0)
                                    mventry["Exch_Trans_CD"].Value = "0";//None
                                else if (intmsExchRTD == 128 || intmsExchRTD == 32768)
                                    mventry["Exch_Trans_CD"].Value = "9";//Clear all mail properties
                                else if (intmsExchRTD == 2 || intmsExchRTD == 1 || intmsExchRTD == 2147483648)
                                {
                                    if (mventry["msExchHideFromAddressList"].IsPresent)
                                    {
                                        if (mventry["msExchHideFromAddressList"].BooleanValue == false)
                                        {
                                            mventry["Exch_Trans_CD"].Value = "10";//Mailbox to none (day 0)
                                        }
                                        else
                                        {
                                            if (mventry["msExchtransitionDate"].IsPresent)
                                            {
                                                if (AccountTTLExpired(mventry["msExchtransitionDate"].Value, strexchtransitiontime))
                                                {
                                                    if (!(mventry["msExchUMEnabledFlags"].IsPresent))
                                                    {
                                                        mventry["Exch_Trans_CD"].Value = "9";//Clear all mail properties
                                                        LogExchangeAttributes(csentry, mventry);
                                                    }
                                                    else
                                                    {
                                                        //Log an event in eventviewer

                                                        sSource = "SAD MA";
                                                        sLog = "Application";
                                                        sEvent = mventry["sAMAccountName"].Value + " - Mailbox can't be deprovisioned because UM was not successfully disabled for the Account";

                                                        if (!EventLog.SourceExists(sSource))
                                                            EventLog.CreateEventSource(sSource, sLog);

                                                        EventLog.WriteEntry(sSource, sEvent, EventLogEntryType.Error, 8191);
                                                    }
                                                }

                                            }
                                        }
                                    }
                                    else
                                    {
                                        mventry["Exch_Trans_CD"].Value = "10";//Mailbox to none (day 0)
                                    }
                                }
                                break;
                        }
                    }
                    break;
                #endregion

                #region mailNickname
				// HCM - Phase1 - mailnickName new code logic
                //Code to generate unique mailNickname 
                //Exch_Trans_CD values 
                //0 - None (Do nothing)
                //1 - Create Mailbox (after tipping point) - Generate MailNickname
                //2 - Maintain Mailbox
                //3 - Unhide Mailbox
                //4 - Create Contact (after tipping point) - SAD Auth and Ext Contact.- Generate MailNickname
                //5 - Maintain Contact 
                //6 - Clear Mailbox properties
                //7 - Mailbox to Contact (day 0) - Generate MailNickname
                //8 - Maintain Contact - after transition from mailbox (within 7 days)
                //9 - Clear all mail properties
                //10 - Mailbox to none (day 0) - None - Generate MailNickname
                case "cd.person:ANGLCZD_FIRST_NM,ANGLCZD_LAST_NM,ANGLCZD_MDL_NM,BUSINESS_EMAIL_TXT,EMPLY_GRP_CD,EXT_AUTH_FLG,INTERNET_STYLE_ADRS,LAST_UPDT_DT,NOVARTIS_FLG,PRSNL_AREA_CD,PRSNL_NBR,SYSTEM_ACCESS_FLG,SYSTEM_ID,UNIFIED_COMM_FLG->mv.person:mailNickname":

                   // if (!mventry["mailNickname"].IsPresent) // commneted the code for HCM Phase1 Bug Fix 
                   // {
                        //User Type logic
                        if (csentry["PRSNL_AREA_CD"].IsPresent && csentry["EMPLY_GRP_CD"].IsPresent)
                        {
                            //setUserVariables(csentry["PRSNL_AREA_CD"].Value, csentry["EMPLY_GRP_CD"].Value);
                            long intmsExchRTD = 0;
                            if (mventry["msExchRecipientTypeDetails"].IsPresent)
                            {
                                intmsExchRTD = Convert.ToInt64(mventry["msExchRecipientTypeDetails"].Value);
                            }

                            switch (CommonLayer.GetUserType(mventry, csentry, "SAD", "no"))
                            {
                                case CommonLayer.GUT_MAIL_CONTACT:
                                case CommonLayer.GUT_MADS:
                                    if (intmsExchRTD == 0)
                                    {
                                        //mventry["Exch_Trans_CD"].Value = "4";//Create Contact

                                        //HCM - Creating Unique MailNickname for new Mailcontacts.
                                        if (!mventry["mailNickname"].IsPresent)
                                        {
                                            string defaultmNickname = CommonLayer.BuildMailNickName(csentry);
                                            if (defaultmNickname != string.Empty)
                                            {
                                                mventry["mailNickname"].Value = CommonLayer.GetUniqueMailNickname(defaultmNickname.ToLower(), csentry["SYSTEM_ID"].Value.ToLower(), mventry);
                                            }
                                        }
                                    }
                                    /*
                                     * else if (intmsExchRTD == Convert.ToInt32(CommonLayer.MAIL_USER) || intmsExchRTD == Convert.ToInt32(CommonLayer.MAIL_CONTACT))
                                      //  {
                                       //     mventry["Exch_Trans_CD"].Value = "5"; //Maintain Contact
                                       }
                                   */
                                    else if (intmsExchRTD == Convert.ToInt32(CommonLayer.LINKED_MAILBOX) || intmsExchRTD == Convert.ToInt64(CommonLayer.REMOTE_MAILBOX))
                                    {
                                        #region targetaddress is present
                                        /*
                                        if (mventry["targetaddress"].IsPresent)
                                        {
                                            if (mventry["msExchtransitionDate"].IsPresent)
                                            {
                                                if (AccountTTLExpired(mventry["msExchtransitionDate"].Value, strexchtransitiontime))
                                                {
                                                    if (!(mventry["msExchUMEnabledFlags"].IsPresent))
                                                    {
                                                        mventry["Exch_Trans_CD"].Value = "6";
                                                        LogExchangeAttributes(csentry, mventry);
                                                    }
                                                    else
                                                    {
                                                        //Log an event in eventviewer

                                                        sSource = "SAD MA";
                                                        sLog = "Application";
                                                        sEvent = mventry["sAMAccountName"].Value + " - Mailbox can't be deprovisioned because UM was not successfully disabled for the Account";

                                                        if (!EventLog.SourceExists(sSource))
                                                            EventLog.CreateEventSource(sSource, sLog);

                                                        EventLog.WriteEntry(sSource, sEvent, EventLogEntryType.Error, 8191);
                                                    }
                                                }
                                                else
                                                {
                                                    mventry["Exch_Trans_CD"].Value = "8";
                                                }
                                            }
                                        }
                                         */
                                        #endregion

                                        #region targetaddress is not present
                                        //else
                                        if (!mventry["targetaddress"].IsPresent)
                                        {
                                            //mventry["Exch_Trans_CD"].Value = "7";
                                            //HCM - Creating Unique MailNickname 
                                            if (!mventry["mailNickname"].IsPresent)
                                            {
                                                string defaultmNickname = CommonLayer.BuildMailNickName(csentry);
                                                if (defaultmNickname != string.Empty)
                                                {
                                                    mventry["mailNickname"].Value = CommonLayer.GetUniqueMailNickname(defaultmNickname.ToLower(), csentry["SYSTEM_ID"].Value.ToLower(), mventry);
                                                }
                                            }

                                        }
                                        #endregion
                                    }

                                    break;

                                case CommonLayer.GUT_LOCAL_MBX:
                                case CommonLayer.GUT_REMOTE_MBX:
                                    if (csentry["SYSTEM_ACCESS_FLG"].IsPresent && csentry["SYSTEM_ACCESS_FLG"].Value.ToUpper() == "Y")
                                    {
                                        if (!mventry["msExchMailboxGuid"].IsPresent)
                                        {
                                            //mventry["Exch_Trans_CD"].Value = "1";//Create Mailbox
                                            //HCM - Creating Unique MailNickname for new Mailcontacts.
                                            if (!mventry["mailNickname"].IsPresent)
                                            {
                                                string defaultmNickname = CommonLayer.BuildMailNickName(csentry);
                                                if (defaultmNickname != string.Empty)
                                                {
                                                    mventry["mailNickname"].Value = CommonLayer.GetUniqueMailNickname(defaultmNickname.ToLower(), csentry["SYSTEM_ID"].Value.ToLower(),mventry);
                                                }
                                            }
                                        }
                                        /*else
                                        {
                                            if (mventry["msExchHideFromAddressList"].IsPresent && mventry["msExchHideFromAddressList"].BooleanValue == true)
                                            {
                                                mventry["Exch_Trans_CD"].Value = "3";
                                            }
                                            else
                                            {
                                                mventry["Exch_Trans_CD"].Value = "2";//Maintain mailbox
                                            }
                                        }*/
                                    }
                                    else
                                    {
                                        if (mventry["msExchMailboxGuid"].IsPresent)
                                        {
                                            // mventry["Exch_Trans_CD"].Value = "10"; //Mailbox to none

                                            //HCM - Creating Unique MailNickname
                                            if (!mventry["mailNickname"].IsPresent)
                                            {
                                                string defaultmNickname = CommonLayer.BuildMailNickName(csentry);
                                                if (defaultmNickname != string.Empty)
                                                {
                                                    mventry["mailNickname"].Value = CommonLayer.GetUniqueMailNickname(defaultmNickname.ToLower(), csentry["SYSTEM_ID"].Value.ToLower(), mventry);
                                                }
                                            }
                                        }
                                        else
                                        {
                                            //mventry["Exch_Trans_CD"].Value = "0"; //None

                                            //HCM - Clear MailNickname for NONE type users
                                            if (mventry["mailNickname"].IsPresent)
                                            {
                                                mventry["mailNickname"].Delete();
                                            }
                                        }
                                    }
                                    break;

                                case CommonLayer.GUT_None:
                                    if (intmsExchRTD == 0)
                                    {
                                        //mventry["Exch_Trans_CD"].Value = "0";//None
                                        //HCM - Deleting MailNickname
                                        if (mventry["mailNickname"].IsPresent)
                                        {
                                            mventry["mailNickname"].Delete();
                                        }
                                    }
                                    else if (intmsExchRTD == 128 || intmsExchRTD == 32768)
                                    {
                                        // mventry["Exch_Trans_CD"].Value = "9";//Clear all mail properties

                                        //HCM - Deleting MailNickname
                                        if (mventry["mailNickname"].IsPresent)
                                        {
                                            mventry["mailNickname"].Delete();
                                        }
                                    }
                                    else if (intmsExchRTD == 2 || intmsExchRTD == 1 || intmsExchRTD == 2147483648)
                                    {
                                        if (mventry["msExchHideFromAddressList"].IsPresent)
                                        {
                                            if (mventry["msExchHideFromAddressList"].BooleanValue == false)
                                            {
                                                //  mventry["Exch_Trans_CD"].Value = "10";//Mailbox to none (day 0)

                                                //HCM - Creating Unique MailNickname
                                                if (!mventry["mailNickname"].IsPresent)
                                                {
                                                    string defaultmNickname = CommonLayer.BuildMailNickName(csentry);
                                                    if (defaultmNickname != string.Empty)
                                                    {
                                                        mventry["mailNickname"].Value = CommonLayer.GetUniqueMailNickname(defaultmNickname.ToLower(), csentry["SYSTEM_ID"].Value.ToLower(), mventry);
                                                    }
                                                }
                                            }
                                            else// if msExchHideFromAddressList is true
                                            {
                                                if (mventry["msExchtransitionDate"].IsPresent)
                                                {
                                                    if (AccountTTLExpired(mventry["msExchtransitionDate"].Value, strexchtransitiontime))
                                                    {
                                                        if (!(mventry["msExchUMEnabledFlags"].IsPresent))
                                                        {
                                                            // mventry["Exch_Trans_CD"].Value = "9";//Clear all mail properties

                                                            //HCM - Clear MailNickname for NONE type users
                                                            if (mventry["mailNickname"].IsPresent)
                                                            {
                                                                mventry["mailNickname"].Delete();
                                                            }

                                                            // LogExchangeAttributes(csentry, mventry);
                                                        }
                                                        /*
                                                        else
                                                        {
                                                            //Log an event in eventviewer

                                                            sSource = "SAD MA";
                                                            sLog = "Application";
                                                            sEvent = mventry["sAMAccountName"].Value + " - Mailbox can't be deprovisioned because UM was not successfully disabled for the Account";

                                                            if (!EventLog.SourceExists(sSource))
                                                                EventLog.CreateEventSource(sSource, sLog);

                                                            EventLog.WriteEntry(sSource, sEvent, EventLogEntryType.Error, 8191);
                                                        }
                                                        */
                                                    }

                                                }
                                            }
                                        }
                                        else // if msExchHideFromAddressList is not present
                                        {
                                            //  mventry["Exch_Trans_CD"].Value = "10";//Mailbox to none (day 0)

                                            //HCM - Creating Unique MailNickname
                                            if (!mventry["mailNickname"].IsPresent)
                                            {
                                                string defaultmNickname = CommonLayer.BuildMailNickName(csentry);
                                                if (defaultmNickname != string.Empty)
                                                {
                                                    mventry["mailNickname"].Value = CommonLayer.GetUniqueMailNickname(defaultmNickname.ToLower(), csentry["SYSTEM_ID"].Value.ToLower(), mventry);
                                                }
                                            }
                                        }
                                    }
                                    break;
                            }
                        }
                    //}
                    break;
                #endregion

                // HCM Comments - TBD for deletion									  
                #region Ext_Auth_Flag
                case "cd.person:EXT_AUTH_FLG,LAST_UPDT_DT,NOVARTIS_FLG,SYSTEM_ID,UNIFIED_COMM_FLG->mv.person:Ext_Auth_Flag":
                    if (CommonLayer.IsMigrationOverride(mventry, csentry, "SAD") || CommonLayer.IsMigrationAcquisition(mventry, csentry, "SAD"))
                    {
                        mventry["Ext_Auth_Flag"].Value = "N";
                    }
                    else
                    {
                        if (csentry["NOVARTIS_FLG"].IsPresent && csentry["NOVARTIS_FLG"].Value.ToUpper() == "Y")
                        {
                            if (mventry["msExchMailboxGuid"].IsPresent)
                            {
                                mventry["Ext_Auth_Flag"].Value = "N";
                            }
                            else
                            {
                                if (!mventry["Ext_Auth_Flag"].IsPresent)
                                    mventry["Ext_Auth_Flag"].Value = csentry["EXT_AUTH_FLG"].Value;
                            }
                        }
                        else
                        {
                            mventry["Ext_Auth_Flag"].Value = csentry["EXT_AUTH_FLG"].Value;
                        }
                    }

                    break;
                #endregion

                #region External_Email_Address
                case "cd.person:BUSINESS_EMAIL_TXT,PRSNL_NBR->mv.person:External_Email_Address":
                    if (csentry["BUSINESS_EMAIL_TXT"].IsPresent && !String.IsNullOrEmpty(csentry["BUSINESS_EMAIL_TXT"].Value))
                    {
                        mventry["External_Email_Address"].Value = csentry["BUSINESS_EMAIL_TXT"].Value.ToString();
                    }
                    else if (mventry["IsBusinessEmailFromSAD"].IsPresent && mventry["IsBusinessEmailFromSAD"].BooleanValue == true)
                    {//Clear out the External Email Address
                        mventry["External_Email_Address"].Delete();
                    }
                    break;
                #endregion

                #region GetUserType
                case "cd.person:BUSINESS_EMAIL_TXT,EMPLY_GRP_CD,EXT_AUTH_FLG,INTERNET_STYLE_ADRS,LAST_UPDT_DT,NOVARTIS_FLG,PRSNL_AREA_CD,PRSNL_NBR,SYSTEM_ID,UNIFIED_COMM_FLG->mv.person:GetUserType":

                    setUserVariables(csentry["PRSNL_AREA_CD"].Value, csentry["EMPLY_GRP_CD"].Value);

                    mventry["GetUserType"].Value = CommonLayer.GetUserType(mventry, csentry, "SAD", "no");

                    break;
                #endregion

				//HCM - Phase 1 - Retiring KnownasMA
                #region givenName
                case "cd.person:FRST_NM,SYSTEM_ID->mv.person:givenName":
                    //if (!mventry["knownAsFirstName"].IsPresent && csentry["FRST_NM"].IsPresent)
					if(csentry["FRST_NM"].IsPresent)	
                    {
                        mventry["givenName"].Value = csentry["FRST_NM"].Value.ToString();
                    }
                    break;
                #endregion

				//HCM Comment - TBD for Novartis logic									  
                #region homeMDB
                case "cd.person:EMPLY_GRP_CD,NOVARTIS_FLG,PRSNL_AREA_CD,SYSTEM_ID,UNIFIED_COMM_FLG->mv.person:homeMDB":
                    // O365 bug fix - added additional condition to fix - Implement change in SAD MA to manage the flip logic of Novartis users. 

                    if ((csentry["UNIFIED_COMM_FLG"].IsPresent && csentry["UNIFIED_COMM_FLG"].Value.ToUpper() == "Y") || CommonLayer.IsMigrationAcquisition(mventry, csentry, "SAD"))
                    {
                        setUserVariables(csentry["PRSNL_AREA_CD"].Value.ToString(), csentry["EMPLY_GRP_CD"].Value.ToString());

                        if (!((strArrServerValues[1] == null) || (strArrServerValues[1] == "")))
                            mventry["homeMDB"].Value = strArrServerValues[1];
                    }
                    else
                    {
                        mventry["homeMDB"].Delete();
                    }

                    break;

                #endregion

				//HCM Comment - TBD for Novartis logic									  
                #region msExchHomeServerName
                case "cd.person:EMPLY_GRP_CD,NOVARTIS_FLG,PRSNL_AREA_CD,SYSTEM_ID,UNIFIED_COMM_FLG->mv.person:msExchHomeServerName":
                    // O365 bug fix - added additional condition to fix - Users who transition from contact to mailbox are not getting an msExchHomeServerName generated (they get homeMDB), so this is different than the Novartis logic. 

                    if ((csentry["UNIFIED_COMM_FLG"].IsPresent && csentry["UNIFIED_COMM_FLG"].Value.ToUpper() == "Y") || CommonLayer.IsMigrationAcquisition(mventry, csentry, "SAD"))
                    {
                        setUserVariables(csentry["PRSNL_AREA_CD"].Value.ToString(), csentry["EMPLY_GRP_CD"].Value.ToString());

                        if (!((strArrServerValues[0] == null) || (strArrServerValues[0] == "")))
                            mventry["msExchHomeServerName"].Value = strArrServerValues[0];
                    }
                    else
                    {
                        mventry["msExchHomeServerName"].Delete();
                    }

                    break;

                #endregion

                //CHG1471559 - Initial Password Update. Modifying initial password char limits for new AD password Policies
                //HCM - phase1 - Modified the code to regenerate init password for none to mailbox/mailcontact transitions
                #region initPassword
                //case "cd.person:PRSNL_NBR->mv.person:initPassword":
                case "cd.person:PRSNL_NBR->mv.person:initPassword":
                    //Initpassword set for all users only System Id changed and password not already present
                    //Not deleted - since logically this will execute only once during inital MV object creation
                    if (!mventry["initPassword"].IsPresent)
                    {
                        mventry["initPassword"].Value = RandomPassword.Generate(MinInitPwdLength, MaxInitPwdLength);
                    }
                    // HCM Phase1 - code commented- Password will be updated with separate MA.
                   // else if(mventry["initPassword"].IsPresent && (mventry["initPassword"].Value.Length < MinInitPwdLength) && mventry["initPassword"].Value != "CHANGED")
                  //  {
                      //  mventry["initPassword"].Value = RandomPassword.Generate(MinInitPwdLength, MaxInitPwdLength);
                  //  }
                    break;
                #endregion

                #region IsAcquisitionMigrate

				// (HCM Comments) - Novartis migration Logic - Can be removed															 
                case "cd.person:LAST_UPDT_DT,SYSTEM_ID->mv.person:IsAcquisitionMigrate":
                    mventry["IsAcquisitionMigrate"].BooleanValue = false;
                    if (csentry["SYSTEM_ID"].IsPresent)
                    {
                        for (int i = 0; i < arrNovartisList.Length; i++)
                        {
                            if (arrNovartisList[i].ToUpper() == csentry["SYSTEM_ID"].Value.ToUpper())
                            {
                                mventry["IsAcquisitionMigrate"].BooleanValue = true;
                                break;
                            }
                        }
                    }
                    break;
                #endregion

                #region IsBusinessEmailFromSAD
                case "cd.person:BUSINESS_EMAIL_TXT,PRSNL_NBR->mv.person:IsBusinessEmailFromSAD":
                    ////One time value set
                    if (csentry["BUSINESS_EMAIL_TXT"].IsPresent)//triple check
                    {
                        mventry["IsBusinessEmailFromSAD"].BooleanValue = true;
                    }
                    break;
                #endregion

                #region IsEmailDomainChange
                //HCM - Phase1 - Removing BUSINESS_AREA_CD from the flow for IDM SAD mapping updates
                //HCM - TBD to check the rules in Sync Flow
                // case "cd.person:BUSINESS_AREA_CD,EMPLY_GRP_CD,ORG_UNIT_CD,SYSTEM_ID->mv.person:IsEmailDomainChange":
                //HCM - No rules extension present for the code . Commenting
                /*
                case "cd.person:EMPLY_GRP_CD,ORG_UNIT_CD,SYSTEM_ID->mv.person:IsEmailDomainChange":

                    if (mventry["mail"].IsPresent)
                    {
                       
                        string strExistingEmailDomain = mventry["mail"].Value.ToLower().Split('@')[1];
                        if (strExistingEmailDomain == "lilly.com")
                        {//Lilly to Elanco (Employee)
                            if (GetBusinessAreaCode(csentry).ToString().ToUpper() == "K")
                                mventry["IsEmailDomainChange"].Value = "Y";
                        }
                        else if (strExistingEmailDomain == "network.lilly.com")
                        {//Lilly contractor to Elanco (Contractor) or Lilly employee
                            if (GetBusinessAreaCode(csentry).ToString().ToUpper() == "K" || csentry["EMPLY_GRP_CD"].Value.ToUpper() != "D")
                                mventry["IsEmailDomainChange"].Value = "Y";
                        }
                        else if (strExistingEmailDomain == "elanco.com")
                        {//Elanco employee to Lilly employee or elanco contractor
                            if (GetBusinessAreaCode(csentry).ToString().ToUpper() == "A" || csentry["EMPLY_GRP_CD"].Value.ToUpper() == "D")
                                mventry["IsEmailDomainChange"].Value = "Y";
                        }
                        else if (strExistingEmailDomain == "network.elanco.com")
                        {//Elanco contractor to Lilly employee or elanco employee
                            if (GetBusinessAreaCode(csentry).ToString().ToUpper() == "A" || csentry["EMPLY_GRP_CD"].Value.ToUpper() != "D")
                                mventry["IsEmailDomainChange"].Value = "Y";
                        }
                    }

                    break;
                    */
                #endregion

                #region newEmailDomain
                //HCM - Phase1 - Removing BUSINESS_AREA_CD from the flow for IDM SAD mapping updates
                //HCM - TBD to check the rules in Sync Flow
                //case "cd.person:BUSINESS_AREA_CD,EMPLY_GRP_CD,ORG_UNIT_CD,SYSTEM_ID->mv.person:newEmailDomain":
                //HCM - No rules extension present for the code . Commenting
                /*
                case "cd.person:EMPLY_GRP_CD,ORG_UNIT_CD,SYSTEM_ID->mv.person:newEmailDomain":

                    if (mventry["mail"].IsPresent)
                    {                       

                        string strExistingEmailDomain = mventry["mail"].Value.ToLower().Split('@')[1];
                        if (strExistingEmailDomain == "lilly.com")
                        {//Lilly to Elanco (Employee)
                            if (GetBusinessAreaCode(csentry).ToString().ToUpper() == "K")
                                mventry["newEmailDomain"].Value = "elanco.com";
                        }
                        else if (strExistingEmailDomain == "network.lilly.com")
                        {//Lilly contractor to Elanco (Contractor or employee) or Lilly employee
                            if (GetBusinessAreaCode(csentry).ToString().ToUpper() == "K")
                                if (csentry["EMPLY_GRP_CD"].Value.ToUpper() == "D")
                                    mventry["newEmailDomain"].Value = "network.elanco.com";
                                else
                                    mventry["newEmailDomain"].Value = "elanco.com";
                            if (csentry["EMPLY_GRP_CD"].Value.ToUpper() != "D")//Lilly employee
                                mventry["newEmailDomain"].Value = "lilly.com";
                        }
                        else if (strExistingEmailDomain == "elanco.com")
                        {//Elanco employee to Lilly employee
                            if (GetBusinessAreaCode(csentry).ToString().ToUpper() == "A")
                                mventry["newEmailDomain"].Value = "lilly.com";
                        }
                        else if (strExistingEmailDomain == "network.elanco.com")
                        {//Elanco contractor to Lilly contractor
                            if (GetBusinessAreaCode(csentry).ToString().ToUpper() == "A" && csentry["EMPLY_GRP_CD"].Value.ToUpper() == "D")
                                mventry["newEmailDomain"].Value = "network.lilly.com";
                        }
                    }

                    break;
                    */
                #endregion

                #region IsPhoneNumberChanged
                //BOC NLW mini Release
				//HCM - TBD verify if IsPhoneNumberChanged is being used anywhere else. if not delete it.
                case "cd.person:FULL_WORK_PHONE_NBR,LAST_UPDT_DT->mv.person:IsPhoneNumberChanged":
                    if (csentry["FULL_WORK_PHONE_NBR"].IsPresent && mventry["Full_Work_Phone_nbr"].IsPresent)
                    {
                        if (csentry["FULL_WORK_PHONE_NBR"].Value != mventry["Full_Work_Phone_nbr"].Value)
                            mventry["IsPhoneNumberChanged"].Value = "1";
                        else
                            mventry["IsPhoneNumberChanged"].Value = "0";
                    }
                    else
                        mventry["IsPhoneNumberChanged"].Value = "0";
                    break;

                //EOC NLW mini Release
                #endregion

                #region IsSamAccountNameChanged
                //Privileged Release code change - BOC
                case "cd.person:EMPLY_GRP_CD,LAST_UPDT_DT,SYSTEM_ID->mv.person:IsSamAccountNameChanged":
                    if (csentry["EMPLY_GRP_CD"].IsPresent && mventry["employeeType"].IsPresent && csentry["EMPLY_GRP_CD"].Value.ToUpper() != mventry["employeeType"].Value.ToUpper())
                    {
                        mventry["IsSamAccountNameChanged"].BooleanValue = true;
                    }
                    else
                    {
                        mventry["IsSamAccountNameChanged"].BooleanValue = false;
                    }
                    break;
                //Privileged Release code change - EOC
                #endregion

                #region lly_cntct_email_adrs
                //This attribute flow is for Password dissimination functionality , lilly contact email address is passed in Notification table 
                case "cd.person:LY_CONTACT_PRSNL_NBR->mv.person:lly_cntct_email_adrs":
                    string strLlyCnctAddr = GetEmailAddressFromID(mventry, csentry["LY_CONTACT_PRSNL_NBR"].Value);
                    if (strLlyCnctAddr == "")
                    {
                        mventry["lly_cntct_email_adrs"].Delete();
                    }
                    else
                    {
                        mventry["lly_cntct_email_adrs"].Value = strLlyCnctAddr;
                    }
                    break;
                #endregion

                #region lly_sprvsr_email_adrs
                //This attribute flow is for Password dissimination functionality , supervisor email address  is passed in Notification table     
                case "cd.person:SUPERVISOR_PRSNL_NBR->mv.person:lly_sprvsr_email_adrs":
                    string strSprvsrAddr = GetEmailAddressFromID(mventry, csentry["SUPERVISOR_PRSNL_NBR"].Value);
                    if (strSprvsrAddr == "")
                    {
                        mventry["lly_sprvsr_email_adrs"].Delete();
                    }
                    else
                    {
                        mventry["lly_sprvsr_email_adrs"].Value = strSprvsrAddr;
                    }
                    break;
                #endregion

                #region mailfromInternetAddress
                case "cd.person:EXT_AUTH_FLG,INTERNET_STYLE_ADRS,SYSTEM_ID->mv.person:mailfromInternetAddress":
                    if (csentry["EXT_AUTH_FLG"].Value.ToLower() == "y")
                    {
                        if ((csentry["INTERNET_STYLE_ADRS"].IsPresent) && (csentry["INTERNET_STYLE_ADRS"].Value.ToLower().Contains("@lilly")))
                        {
                            string[] strinternetaddress;
                            strinternetaddress = csentry["INTERNET_STYLE_ADRS"].Value.Split('@');
                            mventry["mailfromInternetAddress"].Value = strinternetaddress[0] + "@" + strexchangemailprefix;
                        }
                    }
                    break;
                #endregion

                #region ManagerPrsnlNbr
                case "cd.person:SUPERVISOR_PRSNL_NBR->mv.person:ManagerPrsnlNbr":
                    if (csentry["SUPERVISOR_PRSNL_NBR"].IsPresent)
                    {
                        mventry["ManagerPrsnlNbr"].Value = csentry["SUPERVISOR_PRSNL_NBR"].Value.ToString();
                    }

                    break;
                #endregion

				//HCM - Phase 1 - Retiring KnownasMA
                #region middleName
                case "cd.person:MDL_NM,SYSTEM_ID->mv.person:middleName":
                    //if (!mventry["knownAsMiddleName"].IsPresent && csentry["MDL_NM"].IsPresent)
					if (csentry["MDL_NM"].IsPresent)
                    {
                        mventry["middleName"].Value = csentry["MDL_NM"].Value.ToString();
                    }
                    else
                    {
                        mventry["middleName"].Delete();
                    }
                    break;
                //Release 2.2 changes -EOC 
                #endregion

                #region miis_mod_dt
                //Update the MIIS Modification date for a record

                //Change for AUDF1.5. Check applied on PAC change (cleanup release)

                //HCM - Phase1 - Removing EMPLY_SUB_GRP_CD and HAS_LILLY_ASSET_FLG from the flow for IDM SAD mapping updates
                //HCM - TBD - check the rules in Sync Flow
                case "cd.person:CNTRY_CD,EMPLY_GRP_CD,PRSNL_AREA_CD,PRSNL_NBR,STATUS_CD,SYSTEM_ID->mv.person:miis_mod_dt":
                    {
                        mventry["miis_mod_dt"].Value = System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"); //Stores Updated Modification date
                    }
                    break;
                #endregion

                #region ms-DS-ConsistencyGuid
                case "cd.person:SYSTEM_ID->mv.person:ms-DS-ConsistencyGuid":
                    //generate new GUID to set in Consistency Guid for cloud
                    if (!mventry["ms-DS-ConsistencyGuid"].IsPresent)
                        mventry["ms-DS-ConsistencyGuid"].BinaryValue = Guid.NewGuid().ToByteArray();
                    break;

                #endregion

                #region otherMailbox
                //This function has been introduced to test exchange functinality in Dev and QA
                //in Prod strexchangemailprefix value should be LILLY.COM
                //HCM - Phase1 - Removing COLLAB_EXTRNL_EMAIL_ADRS from the flow for IDM SAD mapping updates
                //HCM - TBD - check the rules in Sync Flow
                case "cd.person:BUSINESS_EMAIL_TXT,NOVARTIS_FLG,SYSTEM_ID->mv.person:otherMailbox":

                    string strOtherMailbox = string.Empty;
                    //Cannot use GetUserType function as Secure email check is not required here
                    if ((csentry["BUSINESS_EMAIL_TXT"].IsPresent))
                        strOtherMailbox = csentry["BUSINESS_EMAIL_TXT"].Value.ToString();
                    else if (mventry["IsBusinessEmailFromSAD"].IsPresent)
                        strOtherMailbox = string.Empty;
                    else if (mventry["External_Email_Address"].IsPresent)
                        strOtherMailbox = mventry["External_Email_Address"].Value.ToString();
                    //if (strOtherMailbox == string.Empty && csentry["COLLAB_EXTRNL_EMAIL_ADRS"].IsPresent)
                        //strOtherMailbox = csentry["COLLAB_EXTRNL_EMAIL_ADRS"].Value.ToString();
                    if (strOtherMailbox != string.Empty && !strOtherMailbox.ToLower().EndsWith("@lilly.com") && CommonLayer.IsDomainSecure(strOtherMailbox))
                        mventry["otherMailbox"].Value = strOtherMailbox;
                    else
                        mventry["otherMailbox"].Delete();
                    break;
                //NLW Mini Release
                #endregion

                #region SADMVCreate_dt
                //This attribute flow is for notification and reporting functionality ,today's date  is passed to metaverse     
                case "cd.person:PRSNL_NBR->mv.person:SADMVCreate_dt":
                    if (!(mventry["SADMVCreate_dt"].IsPresent))
                    {
                        mventry["SADMVCreate_dt"].Value = System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    }
                    break;
                #endregion

                #region SADSysIDCreate_dt
                //This attribute flow is for notification and reporting functionality , today's date  is passed to MIIS create date in Notification table
                //When we have system ID populated in MIIS , MIIS create date will be inserted in Notification table
                case "cd.person:SYSTEM_ID->mv.person:SADSysIDCreate_dt":
                    if (!(mventry["SADSysIDCreate_dt"].IsPresent))
                    {
                        mventry["SADSysIDCreate_dt"].Value = System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    }
                    break;
                #endregion

                #region sn
                //Release 2.2 changes -BOC 
				//HCM - Phase1 - Retiring KnownasMA
                case "cd.person:LST_NM,SYSTEM_ID->mv.person:sn":
                    //if (!mventry["knownAsLastName"].IsPresent && csentry["LST_NM"].IsPresent)
					if (csentry["LST_NM"].IsPresent)
                    {
                        mventry["sn"].Value = csentry["LST_NM"].Value.ToString();
                    }
                    break;
                #endregion

                #region spnsr_email_adrs
                //This attribute flow is for Contractor End date notification functionality ,lilly sponser email address  is passed in Notification table     
                case "cd.person:LY_CONTACT_PRSNL_NBR->mv.person:spnsr_email_adrs":
                    string strSpnsrAddr = GetEmailAddressFromID(mventry, csentry["LY_CONTACT_PRSNL_NBR"].Value);
                    if (strSpnsrAddr == "")
                    {
                        mventry["spnsr_email_adrs"].Delete();
                    }
                    else
                    {
                        mventry["spnsr_email_adrs"].Value = strSpnsrAddr;
                    }
                    break;
                #endregion

                #region spnsr_nm
                //This attribute flow is for Contractor End date notification functionality ,lilly sponser name  is passed in Notification table     
                case "cd.person:LY_CONTACT_PRSNL_NBR->mv.person:spnsr_nm":
                    string strSprvsrNm = GetNameFromID(mventry, csentry["LY_CONTACT_PRSNL_NBR"].Value);
                    if (strSprvsrNm == "")
                    {
                        mventry["spnsr_nm"].Delete();
                    }
                    else
                    {
                        mventry["spnsr_nm"].Value = strSprvsrNm;
                    }
                    break;
                #endregion

                #region targetAddress

                case "cd.person:BUSINESS_EMAIL_TXT,EMPLY_GRP_CD,EXT_AUTH_FLG,INTERNET_STYLE_ADRS,LAST_UPDT_DT,NOVARTIS_FLG,PRSNL_AREA_CD,PRSNL_NBR,SYSTEM_ID,UNIFIED_COMM_FLG->mv.person:TargetAddress":
                    setUserVariables(csentry["PRSNL_AREA_CD"].Value, csentry["EMPLY_GRP_CD"].Value);
                    switch (CommonLayer.GetUserType(mventry, csentry, "SAD", "no"))
                    {

                        case "MADS":
                            if (!CommonLayer.IsMigrationOverride(mventry, csentry, "SAD"))
                            {
                                mventry["TargetAddress"].Value = "SMTP:" + csentry["INTERNET_STYLE_ADRS"].Value.ToLower();
                            }
                            break;
                        case "LocalMailbox":
                            mventry["TargetAddress"].Delete();
                            break;
                        case "MailContact":
                            //BOC NLW mini Release
                            mventry["TargetAddress"].Value = "SMTP:" + csentry["BUSINESS_EMAIL_TXT"].Value.ToString().ToLower();
                            break;
                        //EOC NLW mini Release
                        case "None":
                            mventry["TargetAddress"].Delete();
                            break;
                        case "RemoteMailbox":
                            mventry["TargetAddress"].Delete();
                            break;
                    }
                    break;

                #endregion

                #region TNMSTransCode

                //0- Phone number is assigned through TNMS
                //1- Phone number is just disconnected for user (day 0)
                //2- User does not have phone number (unassigned)
                case "cd.person:LAST_UPDT_DT,PRSNL_NBR->mv.person:TNMSTransCode":
                    ConnectedMA tnmsDataMA = mventry.ConnectedMAs["TNMS Data MA"];
                    int tnmsDataconnectors = tnmsDataMA.Connectors.Count;
                    if (tnmsDataconnectors == 1)
                    {
                        mventry["TNMSTransCode"].Value = "0";
                    }
                    else
                    {
                        if (!mventry["TNMSTransCode"].IsPresent || (mventry["TNMSTransCode"].IsPresent && mventry["TNMSTransCode"].Value == "1"))
                            mventry["TNMSTransCode"].Value = "2";//new user
                        else
                            if (mventry["TNMSTransCode"].Value == "0")//If the user had TNMS Data connector (phone assigned)
                            {
                                mventry["TNMSTransCode"].Value = "1";//just disconnected user
                            }
                    }
                    break;

                #endregion

				//HCM Comments- With Quick Connect MA retired quickConnectFlgBln will flow directly to Reporting MA (direct flow rule)
                #region quickConnectFlgBln
                case "cd.person:EMPLY_GRP_CD,LAST_UPDT_DT,QUICKCNNCT_FLG,SYSTEM_ID->mv.person:quickConnectFlgBln":

                    if (csentry["QUICKCNNCT_FLG"].IsPresent)
                    {
                        if (csentry["QUICKCNNCT_FLG"].Value == "Y")
                            mventry["quickConnectFlgBln"].BooleanValue = true;
                        else if (csentry["QUICKCNNCT_FLG"].Value == "N")
                            mventry["quickConnectFlgBln"].BooleanValue = false;
                    }
                 
                    break;
                #endregion

                #region UC_Flag
                case "cd.person:EXT_AUTH_FLG,LAST_UPDT_DT,NOVARTIS_FLG,SYSTEM_ID,UNIFIED_COMM_FLG->mv.person:UC_Flag":
                    if (CommonLayer.IsMigrationOverride(mventry, csentry, "SAD") || CommonLayer.IsMigrationAcquisition(mventry, csentry, "SAD"))
                    {
                        mventry["UC_Flag"].Value = "Y";
                    }
                    else
                    {
                        if (csentry["NOVARTIS_FLG"].IsPresent && csentry["NOVARTIS_FLG"].Value.ToUpper() == "Y")
                        {
                            if (mventry["msExchMailboxGuid"].IsPresent)
                            {
                                mventry["UC_Flag"].Value = "Y";
                            }
                            else
                            {
                                if (!mventry["UC_Flag"].IsPresent)
                                    mventry["UC_Flag"].Value = csentry["UNIFIED_COMM_FLG"].Value;
                            }
                        }
                        else
                        {
                            mventry["UC_Flag"].Value = csentry["UNIFIED_COMM_FLG"].Value;
                        }
                    }

                    break;
                #endregion

                //HCM - Phase1 - Removing BUSINESS_AREA_CD from the flow for IDM SAD mapping updates
                //HCM - TBD to check the rules in Sync Flow												  
                #region "updateRecipientFlag"
                //case "cd.person:ANGLCZD_FIRST_NM,ANGLCZD_LAST_NM,ANGLCZD_MDL_NM,BUSINESS_AREA_CD,BUSINESS_EMAIL_TXT,ELANCO_FLAG,EMPLY_GRP_CD,EXT_AUTH_FLG,INTERNET_STYLE_ADRS,NOVARTIS_FLG,ORG_UNIT_CD,PRSNL_AREA_CD,PRSNL_NBR,SYSTEM_ACCESS_FLG,SYSTEM_ID,UNIFIED_COMM_FLG->mv.person:updateRecipientFlag":
                case "cd.person:ANGLCZD_FIRST_NM,ANGLCZD_LAST_NM,ANGLCZD_MDL_NM,BUSINESS_EMAIL_TXT,ELANCO_FLAG,EMPLY_GRP_CD,EXT_AUTH_FLG,INTERNET_STYLE_ADRS,NOVARTIS_FLG,ORG_UNIT_CD,PRSNL_AREA_CD,PRSNL_NBR,SYSTEM_ACCESS_FLG,SYSTEM_ID,UNIFIED_COMM_FLG->mv.person:updateRecipientFlag":
                    string systemAccessFlag = null;
                    string strUpdateRecipientFlag = null;
                    string strUPN = null;
                    string strTargetAddress = null;
                    setUserVariables(csentry["PRSNL_AREA_CD"].Value, csentry["EMPLY_GRP_CD"].Value);
                    if (csentry["SYSTEM_ACCESS_FLG"].IsPresent)
                    {
                        systemAccessFlag = csentry["SYSTEM_ACCESS_FLG"].Value;
                    }
                    if (mventry["userPrincipalName"].IsPresent)
                    {
                        strUPN = mventry["userPrincipalName"].Value;
                    }
                    if (mventry["External_Email_Address"].IsPresent)
                    {
                        strTargetAddress = mventry["External_Email_Address"].Value;
                    }
                    strUpdateRecipientFlag = System.DateTime.Now.ToString() + ",userType=" + CommonLayer.GetUserType(mventry, csentry, "SAD", "no") + ",systemAccessFlag=" + systemAccessFlag + ",UPN=" + strUPN + ",RestoredTargetAddress=" + strTargetAddress;
                        

                    mventry["updateRecipientFlag"].Value = strUpdateRecipientFlag;
                    break;

                #endregion

                #region WhereToProvision
                case "cd.person:EMPLY_GRP_CD,PRSNL_AREA_CD,SYSTEM_ID->mv.person:WhereToProvision":
                    if (strGlobalPAC == "PAC")
                    {
                        if (mventry["msExchRecipientTypeDetails"].IsPresent && mventry["msExchRecipientTypeDetails"].Value.ToString() == CommonLayer.REMOTE_MAILBOX)
                        {
                            mventry["WhereToProvision"].Value = "Remote";
                        }
                        else
                        {
                            setUserVariables(csentry["PRSNL_AREA_CD"].Value, csentry["EMPLY_GRP_CD"].Value);
                            mventry["WhereToProvision"].Value = strPACWhereToprovision;
                        }
                    }
                    else
                        mventry["WhereToProvision"].Value = "Remote";
                    break;
                #endregion

                default:
                    throw new EntryPointNotImplementedException();
            }
        }

        void IMASynchronization.MapAttributesForExport(string FlowRuleName, MVEntry mventry, CSEntry csentry)
        {
            //
            // TODO: write your export attribute flow code
            //
            throw new EntryPointNotImplementedException();
        }

        //Release 5 Lilac - The function takes personal number as an input and returns the email address
        // Data is being retrived from Metaverse
        //Only Active users to get email notification (cleanup release)
        public string GetEmailAddressFromID(MVEntry mventry, string strPrsnlID)
        {
            MVEntry[] findResultList = null;
            findResultList = Utils.FindMVEntries("employeeID", strPrsnlID);
            string strReturnValue = "";
            string strPersonStatus = string.Empty;
            if (findResultList.Length == 0) // display name never used, so take this one
            {
                strReturnValue = "";
            }
            else if (findResultList.Length == 1) // If a metaverse entry is found with the specified commonName, check if this is the entry that is being processed
            {
                MVEntry mvEntryFound = findResultList[0];
                strPersonStatus = mvEntryFound["employeeStatus"].Value;
                if (strPersonStatus == "3")
                {
                    if (mvEntryFound["mail"].IsPresent)
                    {
                        strReturnValue = mvEntryFound["mail"].Value.ToString();
                    }
                }
            }
            return strReturnValue;
        }

        public string GetNameFromID(MVEntry mventry, string strPrsnlID)
        {
            MVEntry[] findResultList = null;
            findResultList = Utils.FindMVEntries("employeeID", strPrsnlID);
            string strReturnValue = "";
            if (findResultList.Length == 0) // display name never used, so take this one
            {
                strReturnValue = "";
            }
            else if (findResultList.Length == 1) // If a metaverse entry is found with the specified commonName, check if this is the entry that is being processed
            {
                MVEntry mvEntryFound = findResultList[0];
                strReturnValue = mvEntryFound["displayName"].Value;
            }
            return strReturnValue;
        }

        //Release 5 Lilac

        public static string ConstructUniqueCommonName(CSEntry csentry, MVEntry mventry)
        {
            //
            // Unique CN is constructed by comapring the MV cn values.
            //
            string defaultCN = "";
            string commonName = "";
            string sSource, sLog, sEvent;

            if (csentry["ANGLCZD_LAST_NM"].IsPresent)
                defaultCN = csentry["ANGLCZD_LAST_NM"].Value;
            if (csentry["ANGLCZD_FIRST_NM"].IsPresent)
                defaultCN = defaultCN.Trim() + " " + csentry["ANGLCZD_FIRST_NM"].Value;
            if (csentry["ANGLCZD_MDL_NM"].IsPresent)
                defaultCN = defaultCN.Trim() + " " + csentry["ANGLCZD_MDL_NM"].Value;

            //If CN cannot be build based on ANGLICZED values
            if (defaultCN.Equals(""))
            {
                if (csentry["LST_NM"].IsPresent)
                    defaultCN = csentry["LST_NM"].Value;
                if (csentry["FRST_NM"].IsPresent)
                    defaultCN = defaultCN.Trim() + " " + csentry["FRST_NM"].Value;
                if (csentry["MDL_NM"].IsPresent)
                    defaultCN = defaultCN.Trim() + " " + csentry["MDL_NM"].Value;
            }

            int xCount = 0;
            bool uniqueCN = false;

            while (!uniqueCN)
            {
                switch (xCount)
                {
                    case 0:
                        commonName = defaultCN.ToUpper().Trim();
                        break;
                    default:
                        commonName = defaultCN.ToUpper().Trim() + " X" + xCount;
                        break;
                }

                xCount++;

                MVEntry[] findResultList = null;
                try
                {
                    //CN cannot be null, throw an excption
                    if (!commonName.Equals(""))
                    {
                        findResultList = Utils.FindMVEntries("cn", commonName);
                    }
                    else
                    {
                        throw new UnexpectedDataException();
                    }
                    if (findResultList.Length == 0) // display name never used, so take this one
                    {
                        return commonName;
                    }
                    else if (findResultList.Length == 1) // If a metaverse entry is found with the specified commonName, check if this is the entry that is being processed
                    {
                        MVEntry mvEntryFound = findResultList[0];
                        if (mvEntryFound.Equals(mventry))        // Yes this is the same entry
                            return commonName;
                    }
                }
                catch (UnexpectedDataException unex)
                {
                    //CN cannot be null, exception thrown is captured to a Event Log
                    sSource = "SAD MA";
                    sLog = "Application";
                    sEvent = csentry["PRSNL_NBR"].Value + " - CN is null";

                    if (!EventLog.SourceExists(sSource))
                        EventLog.CreateEventSource(sSource, sLog);

                    EventLog.WriteEntry(sSource, sEvent, EventLogEntryType.Error, 8101);

                    throw unex;
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }

            return defaultCN;
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
                        userpolicyArray.Add(strvalue.ToUpper());
                    }
                    valhash.Add("USERPOLICY", userpolicyArray);
                }
                if (xmlReader.ReadToFollowing("MAILBOXPOLICY"))
                {
                    xmlReader.Read();
                    string usermailboxpolicy = xmlReader.Value;
                    valhash.Add("MAILBOXPOLICY", usermailboxpolicy);
                }

                //CleanUp Release Code modified
                if (xmlReader.ReadToFollowing("OPTIONFLAGS"))
                {
                    xmlReader.Read();
                    string useroptionflags = xmlReader.Value;
                    valhash.Add("OPTIONFLAGS", useroptionflags);
                }

                //CleanUp Release Code modified
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
                strskipmail = null;
                if (listarea.Contains(paccode) && listtype.Contains(employeeType.ToUpper()))
                {
                    publicOu = afconfig["PATH"].ToString();
                    domain = afconfig["DOMAIN"].ToString().ToUpper();
                    upn = afconfig["UPN"].ToString();
                    sipHomeServer = afconfig["SIPHOMESERVER"].ToString();
                    emailtemplate = afconfig["EMAILTEMPLATE"].ToString();              //Release 5 Lilac - Setting EMAILTEMPLATE value to emailtemplate
                    archivingenabled = afconfig["ARCHIVINGENABLED"].ToString();        //Exchange OCS code
                    userlocationprofile = afconfig["USERLOCATIONPROFILE"].ToString();  //Exchange OCS code
                    userpolicy = (ArrayList)afconfig["USERPOLICY"];     //Exchange OCS code
                    strskipmail = afconfig["skipmail"].ToString(); //Exchange
                    useroptionflags = afconfig["OPTIONFLAGS"].ToString();        //CleanUp Release Code modified
                    userfederationenabled = afconfig["FEDERATIONENABLED"].ToString(); //CleanUp Release Code modified
                    strPACWhereToprovision = afconfig["WhereToProvision"].ToString();

                    strArrServerValues = getRandomizedValues((string[,])afconfig["RANDOMIZEDARRAY"]);
                    impstrArrServerValues = (ArrayList)afconfig["MSEXCHHOMESERVER"];

                    break;
                }
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

        public void LogExchangeAttributes(CSEntry csentry, MVEntry mventry)
        {
            ConnectedMA RFADMA;
            int RFconnectors = 0;
            string strmsExchUMRecipientDialPlanLink = null;
            string strmsExchUMTemplateLink = null;
            CSEntry csentryRF;

            RFADMA = mventry.ConnectedMAs["Resource Forest AD MA"];
            RFconnectors = RFADMA.Connectors.Count;

            CheckAndCreateFile();

            string strtempobjestsid, strtempobjestguid, strtempmsexchmailboxguid; //exchange
            objStreamWriter = new StreamWriter(strexchtransitionlogspath, true);

            if (mventry["msExchMailboxGuid"].IsPresent)
            {
                strtempmsexchmailboxguid = ConvertByteToStringGUID(mventry["msExchMailboxGuid"].IsPresent == false ? null : mventry["msExchMailboxGuid"].BinaryValue);
            }
            else
            {
                strtempmsexchmailboxguid = string.Empty;
            }

            string strtemp = "-------------------------------------------------------" + "\r\n";

            //RF Connector added to retrieve UM Reference variables (Exchange)
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

			//HCM Comments -Phase1 - replacing the EDS_Supervisor_Prsnl_Nbr with ManagerPrsnlNbr, since EDS MA is retiring.								   
            strtemp = strtemp
                                           + " SystemID:        " + (mventry["sAMAccountName"].IsPresent == false ? string.Empty : mventry["sAMAccountName"].Value)
                                           + "\r\n EmployeeID:      " + (mventry["employeeID"].IsPresent == false ? string.Empty : mventry["employeeID"].Value)
                                           + "\r\n DisplayName:     " + (mventry["displayName"].IsPresent == false ? string.Empty : mventry["displayName"].Value)
                                           + "\r\n EmployeeType:    " + (mventry["employeeType"].IsPresent == false ? string.Empty : mventry["employeeType"].Value)
                                           + "\r\n C:               " + (mventry["c"].IsPresent == false ? string.Empty : mventry["c"].Value)
                                           + "\r\n Personnel Area Code:  " + (mventry["personnel_area_cd"].IsPresent == false ? string.Empty : mventry["personnel_area_cd"].Value)
                                           + "\r\n Home Directory:  " + (mventry["HomeDirectory"].IsPresent == false ? string.Empty : mventry["HomeDirectory"].Value)
                                           + "\r\n DN:              " + (mventry["msDS-SourceObjectDN"].IsPresent == false ? string.Empty : mventry["msDS-SourceObjectDN"].Value)
                                           + "\r\n Supervisor Personal Number:  " + (mventry["ManagerPrsnlNbr"].IsPresent == false ? string.Empty : mventry["ManagerPrsnlNbr"].Value)
                                           + "\r\n MailNickname:     " + (mventry["mailNickname"].IsPresent == false ? string.Empty : mventry["mailNickname"].Value)   //exchange
                                           + "\r\n msExchMailboxGuid:   " + strtempmsexchmailboxguid  //exchange
                                           + "\r\n msExchHomeServerName:     " + (mventry["msExchHomeServerName"].IsPresent == false ? string.Empty : mventry["msExchHomeServerName"].Value)   //exchange
                                           + "\r\n msExchHomeMDB:      " + (mventry["HomeMDB"].IsPresent == false ? string.Empty : mventry["HomeMDB"].Value)   //exchange
                                           + "\r\n legacyExchangeDn:    " + (mventry["legacyExchangeDn"].IsPresent == false ? string.Empty : mventry["legacyExchangeDn"].Value) //exchange
                                           + "\r\n proxyAddresses:     " + (mventry["proxyAddresses"].IsPresent == false ? string.Empty : getProxyAddresses(mventry["proxyAddresses"].Values)) //exchange
                                           + "\r\n msExchUMTemplateLink:      " + (!(strmsExchUMTemplateLink == null) == false ? string.Empty : strmsExchUMTemplateLink)   //exchange
                                           + "\r\n msExchUMRecipientDialPlanLink:      " + (!(strmsExchUMRecipientDialPlanLink == null) == false ? string.Empty : strmsExchUMRecipientDialPlanLink)   //exchange
                                           + "\r\n msExchUMEnabledFlags:      " + (mventry["msExchUMEnabledFlags"].IsPresent == false ? string.Empty : mventry["msExchUMEnabledFlags"].Value)   //exchange
                                           + "\r\n";
            strtemp = strtemp + "\r\n";
            
            objStreamWriter.WriteLine(strtemp);
            objStreamWriter.Close();


        }

        public void CheckAndCreateFile()
        {
            string strFileName = BuildFileName(strexchtransitionlogs);
            //strexchtransitionlogs = strFileName;
            strexchtransitionlogspath = strFileName;
            //string strFileName = strexchtransitionlogs;
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

        public string BuildFileName(string Logslocation)
        {
            string strFileName = Logslocation + "_" + (DateTime.Now).Year.ToString() + "-" + (((DateTime.Now).Month <= 9) ? "0" + (DateTime.Now).Month.ToString() : (DateTime.Now).Month.ToString()) + "-" + ((DateTime.Now).Day.ToString().Length == 1 ? ("0" + (DateTime.Now).Day.ToString()) : (DateTime.Now).Day.ToString()) + ".txt";
            return strFileName;
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

        /// <summary>
        /// #CHG1324643 - Added this method to updated the Business Area Code value based on Elanco Flag        /// 
        /// </summary>
        /// <param name="csObject"></param>
        /// <returns>string</returns>
        //HCM - Phase1 - Removing BUSINESS_AREA_CD from the flow for IDM SAD mapping updates
        //HCM - TBD to check the rules in Sync Flow
        public string GetBusinessAreaCode(CSEntry csentry)
        {
            // Update Business area code based on Elanco flag only if Business Area Code is A or K
            //if ((csentry["ELANCO_FLAG"].IsPresent) && ((csentry["BUSINESS_AREA_CD"].Value.ToString().ToUpper() == "K") || (csentry["BUSINESS_AREA_CD"].Value.ToString().ToUpper() == "A")))
                if (csentry["ELANCO_FLAG"].IsPresent)
                {
                    if (csentry["ELANCO_FLAG"].Value.ToString() == "Y")
                    {
                        //Update Business_Area_CD in MV to K
                        return "K";
                    }
                    else if (csentry["ELANCO_FLAG"].Value.ToString() == "N")
                    {
                        //Update Business_Area_CD in MV to A
                        return "A";
                    } // If Elanco Flag value is other than Y or N,return A to MV
                    else
                    {
                         //return csentry["BUSINESS_AREA_CD"].Value.ToString();
                        return "A";
                    }
                } // If Elanco Flag value is not present,thenreturn A to MV
                else
                {
                   // return csentry["BUSINESS_AREA_CD"].Value.ToString();
                    return "A";
                }
        }

    }
}
