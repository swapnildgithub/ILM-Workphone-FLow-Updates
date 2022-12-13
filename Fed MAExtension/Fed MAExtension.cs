
using System;
using System.Xml;
using Microsoft.MetadirectoryServices;
using System.Collections;
using CommonLayer_NameSpace;

namespace Mms_ManagementAgent_Fed_MAExtension
{
    /// <summary>
    /// Summary description for MAExtensionObject.
    /// </summary>
    public class MAExtensionObject : IMASynchronization
    {
        string timetolive;
        string version, lcsflagoff, etypeNodeValue, publicOu, domain, upn, sipHomeServer, strexchangemailprefix, strlegacyExchangeDN, strAcceptedDomains, strAcceptedDomainsCN, strRBACPolicyLink, strbogusdomain, strOVAFlagSwitch;
        string archivingenabled, userlocationprofile, userinternetaccessenabled, userfederationenabled, useroptionflags, strUMMailboxPolicy, strO365List;
        ArrayList arrproxyadd, userpolicy, impstrArrServerValues;
        // string strLillyDomainSuffix, strElancoDomainSuffix, strAgspanDomainSuffix, strContractorEmailDomainPrefix, strRFDomain;
        string strLillyDomainSuffix, strElancoDomainSuffix, strAgspanDomainSuffix, strContractorEmailDomainPrefix, strRFDomain, strAllowableUPNDomains;//CHG-CHG1178839, added string "strAllowableUPNDomains".
        string[] strArrServerValues; //Exchange Random
        string strskipmail;
        Hashtable afvalhash;
        public MAExtensionObject()
        {
            //
            // TODO: Add constructor logic here
            //
        }
        void IMASynchronization.Initialize()
        {
            XmlDocument config;
            XmlNode rnode, node, rnode1;
            string dir;
            try
            {
                const string XML_CONFIG_FILE = @"\rules-config.xml";
                config = new XmlDocument();
                dir = Utils.ExtensionsDirectory;
                config.Load(dir + XML_CONFIG_FILE);

                rnode = config.SelectSingleNode("rules-extension-properties");
                node = rnode.SelectSingleNode("environment");
                string env = node.InnerText;
                rnode = config.SelectSingleNode
                    ("rules-extension-properties/management-agents/" + env + "/ad-ma");
                XmlNode confignode = rnode.SelectSingleNode("siteconfigfile");
                afvalhash = LoadAFConfig(confignode.InnerText);

				//INIT PWD Depro Error Fix
				//Get the timetolive(ttyl) value from config file
                node = rnode.SelectSingleNode("ttl");
                timetolive = node.InnerText;
				
                //ERMA change

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
                //ERMA code EOC

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
            int sadconnectors = 0, fedconnectors = 0;
            ConnectedMA sadMA = mventry.ConnectedMAs["Staging Area Database MA"];
            sadconnectors = sadMA.Connectors.Count;
            ConnectedMA FedMA = mventry.ConnectedMAs["Fed MA"];
            fedconnectors = FedMA.Connectors.Count;


            switch (FlowRuleName)
            {

                case "cd.user:altSecurityIdentities<-mv.PrivilegedAccount:sAMAccountName":
                    if (mventry["sAMAccountName"].IsPresent && mventry["sAMAccountName"].Value.ToLower().EndsWith("-ca"))
                        csentry["altSecurityIdentities"].Value = "CloudAdministrators";
                    if (mventry["sAMAccountName"].IsPresent && mventry["sAMAccountName"].Value.ToLower().EndsWith("-ds"))
                        csentry["altSecurityIdentities"].Value = "DSAccounts";
                    break;


                #region userPrincipalName


                case "cd.user:userPrincipalName<-mv.person:Anglczd_first_nm,Anglczd_last_nm,Anglczd_middle_nm,Business_Area_CD,employeeID,employeeType,Exch_Trans_CD,mail,msDS-SourceObjectDN,personnel_area_cd,sAMAccountName":
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
                        } //End - CHG-CHG1178839
                        else
                            //If email is not present set UPN with System ID
                            if (!string.IsNullOrEmpty(tempupn))
                            csentry["userPrincipalName"].Value = mventry["sAMAccountName"].Value + "@" + tempupn;
                            else // Updated for indentation
                            csentry["userPrincipalName"].Value = mventry["sAMAccountName"].Value + "@" + strRFDomain;
                    }
                    break;
                #endregion
                case "cd.user:msDS-UserAccountDisabled<-mv.person:admin_override_flag,deprovisionedDate,employeeID,employeeType,emply_sub_grp_cd,System_Access_Flag":
                    //Release 5 - Lilac. Process All employee types mentioned in Rules.config
                    //The account if disabled manually for Admin purpose then it should not be reenabled by MIIS. 
                    //So while disabling the account manually ,perticular Admin Override Text should be present in desciption field e.g. "ADMIN" this Text configurable from rules.config file.
                    //If description is present and it is "ADMIN" (whatever text mentioned in rules.config) then make the admin_override_flag as True. If True then skip processing.
                    if (mventry["admin_override_flag"].Value == "False")
                    {
                        if (mventry["deprovisionedDate"].IsPresent)
                        {
                            if (!AccountTTLExpired(mventry["deprovisionedDate"].Value, timetolive))
                            {
                                csentry["msDS-UserAccountDisabled"].BooleanValue = true;
                            }
                        }
                        else
                        {
                            csentry["msDS-UserAccountDisabled"].BooleanValue = false;
                        }

                    }
                    break;

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
    }
}
