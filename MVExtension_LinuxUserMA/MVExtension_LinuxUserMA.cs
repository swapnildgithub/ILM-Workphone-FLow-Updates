using System;
using Microsoft.MetadirectoryServices;
using System.Collections;
using System.Xml;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Collections.Specialized;
using System.DirectoryServices;
using System.Management;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Collections.ObjectModel;
using System.DirectoryServices.ActiveDirectory;
using System.DirectoryServices.AccountManagement;


namespace Mms_Metaverse
{
    /// <summary>
    /// Summary description for MVExtensionObject.
    /// </summary>
    public class MVExtensionObject : IMVSynchronization
    {

        XmlNode rnode;
        XmlNode node;
        Hashtable afvalhash;
        string version, etypeNodeValue, publicOu, domain, upn, sipHomeServer, grpoU, scriptpath, CollabOU, subDomainOU, linuxou, linuxAutoMountOU, linuxHomeServer, strMembers, linuxGroup, linuxDomain, NAMPPath;
        string[] systemIDs;
        StringCollection memberCollection;
        String strLinuxSystemIDs = "";

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
                //get the Users to be provisioned from config                       
                node = rnode.SelectSingleNode("provision");
                etypeNodeValue = node.InnerText.ToUpper();
                //Provisioning Version number, any changes to the provisioning code
                //The Version should be changed in rules-config.xml
                node = rnode.SelectSingleNode("version");
                version = node.InnerText;
                XmlNode confignode = rnode.SelectSingleNode("siteconfigfile");
                afvalhash = LoadAFConfig(confignode.InnerText);

                node = rnode.SelectSingleNode("grpOU");
                grpoU = node.InnerText;

                node = rnode.SelectSingleNode("NLWSubGroupTypeOU");
                CollabOU = node.InnerText;

                node = rnode.SelectSingleNode("NLWSubGroupTypeSubDomainOU");
                subDomainOU = node.InnerText;

                node = rnode.SelectSingleNode("linuxOU");
                linuxou = node.InnerText;

                node = rnode.SelectSingleNode("linuxAutoMountOU");
                linuxAutoMountOU = node.InnerText;

                node = rnode.SelectSingleNode("linuxHomeServer");
                linuxHomeServer = node.InnerText;


                node = rnode.SelectSingleNode("linuxGroup");
                linuxGroup = node.InnerText;


                node = rnode.SelectSingleNode("linuxDomain");
                linuxDomain = node.InnerText;


                node = rnode.SelectSingleNode("NAMPPath");
                NAMPPath = node.InnerText;

                
                strMembers = getSystemIDsForGroup();

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
        public string getSystemIDsForGroup()
        {
            string strSystemId = string.Empty;
            using (var context = new PrincipalContext(ContextType.Domain, linuxDomain))
            {
                using (var group = GroupPrincipal.FindByIdentity(context, linuxGroup))
                {
                        var users = group.GetMembers(true);
                        foreach (UserPrincipal user in users)
                        {

                            strSystemId += user.SamAccountName + ";";
                            //user variable has the details about the user 

                        }
                
                }
            }
            //string strLDAPFilterForLinux = "memberOf:1.2.840.113556.1.4.1941:=CN=LinuxGroupMembers,OU=Universal Groups,OU=Groups,DC=rfd,DC=lilly,DC=com";

            //string strSystemId = string.Empty;
            //DirectoryEntry dEntryhighlevel = new DirectoryEntry("LDAP://vs1rfdad01.rfd.lilly.com:389",null,null,AuthenticationTypes.Secure);
            //DirectorySearcher dSearcher = new DirectorySearcher(dEntryhighlevel);

            ////filter just user objects
            //dSearcher.Filter = strLDAPFilterForLinux;
            //dSearcher.PageSize = 1000;
            //dSearcher.SearchScope = SearchScope.Subtree;
            //dSearcher.ReferralChasing = ReferralChasingOption.All;
            //SearchResultCollection resultCollection = dSearcher.FindAll();
            //int count = resultCollection.Count;
            //foreach (SearchResult userResults in resultCollection)
            //{

            //    strSystemId += userResults.Properties["samAccountName"][0].ToString() + ";";

            //}
            return strSystemId;

        }
        #region Commn Functions
       

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
                //Release 5 Lilac - Reading 2 new tags LOGINSCRIPT and EMAILTEMPLATE from config file 
                if (xmlReader.ReadToFollowing("LOGINSCRIPT"))
                {
                    xmlReader.Read();
                    string loginscript = xmlReader.Value;
                    valhash.Add("LOGINSCRIPT", loginscript);
                }
                if (xmlReader.ReadToFollowing("EMAILTEMPLATE"))
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
                    scriptpath = afconfig["LOGINSCRIPT"].ToString();  //Release 5 Lilac - Setting LOGINSCRIPT value to scriptpath
                    break;
                }

            }
        }
        //setUserVariables(mventry["personnel_area_cd"].Value, mventry["employeeType"].Value, mventry["emply_sub_grp_cd"].Value);
        public void setUserVariables(string paccode, string employeeType, string employeeSubGroup)
        {
            //
            // Initializing the variable necessary to build a AD account from the already initialized AD config file in a Hashtable
            //
            if (!string.IsNullOrEmpty(paccode) && (!string.IsNullOrEmpty(employeeType)) && (!string.IsNullOrEmpty(employeeSubGroup)))
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
                        scriptpath = afconfig["LOGINSCRIPT"].ToString();  //Release 5 Lilac - Setting LOGINSCRIPT value to scriptpath
                        break;
                    }
                    if (employeeType == "D" && (employeeSubGroup == "63" || employeeSubGroup == "64" || employeeSubGroup == "65"))
                    {
                        publicOu = "OU=Collaborators,OU=Domain Accounts,DC=amp,DC=icepoc,DC=com";
                    }

                }
        }


        private string invokePowerShell(string methodName, string systemID)
        {
            InitialSessionState initial = InitialSessionState.CreateDefault();
            RunspaceConfiguration rsp = RunspaceConfiguration.Create();
            initial.AuthorizationManager = new AuthorizationManager(rsp.ShellId);
            Runspace runspace = RunspaceFactory.CreateRunspace(initial);
            runspace.Open();
            PowerShell ps = PowerShell.Create();
            ps.Runspace = runspace;
            Command myCommand = new Command(NAMPPath);


            CommandParameter methodParam = new CommandParameter("type", methodName);
            CommandParameter systemIDParam = new CommandParameter("name", systemID);

            myCommand.Parameters.Add(methodParam);
            myCommand.Parameters.Add(systemIDParam);
            ps.Commands.AddCommand(myCommand);
            string resultPowershell = "";
            foreach (PSObject result in ps.Invoke())
            {
                resultPowershell = result.ToString();
            }
            string finalOutput = resultPowershell;

            return finalOutput;
        }

        #endregion

        void IMVSynchronization.Terminate()
        {
            //
            // TODO: Add termination logic here
            //
        }

        void IMVSynchronization.Provision(MVEntry mventry)
        {
            ConnectedMA LinuxUserMA;
            ConnectedMA sadMA, fimPortalMA;
            CSEntry csentryPerson, csentryAutomount, csEntryServiceAccount;
            ReferenceValue dn = null;
            ReferenceValue autoMountDN = null;
            string cSEntryType = mventry.ObjectType;
            int connectors;
            string rdn = "", autoMountRDN = "", sSource = "", sLog = "", sEvent = "", oU = "";
            int sadconnectors = 0, fimPortalConnectors = 0, linuxMAConnectors = 0;

            LinuxUserMA = mventry.ConnectedMAs["Linux HR MA"];
            fimPortalMA = mventry.ConnectedMAs["FIM Portal MA"];
            connectors = LinuxUserMA.Connectors.Count;
            sadMA = mventry.ConnectedMAs["Staging Area Database MA"];
            sadconnectors = sadMA.Connectors.Count;
            fimPortalConnectors = fimPortalMA.Connectors.Count;
            linuxMAConnectors = LinuxUserMA.Connectors.Count;
            if (sadconnectors == 1)
            {
                if (connectors == 0)
                {

                        if (strMembers.Contains(mventry["sAMAccountName"].Value))
                        {
                            string strNampDetails = invokePowerShell("Provision-User", mventry["samAccountName"].Value.ToLower());
                            string[] arrUserNampDetails = strNampDetails.Split(';');
                            string strUIDNumber = arrUserNampDetails[0];
                            string strAutomountInformation = arrUserNampDetails[1];

                            if (strUIDNumber == "ERROR" || strAutomountInformation == "ERROR")
                            {
                                string sSource1, sLog1, sEvent1;
                                string ExceptionMessage = mventry["sAMAccountName"].Value
                                           + " - Linux Object could not be created because NAMP script did not return UID Number or Automount homedircetory for the user.";
                                sSource1 = "SAD MA";
                                sLog1 = "Application";
                                sEvent1 = ExceptionMessage;

                                if (!EventLog.SourceExists(sSource1))
                                    EventLog.CreateEventSource(sSource1, sLog1);

                                EventLog.WriteEntry(sSource1, sEvent1, EventLogEntryType.Error, 8033);
                            }
                            else
                            {
                                if (mventry["sAMAccountName"].IsPresent)
                                {
                                    rdn = "uid=" + mventry["sAMAccountName"].Value.ToLower().Trim();
                                    autoMountRDN = "automountKey=" + mventry["samAccountName"].Value.ToLower().Trim();
                                    //HCM - Phase 1 - Modified the code to remvoe the Employee Sub Group Code attribute
								   // if (mventry["personnel_area_cd"].IsPresent && mventry["employeeType"].IsPresent && mventry["emply_sub_grp_cd"].IsPresent)
									 //   setUserVariables(mventry["personnel_area_cd"].Value, mventry["employeeType"].Value, mventry["emply_sub_grp_cd"].Value);
									if (mventry["personnel_area_cd"].IsPresent && mventry["employeeType"].IsPresent)
										setUserVariables(mventry["personnel_area_cd"].Value, mventry["employeeType"].Value);
                                    oU = linuxou;

                                }

                                if (mventry["samAccountName"].IsPresent)
                                {
                                    #region "Person"
                                    ValueCollection ocPerson;
                                    ocPerson = Utils.ValueCollection("top");
                                    ocPerson.Add("posixAccount");
                                    ocPerson.Add("shadowAccount");
                                    ocPerson.Add("account");
                                    ocPerson.Add("person");
                                    ocPerson.Add("inetorgperson");
                                    ocPerson.Add("organizationalPerson");
                                    dn = LinuxUserMA.EscapeDNComponent(rdn).Concat(oU);
                                    csentryPerson = LinuxUserMA.Connectors.StartNewConnector("person", ocPerson);
                                    #region commented code
                                    csentryPerson["loginShell"].Value = "/bin/bash";
                                    csentryPerson["uid"].Value = mventry["samAccountName"].Value.ToLower();
                                    csentryPerson["gidNumber"].IntegerValue = 7546;
                                    csentryPerson["cn"].Value = mventry["samAccountName"].Value.ToLower();
                                    csentryPerson["sn"].Value = mventry["sn"].Value.ToLower();
                                    csentryPerson["employeeNumber"].Value = mventry["employeeID"].Value;

                                    //csentryPerson["seeAlso"].Value = "CR FIM Account Add";
                                    csentryPerson["gecos"].Value = mventry["displayName"].Value.ToLower();

                                    //string strHomeDirectory = arrUserNampDetails[1].Substring(arrUserNampDetails[1].IndexOf('=') + 1);

                                    csentryPerson["homeDirectory"].Value = "/home/" + mventry["samAccountName"].Value.ToLower();

                                    #endregion


                                    csentryPerson.DN = dn;

                                    //string strUIDNumber = hashNampDetails["uidNumber"].ToString();
                                    csentryPerson["uidNumber"].IntegerValue = Convert.ToInt32(strUIDNumber);
                                    //byte[] rawpw = System.Text.UTF8Encoding.UTF8.GetBytes(mventry["initPassword"].Value);//Provisoning Password in Linux User account MA
                                    //csentryPerson["userPassword"].Values.Add(rawpw);

                                    csentryPerson.CommitNewConnector();
                                    #endregion

                                    #region "AutoMount"
                                    ValueCollection ocAutoMountKey;
                                    ocAutoMountKey = Utils.ValueCollection("automount");
                                    ocAutoMountKey.Add("top");
                                    ocAutoMountKey.Add("device");

                                    autoMountDN = LinuxUserMA.EscapeDNComponent(autoMountRDN).Concat(linuxAutoMountOU);
                                    csentryAutomount = LinuxUserMA.Connectors.StartNewConnector("automount", ocAutoMountKey);
                                    csentryAutomount["automountInformation"].Value = arrUserNampDetails[1] + "/" + mventry["samAccountname"].Value.ToLower();
                                    csentryAutomount["automountKey"].Value = mventry["samAccountname"].Value.ToLower();
                                    csentryAutomount["cn"].Value = mventry["samAccountname"].Value.ToLower();
                                    csentryAutomount.DN = autoMountDN;
                                    csentryAutomount.CommitNewConnector();
                                    #endregion

                                }
                            }
                        }
                  

                }

            }
            else
            {
                if (fimPortalConnectors == 1)
                {//If the service account is present in Portal
                    if (linuxMAConnectors == 0)
                    {//If the service account is present in Linux
                        if (mventry["uid"].IsPresent)
                        {
                            if (mventry["uidNumber"].IsPresent)
                            {
                                //Req for version 1.1: The system shall use the same value for Display Name,dn,sn,uid, cn
                                rdn = "uid=" + mventry["uid"].Value.ToString();
                                oU = linuxou;
                                #region "Service Account"
                                //adding objectClasses
                                ValueCollection ocPerson;
                                ocPerson = Utils.ValueCollection("top");
                                ocPerson.Add("posixAccount");
                                ocPerson.Add("shadowAccount");
                                ocPerson.Add("account");
                                ocPerson.Add("person");
                                ocPerson.Add("inetorgperson");
                                ocPerson.Add("organizationalPerson");
                                //creating DN
                                dn = LinuxUserMA.EscapeDNComponent(rdn).Concat(oU);
                                csentryPerson = LinuxUserMA.Connectors.StartNewConnector("account", ocPerson);
                                if (mventry["loginShell"].IsPresent)
                                    csentryPerson["loginShell"].Value = mventry["loginShell"].Value;
                                //Req for version 1.1: The system shall use the same value for Display Name,dn,sn,uid, cn

                                csentryPerson["uid"].Value = mventry["uid"].ToString();
                                if (mventry["gidNumber"].IsPresent)
                                    csentryPerson["gidNumber"].IntegerValue = mventry["gidNumber"].IntegerValue;
                                else
                                    csentryPerson["gidNumber"].IntegerValue = 7546;
                                //Req for version 1.1: The system shall use the same value for Display Name,dn,sn,uid, cn

                                csentryPerson["cn"].Value = mventry["uid"].Value.ToString();

                                //Req for version 1.1: The system shall use the same value for Display Name,dn,sn,uid, cn
                                csentryPerson.DN = dn;
                                if (mventry["uidNumber"].IsPresent)
                                    csentryPerson["uidNumber"].IntegerValue = Convert.ToInt32(mventry["uidNumber"].Value);
                                if (mventry["unixPassword"].IsPresent)
                                {
                                    byte[] rawpw = System.Text.UTF8Encoding.UTF8.GetBytes(mventry["unixPassword"].Value);//Provisoning Password in Linux User account MA
                                    csentryPerson["userPassword"].Values.Add(rawpw);
                                }
                                if (mventry["gecos"].IsPresent)
                                    csentryPerson["gecos"].Value = mventry["gecos"].Value.ToString();
                                //Req for version 1.1: The system shall use the same value for Display Name,dn,sn,uid, cn

                                csentryPerson["sn"].Value = mventry["uid"].Value.ToString();
                                if (mventry["HomeDirectory"].IsPresent)
                                    csentryPerson["homeDirectory"].Value = mventry["HomeDirectory"].Value;
                                //commit connector
                                csentryPerson.CommitNewConnector();
                                #endregion
                            }
                            else 
                            {
                                string sSource1, sLog1, sEvent1;
                                string ExceptionMessage = mventry["sAMAccountName"].Value
                                           + " - Linux Object could not be created because NAMP script did not return UID Number for the service account.";
                                sSource1 = "SAD MA";
                                sLog1 = "Application";
                                sEvent1 = ExceptionMessage;

                                if (!EventLog.SourceExists(sSource1))
                                    EventLog.CreateEventSource(sSource1, sLog1);

                                EventLog.WriteEntry(sSource1, sEvent1, EventLogEntryType.Error, 8033);                            
                            }
                        }
                    }
                }
                else if (fimPortalConnectors == 0)
                { // Else clause is used for de-provisioing of accounts
                    if (linuxMAConnectors == 0)
                    {
                        //Do nothing if Linux account was never provisioned
                    }
                    else
                    {
                        csEntryServiceAccount = LinuxUserMA.Connectors.ByIndex[0];
                        csEntryServiceAccount.Deprovision();
                    }
                }
            }

        }

        bool IMVSynchronization.ShouldDeleteFromMV(CSEntry csentry, MVEntry mventry)
        {
            //
            // TODO: Add MV deletion logic here
            //
            throw new EntryPointNotImplementedException();
        }
    }
}
