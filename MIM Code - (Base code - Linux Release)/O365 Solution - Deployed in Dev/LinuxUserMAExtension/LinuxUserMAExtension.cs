using System.IO;
using System.Xml;
using System.Text;
using System;
using System.Collections;
using System.Collections.Specialized;
using Microsoft.MetadirectoryServices;
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

namespace Mms_ManagementAgent_LinuxUserMAExtension
{
    /// <summary>
    /// Summary description for MAExtensionObject.
    /// </summary>
    public class MAExtensionObject : IMASynchronization
    {
        string strMembers;
        XmlNode rnode;
        XmlNode node;
        Hashtable afvalhash;
        string version, etypeNodeValue, publicOu, domain, upn, sipHomeServer, linuxGrp, scriptpath, CollabOU, subDomainOU, linuxou, linuxAutoMountOU, linuxHomeServer, strLockSwitch, strProjectionSwitch;
        string[] systemIDs;
        
        String strLinuxSystemIDs = "";
        public MAExtensionObject()
        {
            //
            // TODO: Add constructor logic here
            //
        }
        void IMASynchronization.Initialize()
        {
            strMembers = getSystemIDsForGroup("CN=LinuxGroupMembers,OU=Universal Groups,OU=Groups,DC=rfd,DC=lilly,DC=com");
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


            node = rnode.SelectSingleNode("linuxGroup");
            string strLinuxGroup = node.InnerText;
            node = rnode.SelectSingleNode("lockSwitch");
            string strLockSwitch = node.InnerText;

            node = rnode.SelectSingleNode("projectionSwitch");
            strProjectionSwitch = node.InnerText;



        }
        public string getSystemIDsForGroup(string strFilter)
        {
            string strLDAPFilterForLinux = "memberOf:1.2.840.113556.1.4.1941:=CN=LinuxGroupMembers,OU=Universal Groups,OU=Groups,DC=rfd,DC=lilly,DC=com";

            string strSystemId = string.Empty;
            DirectoryEntry dEntryhighlevel = new DirectoryEntry("LDAP://vs1rfdad01.rfd.lilly.com:389", @"RFD\MIISWRITER", "GBP@ck3r$", AuthenticationTypes.Secure);
            DirectorySearcher dSearcher = new DirectorySearcher(dEntryhighlevel);

            //filter just user objects
            dSearcher.Filter = strLDAPFilterForLinux;
            dSearcher.PageSize = 1000;
            dSearcher.SearchScope = SearchScope.Subtree;
            SearchResultCollection resultCollection = dSearcher.FindAll();
            int count = resultCollection.Count;
            foreach (SearchResult userResults in resultCollection)
            {

                strSystemId += userResults.Properties["samAccountName"][0].ToString() + ";";

            }
            return strSystemId;


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
        void IMASynchronization.Terminate()
        {
            //
            // TODO: write termination code
            //
        }

        bool IMASynchronization.ShouldProjectToMV(CSEntry csentry, out string MVObjectType)
        {
            MVObjectType = "ServiceAccounts";
            bool blnIsServiceAccount = false;
            if (strProjectionSwitch == "ON")
            {
                if (csentry["description"].IsPresent)
                {
                    foreach (Object strDescription in csentry["description"].Values)
                    {
                        if (strDescription.ToString().Contains("serviceaccount"))
                        {
                            blnIsServiceAccount = true;
                            break;
                        }
                    }

                }
            }
            return blnIsServiceAccount;
            //throw new EntryPointNotImplementedException();
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
            //String strUIDNumber = "";
            ////if (FlowRuleName == "cd.account#2:description,uidNumber->uidNumber")
            ////{
            ////    if (csentry["description"].IsPresent && csentry["description"].Value.Contains("ServiceAccount"))
            ////    {
            ////        strUIDNumber = csentry["uidNumber"].Value;
            ////        values.Add(strUIDNumber);
            ////    }
            ////}
            //if (FlowRuleName == "cd.account#3:uid->sAMAccountName")
            //{
            //    strUIDNumber = csentry["uid"].Value;
            //    values.Add(strUIDNumber);
            //}
            throw new EntryPointNotImplementedException();
        }

        bool IMASynchronization.ResolveJoinSearch(string joinCriteriaName, CSEntry csentry, MVEntry[] rgmventry, out int imventry, ref string MVObjectType)
        {
            throw new EntryPointNotImplementedException();
        }

        void IMASynchronization.MapAttributesForImport(string FlowRuleName, CSEntry csentry, MVEntry mventry)
        {
            switch (FlowRuleName)
            {

//TBD - LINUX Improvement Release
                case "cd.account:<dn>,description->mv.person:unixDescString":
                     if (csentry["description"].IsPresent)
                    {
                        string strDesc = "";
                        foreach (Object strDescription in csentry["description"].Values)
                        {
                            strDesc = strDesc + strDescription.ToString() + ";";
                        }
                        if (strDesc.Length > 0)
                            strDesc = strDesc.Substring(0, strDesc.Length - 1);
                        mventry["unixDescString"].Value = strDesc;
                    }
                    break;
                case "cd.account:loginShell->mv.ServiceAccounts:loginShell":
                    string strLoginShell = "";
                    if (csentry["loginShell"].IsPresent && csentry["loginShell"].Value == "/bin/sh")
                        strLoginShell = "/bin/sh";
                    else if (csentry["loginShell"].IsPresent && csentry["loginShell"].Value == "/bin/bash")
                        strLoginShell = "/bin/bash";
                    else if (csentry["loginShell"].IsPresent && csentry["loginShell"].Value == "/bin/csh")
                        strLoginShell = "/bin/csh";
                    else if (csentry["loginShell"].IsPresent && csentry["loginShell"].Value == "/bin/ksh")
                        strLoginShell = "/bin/ksh";
                    mventry["loginShell"].Value = strLoginShell;
                    break;
//TBD - LINUX Improvement Release
                case "cd.account:description->mv.ServiceAccounts:unixDescString":
                    if (csentry["description"].IsPresent)
                    {
                        string strDesc = "";
                        foreach (Object strDescription in csentry["description"].Values)
                        {
                            strDesc = strDesc + strDescription.ToString() + ";";
                        }
                        if (strDesc.Length > 0)
                            strDesc = strDesc.Substring(0, strDesc.Length - 1);
                        mventry["unixDescString"].Value = strDesc;
                    }
                    break;
                case "cd.account:<dn>,objectClasses->mv.ServiceAccounts:objectClasses":

                    foreach (Object strObjectClass in csentry.ObjectClass)
                    {
                        mventry["objectClasses"].Values.Add(strObjectClass.ToString());
                    }
                    break;

                case "cd.account:cn,description->mv.ServiceAccounts:ServiceAccountOwners":

                     MVEntry[] findResultListServiceAccountOwner = null;
                  
                    if (csentry["description"].IsPresent)
                    {
                        mventry["ServiceAccountOwners"].Values.Clear();
                        string strServiceAccountOwner = string.Empty;
                        foreach (Value valueMemberItem in csentry["description"].Values)
                        {
                            if (valueMemberItem.ToString().ToLower().StartsWith("owner:"))
                            {
                                strServiceAccountOwner = valueMemberItem.ToString().Split(':')[1];
                                break;
                            }
                        }
                        string[] arrOwnerSystemID = strServiceAccountOwner.Split(',');
                        for (int i = 0; i < arrOwnerSystemID.Length; i++)
                        {
                            findResultListServiceAccountOwner = null;
                            findResultListServiceAccountOwner = Utils.FindMVEntries("samAccountName", arrOwnerSystemID[i].ToString().TrimStart().TrimEnd());
                            if (findResultListServiceAccountOwner.Length > 0)
                            {
                                mventry["ServiceAccountOwners"].Values.Add(findResultListServiceAccountOwner[0]["csObjectID"].Value.ToString());
                                findResultListServiceAccountOwner = null;
                            }
                        }
                    }
                    break;
            }
        }


        void IMASynchronization.MapAttributesForExport(string FlowRuleName, MVEntry mventry, CSEntry csentry)
        {
            switch (FlowRuleName)
            {
				//TBD - LINUX Improvement Release
                case "cd.account:description<-mv.ServiceAccounts:unixDescription":
                    if (mventry["unixDescription"].IsPresent)
                    {
                        //string strDesc = "";
                        foreach (Object strDescription in mventry["unixDescription"].Values)
                        {
                            csentry["description"].Values.Add(strDescription.ToString());
                        }
                        //if (strDesc.Length > 0)
                        //    strDesc = strDesc.Substring(0, strDesc.Length - 1);
                        //csentry["description"].Value = strDesc;
                    }
                    break;
                case "cd.account:uid<-mv.ServiceAccounts:uid":
                    if (mventry["uid"].IsPresent)
                        csentry["uid"].Value = mventry["uid"].Value.ToLower();
                    break;
                case "cd.account:gecos<-mv.ServiceAccounts:gecos":
                    if (mventry["gecos"].IsPresent)
                        csentry["gecos"].Value = mventry["gecos"].Value.ToLower();//TBD - LINUX Improvement Release
                    break;
                case "cd.account:userPassword<-mv.ServiceAccounts:unixPassword":
                    byte[] rawpw = System.Text.UTF8Encoding.UTF8.GetBytes(mventry["unixPassword"].Value);//Provisoning Password in Linux User account MA
                    csentry["userPassword"].Values.Add(rawpw);

                    break;
                case "cd.automount:automountInformation<-mv.person:sAMAccountName":
                    if (csentry["automountInformation"].IsPresent)
                    {
                        string strExistingAutoMount = csentry["automountInformation"].Value;
                        strExistingAutoMount = strExistingAutoMount.Substring(0, strExistingAutoMount.LastIndexOf("/") + 1);
                        strExistingAutoMount = strExistingAutoMount + mventry["sAMAccountName"].Value.ToLower();
                        csentry["automountInformation"].Delete();
                        csentry["automountInformation"].Value = strExistingAutoMount;
                    }
                    break;
                case "cd.automount:automountKey<-mv.person:sAMAccountName":
                    csentry["automountKey"].Value = mventry["sAMAccountName"].Value.ToLower();
                    break;
                case "cd.account:cn<-mv.person:last_updt_date,sAMAccountName":
                    if (!csentry["cn"].IsPresent)
                    csentry["cn"].Value = mventry["sAMAccountName"].Value.ToLower();
                    break;
                case "cd.account:sn<-mv.person:sn":
                     csentry["sn"].Value = mventry["sn"].Value.ToLower();
                    break;
                //case "cd.account:cn<-mv.person:sAMAccountName":

                //    csentry["cn"].Value = mventry["sAMAccountName"].Value.ToLower();
                //    break;
                case "cd.account:gecos<-mv.person:displayName,sAMAccountName":
                    if (!csentry["gecos"].IsPresent)
                        csentry["gecos"].Value = mventry["displayName"].Value.ToLower();
                    break;
                case "cd.account:uid<-mv.person:last_updt_date,sAMAccountName":
                    if (!csentry["uid"].IsPresent)
                        //csentry["uid"].Delete();
                        csentry["uid"].Value = mventry["sAMAccountName"].Value.ToLower();
                    break;

                case "cd.account:homeDirectory<-mv.person:sAMAccountName":
                    if (!csentry["homeDirectory"].IsPresent)
                    {
                        if (mventry["samAccountName"].IsPresent)
                        {

                            csentry["homeDirectory"].Value = @"/home/" + mventry["samAccountName"].Value.ToLower();
                        }
                    }
                    break;
                case "cd.account:nsAccountLock<-mv.person:last_updt_date,sAMAccountName":
                    
                    if (mventry["samAccountName"].IsPresent && strMembers.Contains(mventry["samAccountName"].Value.ToUpper()))
                    {
                        csentry["nsAccountLock"].Delete();
                    }
                    else
                    {
                            csentry["nsAccountLock"].Value = "true";

                    }

                    break;

                case "cd.account:homeDirectory<-mv.ServiceAccounts:HomeDirectory":
                    if (mventry["HomeDirectory"].IsPresent)
                        csentry["homeDirectory"].Value = mventry["HomeDirectory"].Value.ToLower();
                    break;
//TBD - LINUX Improvement Release
                case "cd.account:description<-mv.ServiceAccounts:cn,ServiceAccountExportOwners,serviceAccountName,unixDescString":

                    string strSAportalDescription = string.Empty;
                        string strSAPortalDesc = string.Empty;
                        bool IsSADescription = false;
                        string strSAOwner = string.Empty;
                        bool IsSAOwner = false;
                        foreach (Value desc in csentry["description"].Values)
                        {
                            if (desc.ToString().StartsWith("portalDescription:"))
                            {
                                strSAPortalDesc = desc.ToString();
                                IsSADescription = true;
                            }
                            if (desc.ToString().StartsWith("owner:"))
                            {
                                strSAOwner = desc.ToString();
                                IsSAOwner = true;
                            }
                        }

                        MVEntry[] findResultListSAOwner = null;

                        string strOwners = string.Empty;
                        if (mventry["ServiceAccountExportOwners"].IsPresent)
                        {
                            foreach (Value valueMemberItem in mventry["ServiceAccountExportOwners"].Values)
                            {
                                findResultListSAOwner = Utils.FindMVEntries("csObjectID", valueMemberItem.ToString());
                                if (findResultListSAOwner.Length == 1)
                                {
                                    strOwners = strOwners + findResultListSAOwner[0]["samAccountName"].Value.ToLower() + ",";
                                    findResultListSAOwner = null;
                                }
                            }
                            if (strOwners.Length > 0)
                            {
                                strOwners = "owner: " + strOwners.Substring(0, strOwners.Length - 1);
                            }
                            if (IsSAOwner == false)
                            {
                                csentry["description"].Values.Add(strOwners);
                            }
                            else
                            {
                                csentry["description"].Values.Remove(strSAOwner);
                                csentry["description"].Values.Add(strOwners);
                            }
                        }

                        //
                        //Netgroup description
                        if (mventry["unixDescString"].IsPresent)//TBD - LINUX Improvement Release
                        {
                            ValueCollection vc = csentry["description"].Values;
                            ValueCollection vcdescription = Utils.ValueCollection("initialValue");
                            vcdescription.Remove("initialValue");
                            vcdescription.Add(csentry["description"].Values);
                            //string[] arrmvDesc = strmvDesc.Split(';');
                            strSAportalDescription = "portalDescription:" + mventry["unixDescString"].Value.ToString();

                            ////

                            if (IsSADescription == false)
                            {
                                csentry["description"].Values.Add(strSAportalDescription);
                            }
                            else
                            {
                                csentry["description"].Values.Remove(strSAPortalDesc);
                                csentry["description"].Values.Add(strSAportalDescription);
                            }
                        }
                    break;
            }
        }
    }
}
