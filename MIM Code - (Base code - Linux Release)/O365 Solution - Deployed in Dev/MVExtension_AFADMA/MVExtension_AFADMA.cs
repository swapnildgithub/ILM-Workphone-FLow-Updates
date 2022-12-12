using System;
using Microsoft.MetadirectoryServices;
using System.Collections;
using System.Xml;
using System.Text;
using System.Collections.Specialized;
using System.Diagnostics;
using CommonLayer_NameSpace;

namespace Mms_Metaverse
{
    /// <summary>
    /// Extension for MetaVerse. Accounts provisoned from here.
    /// </summary>
    public class MVExtensionObject : IMVSynchronization
    {
        XmlNode rnode;
        XmlNode node;
        Hashtable afvalhash;
        string version, etypeNodeValue, publicOu, domain, upn, sipHomeServer, grpoU, scriptpath, CollabOU, subDomainOU, serviceAccountOU;

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

              //  node = rnode.SelectSingleNode("NLWSubGroupTypeOU");
               // CollabOU = node.InnerText;

               // node = rnode.SelectSingleNode("serviceAccountOU");
               // serviceAccountOU = node.InnerText;

               // node = rnode.SelectSingleNode("NLWSubGroupTypeSubDomainOU");
              //  subDomainOU = node.InnerText;

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
            /// Provisioning accounts in AF.
            ///

            ConnectedMA afadMA, gmadma;
            ConnectedMA sadMA;
            CSEntry csentry;
            ReferenceValue dn = null;
            string cSEntryType = mventry.ObjectType;
            int connectors;
            string rdn, sSource, sLog, sEvent, oU, formattedDate;
            int sadconnectors = 0, gmaconnectors = 0;

            try
            {
                switch (cSEntryType)
                {                    

                    case "person":
                        {
                            //Ensure CN is present or write to Event Log
                            if (!mventry["cn"].IsPresent)
                            {
                                string ExceptionMessage = mventry["employeeID"].Value
                                    + " - The attribute CN was not present on the MV";
                                sSource = "MIIS AF-ADMA Provisioning";
                                sLog = "Application";
                                sEvent = ExceptionMessage;

                                if (!EventLog.SourceExists(sSource))
                                    EventLog.CreateEventSource(sSource, sLog);

                                EventLog.WriteEntry(sSource, sEvent, EventLogEntryType.Error, 8021);

                                throw new UnexpectedDataException(ExceptionMessage);
                            }

                            else if (!mventry["samAccountname"].IsPresent)
                            {
                                //AF account would be created only if samAccountname is present
                                //else do nothing
                            }
                            else if (etypeNodeValue.Contains(mventry["employeeType"].Value.ToUpper()))
                            {
                                // Ensure that the cn attribute is present.
                                //Provisoning accounts in AF
                                afadMA = mventry.ConnectedMAs["Account Forest AD MA"];
                                //building RDN                               
                                //rdn = "CN=" + mventry["cn"].Value.Trim();
                                //Build RDN with System ID
                                rdn = "CN=" + mventry["sAMAccountName"].Value.Trim();

                                //Get the OU, domain and UPN from AD XML config 
                                //HCM - Phase 1 - Modified the code to remvoe the Employee Sub Group Code attribute
                                setUserVariables(mventry["personnel_area_cd"].Value, mventry["employeeType"].Value);

                                oU = publicOu;

                                //Write to the Event log in case these necessary data isn't available
                                if (publicOu == null | domain == null | upn == null | sipHomeServer == null)
                                {
                                    string ExceptionMessage = mventry["employeeID"].Value
                                        + " - PACCODE and EMPLOYEETYPE combination cannot retrieve data from XML file";
                                    sSource = "MIIS AF-ADMA Provisioning";
                                    sLog = "Application";
                                    sEvent = ExceptionMessage;

                                    if (!EventLog.SourceExists(sSource))
                                        EventLog.CreateEventSource(sSource, sLog);

                                    EventLog.WriteEntry(sSource, sEvent, EventLogEntryType.Error, 8031);
                                    // If a tag does not exist in the xml, then the stopped-extension-dll 
                                    // error will be thrown.
                                    break;
                                }
                                else if (publicOu == string.Empty | domain == string.Empty | upn == string.Empty | sipHomeServer == string.Empty)
                                {
                                    string ExceptionMessage = mventry["employeeID"].Value
                                            + " - AF connector can not be created for this record because it don't have value for required attributes in Environment.xml file.";
                                    sSource = "MIIS AF-ADMA Provisioning";
                                    sLog = "Application";
                                    sEvent = ExceptionMessage;

                                    if (!EventLog.SourceExists(sSource))
                                        EventLog.CreateEventSource(sSource, sLog);

                                    EventLog.WriteEntry(sSource, sEvent, EventLogEntryType.Error, 8031);
                                    // If a tag does not exist in the xml, then the stopped-extension-dll 
                                    // error will be thrown.
                                    break;
                                }

                                //Building a DN using ou and rdn
                                dn = afadMA.EscapeDNComponent(rdn).Concat(oU);

                                connectors = afadMA.Connectors.Count;
                                sadMA = mventry.ConnectedMAs["Staging Area Database MA"];
                                sadconnectors = sadMA.Connectors.Count;
                                //NLW mini Release Change
                                if (sadconnectors == 1)
                                {
                                    if (connectors == 0)
                                    {
                                        if (CommonLayer.ISSysAccessFlagSet(mventry))
                                        {

                                            csentry = afadMA.Connectors.StartNewConnector("user");
                                            //Set the User ID
                                            csentry["sAMAccountName"].Value = mventry["sAMAccountName"].Value;
                                            //Set the initial password
                                            csentry["unicodePwd"].Values.Add(mventry["initPassword"].Value);

                                            //Set the account to Enabled status
                                            csentry["userAccountControl"].IntegerValue = 512;
                                            csentry["pwdLastSet"].IntegerValue = 0;
                                            formattedDate = DateTime.Now.ToString("dd-MMM-yyyy").ToUpper();
                                            //Release 5 Lilac - getting scriptPath value from config and in AD 
                                            if (scriptpath.ToString().Trim().Length > 0)
                                            {
                                                csentry["scriptPath"].Value = scriptpath.ToString();
                                            }

                                            //Log the changes
                                            csentry["info"].Delete();
                                            csentry["info"].Value = "[CREATED~MIISWRITER~" + formattedDate + "~" + version + "]";
                                            csentry.DN = dn;
                                            csentry.CommitNewConnector();


                                        }

                                    }
                                    else if (connectors == 1)
                                    {
                                        // Check if the connector has a different DN and rename if necessary.
                                        // Get the connector.
                                        csentry = afadMA.Connectors.ByIndex[0];

                                        //Change the CN value to system ID for Office 365
                                        //dn = afadMA.EscapeDNComponent("CN=" + mventry["sAMAccountName"].Value).Concat(oU);
                                        //csentry.DN = dn;

                                        //Connectors are same, but the rename might be within the same Domain
                                        //Split the DN to get the domain
                                        string tmpsplt = csentry.DN.ToString().ToUpper().Split(new string[] { "DC=" }, 2, StringSplitOptions.None)[1];
                                        string tmpdmn = tmpsplt.Split(',')[0];
                                        if (tmpdmn.Contains(mventry["domain"].Value.ToUpper()))
                                        {
                                            //Microsoft Identity Integration Server will rename/move if different, if not, nothing will happen.
                                            dn = afadMA.EscapeDNComponent("CN=" + mventry["sAMAccountName"].Value).Concat(oU);
                                            csentry.DN = dn;
                                        }
                                        else
                                        {
                                            oU = "OU=" + csentry.DN.ToString().ToUpper().Split(new string[] { "OU=" }, 2, StringSplitOptions.None)[1];
                                            dn = afadMA.EscapeDNComponent("CN=" + mventry["sAMAccountName"].Value).Concat(oU);
                                            csentry.DN = dn;
                                        }
                                        //NLW Mini Release
                                        if (!CommonLayer.ISSysAccessFlagSet(mventry))
                                        {
                                            csentry.Deprovision();
                                        }
                                    }
                                    else
                                    {
                                        // More than 2 connectors
                                        // Happens only with user deleted from MV but still in AD and reprovisioned in MV again (MIIS capabilites)
                                        // Join with the already avaiable AD account and remove the AD account created by MIIS
                                        if (connectors == 2)
                                        {
                                            int i = 0;
                                            bool multiconnectors = true;
                                            while (i < 2 && multiconnectors)
                                            {
                                                csentry = afadMA.Connectors.ByIndex[i];
                                                i++;
                                                //MIIS Created AD CS will be deleted so already avaialbe AD is the only CS for the user
                                                if (csentry["sAMAccountName"].Value.Equals(mventry["sAMAccountName"].Value)
                                                    && !csentry["whenCreated"].IsPresent
                                                    )
                                                {
                                                    csentry.Deprovision();
                                                    multiconnectors = false;
                                                }
                                            }
                                            //Throw if more than 2 connectors or unable recover from above scenario
                                            if (multiconnectors)
                                            {
                                                throw (new UnexpectedDataException("Multiple connectors in AF:" + afadMA.Connectors.Count.ToString()));
                                            }
                                        }
                                        else
                                        {
                                            //Throw if more than 2 connectors or unable recover from above scenario
                                            throw (new UnexpectedDataException("Multiple connectors in AF:" + afadMA.Connectors.Count.ToString()));
                                        }
                                    }
                                }
                            }
                            break;
                        }

                    case "group":
                        {
                            if (!mventry["sAMAccountName"].IsPresent)
                            {
                                //Log an error                           
                            }
                            else
                            {
                                //Provisoning group accounts in AF
                                afadMA = mventry.ConnectedMAs["Account Forest AD MA"];
                                gmadma = mventry.ConnectedMAs["Group Management MA"];
                                //building RDN                                
                                rdn = "CN=" + mventry["sAMAccountName"].Value.Trim();
                                //Building a DN using ou and rdn
                                dn = afadMA.EscapeDNComponent(rdn).Concat(grpoU);
                                connectors = afadMA.Connectors.Count;
                                gmadma = mventry.ConnectedMAs["Group Management MA"];
                                gmaconnectors = 1; // gmadma.Connectors.Count;

                                if (gmaconnectors == 1)
                                {
                                    if (connectors == 0)
                                    {
                                        csentry = afadMA.Connectors.StartNewConnector("group");
                                        csentry.DN = dn;
                                        csentry.CommitNewConnector();
                                    }
                                    else if (connectors == 1)
                                    {

                                        // Check if the connector has a different DN and rename if necessary.
                                        // Get the connector.
                                        csentry = afadMA.Connectors.ByIndex[0];
                                        csentry.DN = dn;
                                    }

                                }

                                else if (gmaconnectors == 0)
                                {
                                    if (connectors == 0)
                                    {
                                        //Do nothing if the group was never provisioned
                                    }

                                    if (connectors == 1)
                                    {
                                        {
                                            csentry = afadMA.Connectors.ByIndex[0];
                                            csentry.Deprovision();
                                        }
                                    }

                                }
                            }
                            break;
                        }
                }
            }
            // Handle any exceptions
            catch (ObjectAlreadyExistsException objex)
            {
                // Ignore if the object already exists, join rules will join the existing object later
                // Capturing the Duplicate objects in Eventlog

                sSource = "MIIS AF-ADMA Provisioning";
                sLog = "Application";
                sEvent = objex.ToString();

                if (!EventLog.SourceExists(sSource))
                    EventLog.CreateEventSource(sSource, sLog);

                EventLog.WriteEntry(sSource, sEvent, EventLogEntryType.Error, 8001);

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

        bool IMVSynchronization.ShouldDeleteFromMV(CSEntry csentry, MVEntry mventry)
        {
            //
            // TODO: Add MV deletion logic here
            //
            throw new EntryPointNotImplementedException();
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

        public void setUserVariables(string paccode, string employeeType, string employeeSubGroup)
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
                if (employeeType == "D" && (employeeSubGroup == "63" || employeeSubGroup == "64" || employeeSubGroup == "65"))
                {
                    publicOu = "OU=Collaborators,OU=Domain Accounts,DC=amp,DC=icepoc,DC=com";
                }

            }
        }
    }
}