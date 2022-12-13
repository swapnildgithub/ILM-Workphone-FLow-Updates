using System;
using Microsoft.MetadirectoryServices;
using System.Collections;
using System.Xml;
using System.Diagnostics;
//using ActiveDs;

namespace Mms_Metaverse
{
    /// <summary>
    /// Extension for MetaVerse. Accounts provisoned from here.
    /// </summary>
    public class MVExtensionObject : IMVSynchronization
    {
        XmlNode rnode;
        XmlNode node;
        string version, parentOU, lcsflagoff, RFgrpoU;
        string uhomeMTA, umsExchHomeServer, uhomeMDB, archivingenabled, userlocationprofile, userinternetaccessenabled, userfederationenabled, useroptionflags;   //Exchange
        string[] strArrServerValues; //Exchange Random
        ArrayList userpolicy;
        Hashtable afvalhash;

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
                node = rnode.SelectSingleNode("version");
                version = node.InnerText;
                node = rnode.SelectSingleNode("parentOU");
                parentOU = node.InnerText;
                //get the Users to be set with LCS flag false
                node = rnode.SelectSingleNode("lcsflagoff");
                lcsflagoff = node.InnerText.ToUpper();

                node = rnode.SelectSingleNode("RFgrpOU");
                RFgrpoU = node.InnerText;

                XmlNode confignode = rnode.SelectSingleNode("siteconfigfile");
                afvalhash = LoadAFConfig(confignode.InnerText);

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
            /// Provisioning accounts in RF.
            ///

            ConnectedMA rfadMA, gmadma;
            ConnectedMA sadMA;
            CSEntry csentry;
            ReferenceValue dn;
            string cSEntryType = mventry.ObjectType;
            int connectors;
            string rdn, sSource, sLog, sEvent, formattedDate;
            int sadconnectors = 0, gmaconnectors = 0;

            try
            {
                switch (cSEntryType)
                {
                    case "person":
                        {
                            if (!mventry["cn"].IsPresent)
                            {
                                //Ensure CN is present or write to Event Log
                                string ExceptionMessage = mventry["employeeID"].Value
                                    + " - The attribute CN was not present on the MV";
                                sSource = "MIIS RF-ADMA Provisioning";
                                sLog = "Application";
                                sEvent = ExceptionMessage;

                                if (!EventLog.SourceExists(sSource))
                                    EventLog.CreateEventSource(sSource, sLog);

                                EventLog.WriteEntry(sSource, sEvent, EventLogEntryType.Error, 8012);

                                throw new UnexpectedDataException(ExceptionMessage);
                            }

                            else if (!mventry["samAccountname"].IsPresent)
                            {
                                //RF account would be created only if samAccountname is present
                                //else do nothing
                            }

                            else
                            {
                                //Provisoning accounts in RF
                                rfadMA = mventry.ConnectedMAs["Resource Forest AD MA"];
                                connectors = rfadMA.Connectors.Count;
                                sadMA = mventry.ConnectedMAs["Staging Area Database MA"];
                                sadconnectors = sadMA.Connectors.Count;

                                //Change CN value to System ID for Office 365
                                rdn = "CN=" + mventry["sAMAccountName"].Value.Trim();
                                //rdn = "CN=" + mventry["cn"].Value.Trim();
                                dn = rfadMA.EscapeDNComponent(rdn).Concat(parentOU);

                                if (sadconnectors == 1)
                                {
                                    if (connectors == 0)
                                    {
                                        csentry = rfadMA.Connectors.StartNewConnector("user");
                                        //Set the User ID
                                        csentry["sAMAccountName"].Value = mventry["sAMAccountName"].Value;

                                        formattedDate = DateTime.Now.ToString("dd-MMM-yyyy").ToUpper();
                                        //Log the changes
                                        csentry["info"].Delete();
                                        csentry["info"].Value = "[CREATED~MIISWRITER~" + formattedDate + "~" + version + "]";
                                        csentry.DN = dn;
                                        csentry.CommitNewConnector();

                                    }
                                    else if (connectors == 1)
                                    {
                                        // Check if the connector has a different DN and rename if necessary.
                                        // Get the connector.
                                        csentry = rfadMA.Connectors.ByIndex[0];
                                        //Microsoft Identity Integration Server will rename/move if different, if not, nothing will happen.
                                        csentry.DN = dn;
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
                                                csentry = rfadMA.Connectors.ByIndex[i];
                                                i++;
                                                //MIIS Created AD CS will be deleted so already available AD is the only CS for the user
                                                if (csentry["employeeID"].Value.Equals(mventry["employeeID"].Value)
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
                                                throw (new UnexpectedDataException("Multiple connectors in RF:" + rfadMA.Connectors.Count.ToString()));
                                            }
                                        }
                                        else
                                        {
                                            //Throw if more than 2 connectors or unable recover from above scenario
                                            throw (new UnexpectedDataException("Multiple connectors in RF:" + rfadMA.Connectors.Count.ToString()));
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
                                //Provisoning group accounts in RF
                                rfadMA = mventry.ConnectedMAs["Resource Forest AD MA"];
                                gmadma = mventry.ConnectedMAs["Group Management MA"];
                                //building RDN                                
                                rdn = "CN=" + mventry["sAMAccountName"].Value.Trim();
                                //Building a DN using ou and rdn
                                dn = rfadMA.EscapeDNComponent(rdn).Concat(RFgrpoU);
                                connectors = rfadMA.Connectors.Count;
                                gmadma = mventry.ConnectedMAs["Group Management MA"];
                                gmaconnectors = gmadma.Connectors.Count;

                                if (gmaconnectors == 1)
                                {
                                    if (connectors == 0)
                                    {
                                        csentry = rfadMA.Connectors.StartNewConnector("group");
                                        csentry.DN = dn;
                                        csentry.CommitNewConnector();
                                    }
                                    else if (connectors == 1)
                                    {
                                        // Check if the connector has a different DN and rename if necessary.
                                        // Get the connector.
                                        csentry = rfadMA.Connectors.ByIndex[0];
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
                                            csentry = rfadMA.Connectors.ByIndex[0];
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

                sSource = "MIIS RF-ADMA Provisioning";
                sLog = "Application";
                sEvent = objex.ToString();

                if (!EventLog.SourceExists(sSource))
                    EventLog.CreateEventSource(sSource, sLog);

                EventLog.WriteEntry(sSource, sEvent, EventLogEntryType.Error, 8002);
            }
            catch (AttributeNotPresentException)
            {
                // Ignore if the attribute on the mventry object is not available at this time
                // For example if atrbt_flg_1 isn't present then for those users will be ignored, 
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
                    }
                    valhash.Add("RANDOMIZEDARRAY", strArrRandomized);
                }
            }
            return primaryhash;
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


        public void setUserVariables(string paccode, string employeeType)
        {
            //
            // Initializing the variable necessary to build a RF account from the already initialized config file in a Hashtable
            //

            foreach (DictionaryEntry de in afvalhash)
            {
                Hashtable afconfig = (Hashtable)de.Value;
                ArrayList listhomeMDB = (ArrayList)afconfig["homeMDB"];
                ArrayList listmsexchange = (ArrayList)afconfig["msExchHomeServerName"];
                ArrayList listarea = (ArrayList)afconfig["PERSONNELAREA"];
                ArrayList listtype = (ArrayList)afconfig["CONSTITUENT"];
                uhomeMDB = null;
                uhomeMTA = null;
                umsExchHomeServer = null;
                strArrServerValues = null;

                if (listarea.Contains(paccode) && listtype.Contains(employeeType.ToUpper()))
                {
                    uhomeMDB = listhomeMDB[0].ToString();
                    uhomeMTA = afconfig["homeMTA"].ToString();
                    umsExchHomeServer = listmsexchange[0].ToString();
                    archivingenabled = afconfig["ARCHIVINGENABLED"].ToString();        //Exchange OCS code
                    userlocationprofile = afconfig["USERLOCATIONPROFILE"].ToString();  //Exchange OCS code
                    userpolicy = (ArrayList)afconfig["USERPOLICY"];     //Exchange OCS code
                    userinternetaccessenabled = afconfig["INTERNETACCESSENABLED"].ToString();        //Exchange OCS code
                    userfederationenabled = afconfig["FEDERATIONENABLED"].ToString();        //Exchange OCS code
                    useroptionflags = afconfig["OPTIONFLAGS"].ToString();        //Exchange OCS code
                    strArrServerValues = getRandomizedValues((string[,])afconfig["RANDOMIZEDARRAY"]);
                    break;
                }
            }
        }
    }
}