
using System;
using System.Xml;
using System.Diagnostics;
using Microsoft.MetadirectoryServices;

namespace Mms_ManagementAgent_KnownAsToolMAExtension
{
    /// <summary>
    /// Summary description for MAExtensionObject.
    /// </summary>
    public class MAExtensionObject : IMASynchronization
    {
        XmlNode rnode;
        XmlNode node;
        string toggledisplayname, toggledtag, env;

        public MAExtensionObject()
        {
            //
            // TODO: Add constructor logic here
            //
        }
        void IMASynchronization.Initialize()
        {
            try
            {
                const string XML_CONFIG_FILE = @"\rules-config.xml";
                XmlDocument config = new XmlDocument();
                string dir = Utils.ExtensionsDirectory;
                config.Load(dir + XML_CONFIG_FILE);

                rnode = config.SelectSingleNode("rules-extension-properties");
                node = rnode.SelectSingleNode("environment");
                env = node.InnerText;
                rnode = config.SelectSingleNode
                    ("rules-extension-properties/management-agents/" + env + "/ad-ma");
                //get the Users to be excluded from setting Network tag in displayname
                node = rnode.SelectSingleNode("toggledisplayname");
                toggledisplayname = node.InnerText.ToUpper();
                node = rnode.SelectSingleNode("toggledtag");
                toggledtag = node.InnerText;
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
            string sSource, sLog, sEvent;
            switch (FlowRuleName)
            {

                case "cd.person:KNWN_FRST_NM,KNWN_LST_NM,KNWN_MDL_NM,LGL_NM_FLG,SYSTEM_ID->mv.person:displayName":
                    //Displayname constructed and appeneded with " - Network" if the users are non-lilly
                    if (csentry["LGL_NM_FLG"].IsPresent && csentry["LGL_NM_FLG"].Value.ToString() == "1")
                    {
                        string displayname = "";
                        if (csentry["KNWN_FRST_NM"].IsPresent)
                            displayname = csentry["KNWN_FRST_NM"].Value;
                        if (csentry["KNWN_MDL_NM"].IsPresent)
                            displayname = displayname.Trim() + " " + csentry["KNWN_MDL_NM"].Value;
                        if (csentry["KNWN_LST_NM"].IsPresent)
                            displayname = displayname.Trim() + " " + csentry["KNWN_LST_NM"].Value;
                        if (mventry["employeeType"].IsPresent
                            && !toggledisplayname.Contains(mventry["employeeType"].Value.ToUpper())
                            && !displayname.Equals(""))
                        {
                            displayname = displayname.Trim() + toggledtag;
                        }
                        else if (!mventry["employeeType"].IsPresent && !displayname.Equals(""))
                        {
                            displayname = displayname.Trim() + toggledtag;
                        }

                        if (!displayname.Trim().Equals(""))
                        {
                            mventry["displayName"].Value = displayname.Trim();
                        }
                        else
                        {
                            mventry["displayName"].Delete();
                            //Write to the Event log for displayname is empty
                            sSource = "Known As Tool MA";
                            sLog = "Application";
                            sEvent = mventry["employeeID"].Value + " - Displayname is null";

                            if (!EventLog.SourceExists(sSource))
                                EventLog.CreateEventSource(sSource, sLog);

                            EventLog.WriteEntry(sSource, sEvent, EventLogEntryType.Error, 8102);
                        }
                    }
                    break;

                case "cd.person:KNWN_LST_NM,SYSTEM_ID->mv.person:sn":
                    if (csentry["KNWN_LST_NM"].IsPresent && csentry["KNWN_LST_NM"].Value.ToString() != string.Empty)
                    {
                        mventry["sn"].Value = csentry["KNWN_LST_NM"].Value.ToString();
                    }
                    break;

                case "cd.person:KNWN_FRST_NM,SYSTEM_ID->mv.person:givenName":
                    if (csentry["KNWN_FRST_NM"].IsPresent && csentry["KNWN_FRST_NM"].Value.ToString() != string.Empty)
                    {
                        mventry["givenName"].Value = csentry["KNWN_FRST_NM"].Value.ToString();
                    }
                    break;

                case "cd.person:KNWN_MDL_NM,SYSTEM_ID->mv.person:middleName":
                    if (csentry["KNWN_MDL_NM"].IsPresent && csentry["KNWN_MDL_NM"].Value.ToString() != string.Empty)
                    {
                        mventry["middleName"].Value = csentry["KNWN_MDL_NM"].Value.ToString();
                    }
                    break;

                default:
                    throw new EntryPointNotImplementedException();
            }
        }

        void IMASynchronization.MapAttributesForExport(string FlowRuleName, MVEntry mventry, CSEntry csentry)
        {
            switch (FlowRuleName)
            {
                case "cd.person:LGL_FRST_NM<-mv.person:givenName,knownAsFirstName":
                    if (!mventry["knownAsFirstName"].IsPresent)
                    {
                        csentry["LGL_FRST_NM"].Value = mventry["givenName"].Value.ToString();
                    }
                    break;

                case "cd.person:LGL_MDL_NM<-mv.person:knownAsMiddleName,middleName":
                    if (!mventry["knownAsMiddleName"].IsPresent)
                    {
                        csentry["LGL_MDL_NM"].Value = mventry["middleName"].Value.ToString();
                    }
                    break;

                case "cd.person:LGL_LST_NM<-mv.person:knownAsLastName,sn":
                    if (!mventry["knownAsLastName"].IsPresent)
                    {
                        csentry["LGL_LST_NM"].Value = mventry["sn"].Value.ToString();
                    }
                    break;

                default:
                    throw new EntryPointNotImplementedException();
            }
        }

    }
}
