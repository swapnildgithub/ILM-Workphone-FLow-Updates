
using System;
using Microsoft.MetadirectoryServices;
using System.Collections;
using System.Xml;
using System.Diagnostics;

namespace Mms_Metaverse
{
	/// <summary>
	/// Summary description for MVExtensionObject.
	/// </summary>
    public class MVExtensionObject : IMVSynchronization
    {
        string accountOu = "";
        string workerOu = "";

        public MVExtensionObject()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        void IMVSynchronization.Initialize ()
        //The function is used to assign the intial values to the variables that have scope in the entire modules.
        // Data is loaded from rules-config.xml file
        {
            XmlDocument config;
            XmlNode rnode;
            XmlNode node;
            string dir;
            string env;
            try
            {
                const string XML_CONFIG_FILE = @"\rules-config.xml";
                config = new XmlDocument();
                dir = Utils.ExtensionsDirectory;
                config.Load(dir + XML_CONFIG_FILE);

                rnode = config.SelectSingleNode("rules-extension-properties");
                node = rnode.SelectSingleNode("environment");
                env = node.InnerText;
                rnode = config.SelectSingleNode
                    ("rules-extension-properties/management-agents/" + env + "/eds-ma");
                node = rnode.SelectSingleNode("accountou");
                accountOu = node.InnerText; // Account OU for EDS internal account
                node = rnode.SelectSingleNode("workerou");
                workerOu = node.InnerText; // Worker OU for EDS internal account
            }
            catch (NullReferenceException nre)
            {
                //	If a tag does not exist in the xml, the stopped-extension-dll 
                //	error will be thrown
                throw nre;
            }
            catch (Exception e)
            {
                //	The exception would be evident on the operation log of MIIS
                throw e;
            }
            finally
            {
                config = null;
                rnode = null;
                node = null;
            }
        }

        void IMVSynchronization.Terminate ()
        {
            //
            // TODO: Add termination logic here
            //
        }

        void IMVSynchronization.Provision (MVEntry mventry)
        {
            // Provisioning accounts in Internal EDS account branch
         
            ConnectedMA intEDSMA, sadMA;
            CSEntry csentry;
            ReferenceValue dn;
            string cSEntryType = mventry.ObjectType;
            int connectors = 0, sadconnectors = 0;;
            string aNchor, rdn, container;
            

            try
            {
                switch (cSEntryType)
                {
                    case "person":
                        {
                            if (!mventry["samAccountname"].IsPresent)
                            {
                                //If samAccountname doesn't exist for the user don't create an account branch.
                                //No need to log any error to event viewer as there would be many such accounts
                            }
                            else
                            {
                                intEDSMA = mventry.ConnectedMAs["Internal EDS MA"];
                                aNchor = mventry["sAMAccountName"].Value.Trim();
                                rdn = "uid=" + aNchor;
                                container = accountOu;
                                dn = intEDSMA.EscapeDNComponent(rdn).Concat(container);
                                connectors = intEDSMA.Connectors.Count;

                                sadMA = mventry.ConnectedMAs["Staging Area Database MA"];
                                sadconnectors = sadMA.Connectors.Count;

                                if (sadconnectors == 1) //Record exists in the SAD
                                {
                                    //NLW mini Release Change
                                    if (connectors == 0) //Account doesn't exist in EDS yet, to be created
                                    {
                                        if (mventry["employeeType"].IsPresent && mventry["employeeType"].Value.ToUpper() == "D"
                                             && mventry["emply_sub_grp_cd"].IsPresent
                                             && (mventry["emply_sub_grp_cd"].Value == "61" || mventry["emply_sub_grp_cd"].Value == "62" || mventry["emply_sub_grp_cd"].Value == "63" || mventry["emply_sub_grp_cd"].Value == "64" || mventry["emply_sub_grp_cd"].Value == "65")
                                             && mventry["System_Access_Flag"].IsPresent && mventry["System_Access_Flag"].Value.ToUpper() == "N")
                                        {
                                            //Do nothing
                                        }
                                        else
                                        {
                                            ValueCollection oc;
                                            oc = Utils.ValueCollection("top");
                                            oc.Add("simpleSecurityObject");
                                            oc.Add("EdsAccount");

                                            csentry = intEDSMA.Connectors.StartNewConnector(
                                                                                "EdsAccount", oc);
                                            byte[] rawpw = System.Text.UTF8Encoding.UTF8.GetBytes(mventry["initPassword"].Value);//Provisoning Password in Internal EDS account MA
                                            csentry["userPassword"].Values.Add(rawpw);
                                            csentry.DN = dn;
                                            //Only mandotary key attribute(s) has been built from code
                                            //All other attributes should be populated from export flow to make the code execution faster
                                            csentry["EdsAccntActiveFlag"].Value = "Y";
                                            csentry["EdsPasswdCounter"].Value = "0";
                                            csentry["EdsPasswdExpiredFlag"].Value = "N";
                                            csentry.CommitNewConnector();
                                        }
                                    }
                                    else if (connectors == 1)
                                    {
                                        // Check if the connector has a different DN and rename if necessary.
                                        // Get the connector.
                                        csentry = intEDSMA.Connectors.ByIndex[0];
                                        csentry.DN = dn;
                                    }
                                    else
                                    {
                                        //Throw an execption if connectors are more than 1
                                        throw (new UnexpectedDataException("multiple connectors in Internal EDS Account MA:" + intEDSMA.Connectors.Count.ToString()));
                                    }
                                }

                                else if (sadconnectors == 0)
                                // Else clause if used for de-provisioing of accounts
                                {
                                    if (connectors == 0)
                                    {
                                        //Do nothing if EDS account was never provisioned
                                    }

                                    else
                                    {
                                        csentry = intEDSMA.Connectors.ByIndex[0];
                                        //This would perform a disconnect on the CSEntry. So next time when export is executed the record would be deleted from EDS directory
                                        //Deprovision method being called for Internal EDS MA object only
                                        csentry.Deprovision();
                                    }
                                }
                              }

                                break;
                        }
                }
            }
            // Handle any exceptions
            catch (AttributeNotPresentException)
            {
                // Ignore
            }
            catch (ObjectAlreadyExistsException)
            {
                // Ignore
            }
            catch (NoSuchAttributeException)
            {
                // Ignore if the attribute on the mventry object is not available at this time
            }
            catch (Exception ex)
            {
                // All other exceptions re-throw to rollback the object synchronization transaction and 
                // report the error to the run history
                throw ex;
            }
            finally
            {
                intEDSMA = null;
                csentry = null;
                dn = null;
                sadMA = null;
            }
        }	

        bool IMVSynchronization.ShouldDeleteFromMV (CSEntry csentry, MVEntry mventry)
        {
            //
            // TODO: Add MV deletion logic here
            //
            throw new EntryPointNotImplementedException();
        }
    }
}
