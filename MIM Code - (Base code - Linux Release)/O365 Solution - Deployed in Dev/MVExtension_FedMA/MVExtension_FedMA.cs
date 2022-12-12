
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
        string ADAMgdiroU = string.Empty;
        string ADAMPrivAccrdiroU = string.Empty;
        string ADAMDSAcctdiroU = string.Empty;
        string ADAMServiceAcctdiroU = string.Empty;
        string timetolive;
        //CHG1471559 - Initial Password Update
        int MinInitPwdLength, MaxInitPwdLength;

        public MVExtensionObject()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        void IMVSynchronization.Initialize()
        //The function is used to assign the intial values to the variables that have scope in the entire modules.
        // Data is loaded from rules-config.xml file
        {
            XmlDocument config;
            XmlNode rnode, node;
            string env, dir;
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
                    ("rules-extension-properties/management-agents/" + env + "/adam");

                //Get the timetolive(ttyl) value from config file
                node = rnode.SelectSingleNode("ttl");
                timetolive = node.InnerText;

                node = rnode.SelectSingleNode("globaldiroU");
                ADAMgdiroU = node.InnerText; // Account OU for ADAM account
                node = rnode.SelectSingleNode("CAAcctdiroU");
                ADAMPrivAccrdiroU = node.InnerText; // Account OU for ADAM Privleged CA account
                node = rnode.SelectSingleNode("PrivAcctdiroU");
                ADAMDSAcctdiroU = node.InnerText;
                node = rnode.SelectSingleNode("ServiceAcctdiroU");
                ADAMServiceAcctdiroU = node.InnerText;

                //CHG1471559 - Initial Password Update
                node = rnode.SelectSingleNode("MinInitPwdLength");
                MinInitPwdLength = Int32.Parse(node.InnerText);

                //CHG1471559 - Initial Password Update
                node = rnode.SelectSingleNode("MaxInitPwdLength");
                MaxInitPwdLength = Int32.Parse(node.InnerText);

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

        void IMVSynchronization.Terminate()
        {
            //
            // TODO: Add termination logic here
            //
        }

        void IMVSynchronization.Provision(MVEntry mventry)
        {
            ///
            /// Provisioning accounts in ADAM.
            ///

            ConnectedMA FedMA, sadMA;
            CSEntry csentry;
            ReferenceValue dn = null;
            string cSEntryType = mventry.ObjectType;
            string rdn, sSource, sLog, sEvent, aNchor;
            int sadconnectors = 0, fedconnectors = 0;
            try
            {
                switch (cSEntryType)
                {
                    case "person":
                        {
                            if (!mventry["samAccountname"].IsPresent)
                            {
                                //ADAM account would be created only if samAccountname is present
                                //else do nothing
                            }
                            else
                            {

                                //Provisoning accounts in ADAM
                                FedMA = mventry.ConnectedMAs["Fed MA"];
                                //building RDN                                
                                aNchor = mventry["employeeID"].Value.Trim();
                                rdn = "CN=" + aNchor;
                                //Building a DN using ou and rdn
                                dn = FedMA.EscapeDNComponent(rdn).Concat(ADAMgdiroU);
                                fedconnectors = FedMA.Connectors.Count;
                                sadMA = mventry.ConnectedMAs["Staging Area Database MA"];
                                sadconnectors = sadMA.Connectors.Count;

                                if (sadconnectors == 1)//Account exist in SAD
                                {
                                    //NLW mini Release Change
                                    if (fedconnectors == 0)//Account doesn't exist in ADAM
                                    {
                                        //if (mventry["employeeType"].IsPresent && mventry["employeeType"].Value.ToUpper() == "D"
                                          // && mventry["emply_sub_grp_cd"].IsPresent
                                            // && (mventry["emply_sub_grp_cd"].Value == "61" || mventry["emply_sub_grp_cd"].Value == "62" || mventry["emply_sub_grp_cd"].Value == "63" || mventry["emply_sub_grp_cd"].Value == "64" || mventry["emply_sub_grp_cd"].Value == "65")
                                            // && mventry["System_Access_Flag"].IsPresent && mventry["System_Access_Flag"].Value.ToUpper() == "N")
										
										//HCM - Phase 1 - Modified the code to remvoe the Employee Sub Group Code attribute
										if (mventry["employeeType"].IsPresent && mventry["employeeType"].Value.ToUpper() == "D"									
										 && mventry["System_Access_Flag"].IsPresent && mventry["System_Access_Flag"].Value.ToUpper() == "N")
                                        {
                                            //Do nothing
                                        }
                                        else
                                        {
                                            csentry = FedMA.Connectors.StartNewConnector("user");
                                            csentry.DN = dn;
                                            // O365 bug fix - setting up initial password for new Users instead of "CHANGED" coming from Metaverse

                                            // Modifying initial password char limits for new AD password Policies
                                            csentry["unicodePwd"].Values.Add(RandomPassword.Generate(MinInitPwdLength, MaxInitPwdLength));
                                            //csentry["unicodePwd"].Values.Add(mventry["initPassword"].Value); //CR05786249 - ADAM User Account to have Password
                                            csentry.CommitNewConnector();
                                        }
                                    }
                                    else if (fedconnectors == 1)
                                    {
                                        // Account exist in ADAM, No need to check DN since it is built using employeeid
                                    }
                                    else
                                    {
                                        //Throw an execption if connectors are more than 1
                                        throw (new UnexpectedDataException("multiple connectors in ADAM :" + FedMA.Connectors.Count.ToString()));
                                    }
                                }

                                else if (sadconnectors == 0)
                                // Else clause if used for de-provisioing of accounts
                                {
                                    if (fedconnectors == 0)
                                    {
                                        //Do nothing if ADAM account was never provisioned
                                    }

                                    else
                                    {
                                        csentry = FedMA.Connectors.ByIndex[0];
                                        //This would perform a disconnect on the CSEntry. So next time when export is executed the record would be deleted from EDS directory
                                        //Deprovision method being called for ADAM MA
                                        if (AccountTTLExpired(mventry["deprovisionedDate"].Value, timetolive))
                                            csentry.Deprovision();
                                    }
                                }
                            }
                            break;
                        }

                    case "PrivilegedAccount":
                        {
                            if (!mventry["samAccountname"].IsPresent)
                            {
                                //ADAM account would be created only if samAccountname is present
                                //else do nothing
                            }
                            else
                            {

                                //Provisoning accounts in ADAM
                                FedMA = mventry.ConnectedMAs["Fed MA"];
                                //building RDN                                
                                aNchor = mventry["sAMAccountName"].Value.Trim();
                                rdn = "CN=" + aNchor;
                                //Building a DN using ou and rdn
                                if (mventry["sAMAccountName"].Value.ToLower().EndsWith("-ca"))
                                {
                                    dn = FedMA.EscapeDNComponent(rdn).Concat(ADAMPrivAccrdiroU);
                                }
                                else
                                {
                                    dn = FedMA.EscapeDNComponent(rdn).Concat(ADAMDSAcctdiroU);
                                }
                                //dn = FedMA.EscapeDNComponent(rdn).Concat(ADAMgdiroU);
                                fedconnectors = FedMA.Connectors.Count;

                                if (fedconnectors == 0)//Account doesn't exist in ADAM
                                {
                                    csentry = FedMA.Connectors.StartNewConnector("user");
                                    csentry.DN = dn;
                                    csentry["unicodePwd"].Values.Add(mventry["initpwdPrivlgdAcct"].Value);
                                    csentry.CommitNewConnector();
                                }
                                else if (fedconnectors == 1)
                                {
                                    // Account exist in ADAM, No need to check DN since it is built using employeeid
                                }
                                else
                                {
                                    //Throw an execption if connectors are more than 1
                                    throw (new UnexpectedDataException("multiple connectors in ADAM :" + FedMA.Connectors.Count.ToString()));
                                }

                            }

                            break;
                        }

                    case "ServiceAccounts":
                        {
                            if (!mventry["samAccountname"].IsPresent && !mventry["initpwdCloudServiceAcct"].IsPresent)
                            {
                                //ADAM account would be created only if samAccountname is present
                                //else do nothing
                            }
                            else
                            {

                                //Provisoning accounts in ADAM
                                FedMA = mventry.ConnectedMAs["Fed MA"];
                                //building RDN                                
                                aNchor = mventry["sAMAccountName"].Value.Trim();
                                rdn = "CN=" + aNchor;
                                //Building a DN using ou and rdn
                                dn = FedMA.EscapeDNComponent(rdn).Concat(ADAMServiceAcctdiroU);
                                fedconnectors = FedMA.Connectors.Count;

                                if (fedconnectors == 0)//Account doesn't exist in ADAM
                                {
                                    csentry = FedMA.Connectors.StartNewConnector("user");
                                    csentry.DN = dn;
                                    csentry["unicodePwd"].Values.Add(mventry["initpwdCloudServiceAcct"].Value);
                                    csentry.CommitNewConnector();
                                }
                                else if (fedconnectors == 1)
                                {
                                    // Account exist in ADAM, No need to check DN since it is built using employeeid
                                }
                                else
                                {
                                    //Throw an execption if connectors are more than 1
                                    throw (new UnexpectedDataException("multiple connectors in ADAM :" + FedMA.Connectors.Count.ToString()));
                                }

                            }

                            break;
                        }
                }
            }
            catch (ObjectAlreadyExistsException objex)
            {
                // Ignore if the object already exists, join rules will join the existing object later
                // Capturing the Duplicate objects in Eventlog

                sSource = "MIIS ADAM Provisioning";
                sLog = "Application";
                sEvent = objex.ToString();

                if (!EventLog.SourceExists(sSource))
                    EventLog.CreateEventSource(sSource, sLog);

                EventLog.WriteEntry(sSource, sEvent, EventLogEntryType.Error, 8051);

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
        bool IMVSynchronization.ShouldDeleteFromMV(CSEntry csentry, MVEntry mventry)
        {
            //
            // TODO: Add MV deletion logic here
            //
            throw new EntryPointNotImplementedException();
        }
    }
}
