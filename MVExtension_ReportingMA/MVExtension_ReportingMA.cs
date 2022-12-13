
using System;
using Microsoft.MetadirectoryServices;


namespace Mms_Metaverse
{
    /// <summary>
    /// Summary description for MVExtensionObject.
    /// </summary>
    public class MVExtensionObject : IMVSynchronization
    {
        public MVExtensionObject()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        void IMVSynchronization.Initialize()
        {
            //
            // TODO: Add initialization logic here
            //
        }

        void IMVSynchronization.Terminate()
        {
            //
            // TODO: Add termination logic here
            //
        }

        void IMVSynchronization.Provision(MVEntry mventry)
        {
            ConnectedMA ReportingMA, sadMA;  // Instance Creation for MAs
            CSEntry csentry;            // Instance variable for storing SAD record

            string cSEntryType = mventry.ObjectType;    // Instance variable for storing Metaverse record
            int connectors, sadconnectors = 0;  //Initialization variable for connector

            try
            {
                switch (cSEntryType)
                {
                    case "person":
                        {
                            //Store Instance for 'AUDF MA'
                            ReportingMA = mventry.ConnectedMAs["Reporting MA"];
                            connectors = ReportingMA.Connectors.Count; //Store the status for MA Connectivity

                            //Store Instance for 'AUDF MA'
                            sadMA = mventry.ConnectedMAs["Staging Area Database MA"];
                            sadconnectors = sadMA.Connectors.Count; //Store the status for MA Connectivity

                            if (sadconnectors == 1)
                            {
                                if (connectors == 0) // If record is new
                                {
                                    //NLW Mini Release
                                    if (mventry["sAMAccountName"].IsPresent)
                                    {

                                        csentry = ReportingMA.Connectors.StartNewConnector("person");    //Stores connector instance
                                        csentry["SYSTEM_ID"].Value = mventry["sAMAccountName"].Value;   //Maps the ILM's entry 
                                        //csentry["MIIS_MOD_DT"].Value = System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");  //Update Modification date
                                        csentry.CommitNewConnector();   //Committing the particular instance

                                    }
                                }
                                else if (connectors == 1)
                                {
                                   

                                }
                                else
                                {
                                    //Throw an execption if connectors are more than 1
                                    throw (new UnexpectedDataException("multiple connectors in AUDFMA:" + ReportingMA.Connectors.Count.ToString()));
                                }

                            }
                            else if (sadconnectors == 0)
                            // Else clause if used for de-provisioing of accounts
                            {
                                //Do Nothing 
                            }

                        }

                        break;

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
                ReportingMA = null;
                csentry = null;
                sadMA = null;
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
