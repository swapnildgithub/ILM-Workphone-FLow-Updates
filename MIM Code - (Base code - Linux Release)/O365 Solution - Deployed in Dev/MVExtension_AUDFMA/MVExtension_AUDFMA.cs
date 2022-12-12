
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

        void IMVSynchronization.Initialize ()
        {
            //
            // TODO: Add initialization logic here
            //
        }

        void IMVSynchronization.Terminate ()
        {
            //
            // TODO: Add termination logic here
            //
        }

        void IMVSynchronization.Provision (MVEntry mventry)
        {
            ConnectedMA AUDFMA, sadMA;  // Instance Creation for MAs
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
                            AUDFMA = mventry.ConnectedMAs["AUDF MA"];
                            connectors = AUDFMA.Connectors.Count; //Store the status for MA Connectivity

                            //Store Instance for 'AUDF MA'
                            sadMA = mventry.ConnectedMAs["Staging Area Database MA"];
                            sadconnectors = sadMA.Connectors.Count; //Store the status for MA Connectivity

                            if (sadconnectors == 1)
                            {
                                if (connectors == 0) // If record is new
                                {
                                    if (mventry["System_Access_Flag"].IsPresent)
                                    {
                                        if (mventry["System_Access_Flag"].Value.ToString().ToLower().Equals("y"))
                                        {
                                            csentry = AUDFMA.Connectors.StartNewConnector("person");    //Stores connector instance
                                            csentry["PRSNL_NBR"].Value = mventry["employeeID"].Value;   //Maps the ILM's entry 
                                            csentry["MIIS_MOD_DT"].Value = System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");  //Update Modification date
                                            csentry.CommitNewConnector();   //Committing the particular instance
                                        }
                                    }
                                }
                                else if (connectors == 1)
                                {
                                    // Ignore if there is already a connector   
                                    if (mventry["System_Access_Flag"].IsPresent)
                                    {
                                        if (mventry["System_Access_Flag"].Value.ToString().ToLower().Equals("n"))
                                        {
                                            csentry = AUDFMA.Connectors.ByIndex[0];
                                            //This would perform a disconnect on the CSEntry. So next time when export is executed the record would be deleted from SQL Server
                                            //Deprovision method being called for AUDF object only
                                            csentry.Deprovision();
                                        }
                                    }
                                }
                                else
                                {
                                    //Throw an execption if connectors are more than 1
                                    throw (new UnexpectedDataException("multiple connectors in AUDFMA:" + AUDFMA.Connectors.Count.ToString()));
                                }

                            }
                            else if (sadconnectors == 0)
                            // Else clause if used for de-provisioing of accounts
                            {
                                if (connectors == 0)
                                {
                                    //Do nothing if record was never provisioned
                                }

                                else
                                {
                                    csentry = AUDFMA.Connectors.ByIndex[0];
                                    //This would perform a disconnect on the CSEntry. So next time when export is executed the record would be deleted from SQL Server
                                    //Deprovision method being called for AUDF object only
                                    csentry.Deprovision();
                                }
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
                AUDFMA = null;
                csentry = null;
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
