
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
            ConnectedMA pdMA, sadMA;
            CSEntry csentry;
            string cSEntryType = mventry.ObjectType;
            int connectors;
            int sadconnectors = 0;

            try
            {
                switch (cSEntryType)
                {
                    case "person":
                        {

                            if (!mventry["samAccountname"].IsPresent)
                            {
                                //If samAccountname doesn't exist for the user don't insert record in PD table
                            }
                            else
                            {
                                pdMA = mventry.ConnectedMAs["Notification MA"];
                                connectors = pdMA.Connectors.Count;
                                sadMA = mventry.ConnectedMAs["Staging Area Database MA"];
                                sadconnectors = sadMA.Connectors.Count;
                                if (sadconnectors == 1) //Record exists in the SAD
                                {

                                    //CleanUp Release - Code modified
                                    if (connectors == 0) //Account doesn't exist in PD yet, to be inserted 
                                    {
                                        //User should be provisioned in Notification table only if SA flag is 'Y'
                                        if (mventry["System_Access_Flag"].IsPresent)
                                        {
                                            if (mventry["System_Access_Flag"].Value.ToString().ToLower().Equals("y"))
                                            {
                                                csentry = pdMA.Connectors.StartNewConnector("person");
                                                csentry["PRSNL_NBR"].Value = mventry["employeeID"].Value.ToString();
                                                csentry.CommitNewConnector();
                                            }
                                        }
                                    }
                                    else if (connectors == 1)
                                    {
                                        // Ignore if there is already a connector   
                                        if (mventry["System_Access_Flag"].IsPresent)
                                        {
                                            //User should be de-provisioned from Notification table if SA flag is 'n'
                                            //CleanUp Release - Code modified
                                            if (mventry["System_Access_Flag"].Value.ToString().ToLower().Equals("n"))
                                            {
                                                csentry = pdMA.Connectors.ByIndex[0];
                                                //This would perform a disconnect on the CSEntry. So next time when export is executed the record would be deleted from SQL Server
                                                //Deprovision method being called for Notification object only
                                                csentry.Deprovision();
                                            }
                                        }
                                    }
                                }
                            }
                            break;
                        }
                 }
            }
            catch (Exception ex)
            {
                // All other exceptions re-throw to rollback the object synchronization transaction and 
                // report the error to the run history
                throw ex;
            }
            //
			// TODO: Remove this throw statement if you implement this method
			//
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
