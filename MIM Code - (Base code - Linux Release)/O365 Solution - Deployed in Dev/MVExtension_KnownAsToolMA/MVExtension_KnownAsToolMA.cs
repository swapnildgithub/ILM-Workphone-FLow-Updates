
using System;
using Microsoft.MetadirectoryServices;
using CommonLayer_NameSpace;

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
            ConnectedMA knownAsToolMA, sadMA;
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
                                knownAsToolMA = mventry.ConnectedMAs["Known As Tool MA"];
                                connectors = knownAsToolMA.Connectors.Count;
                                sadMA = mventry.ConnectedMAs["Staging Area Database MA"];
                                sadconnectors = sadMA.Connectors.Count;
                                if (sadconnectors == 1) //Record exists in the SAD
                                {
                                    //CleanUp Release - Code modified
                                    if (connectors == 0) //Account doesn't exist in PD yet, to be inserted 
                                    {
                                        //NLW Mini Release
                                        if (CommonLayer.ISSysAccessFlagSet(mventry))
                                        {
                                            //User should be provisioned in Notification table only if SA flag is 'Y'
                                            if (mventry["employeeType"].IsPresent)
                                            {
                                                if (!mventry["deprovisionedDate"].IsPresent)
                                                {
                                                    csentry = knownAsToolMA.Connectors.StartNewConnector("person");
                                                    csentry["SYSTEM_ID"].Value = mventry["sAMAccountName"].Value.ToString();
                                                    csentry.CommitNewConnector();
                                                }
                                            }

                                        }
                                    }
                                    else if (connectors == 1)
                                    {
                                        csentry = knownAsToolMA.Connectors.ByIndex[0];

                                        //NLW Mini Release
                                        if (!CommonLayer.ISSysAccessFlagSet(mventry))
                                        {
                                            csentry.Deprovision();
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

        bool IMVSynchronization.ShouldDeleteFromMV(CSEntry csentry, MVEntry mventry)
        {
            //
            // TODO: Add MV deletion logic here
            //
            throw new EntryPointNotImplementedException();
        }
    }
}
