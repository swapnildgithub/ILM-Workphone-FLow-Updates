
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
            ConnectedMA SADPubMA, AUDFMA, sadMA;   // Instance Creation for MAs
            CSEntry csentry;                       // Instance variable for storing SAD record

            string cSEntryType = mventry.ObjectType;    // Instance variable for storing Metaverse record
            int connectors = 0, sadconnectors = 0, audfconnectors = 0;   //Initialization variable for connector            

            try
            {
                switch (cSEntryType)
                {
                    case "person":
                        {
                            //Store Instance for 'SAD Publication MA'
                            SADPubMA = mventry.ConnectedMAs["SAD Publication MA"];
                            connectors = SADPubMA.Connectors.Count;    //Store the status for MA Connectivity
                            //Store Instance for 'AUDF MA'
                            AUDFMA = mventry.ConnectedMAs["AUDF MA"];
                            audfconnectors = AUDFMA.Connectors.Count;   //Store Instance for 'AUDF MA'
                            //Store Instance for 'AUDF Publication MA'
                            sadMA = mventry.ConnectedMAs["Staging Area Database MA"];
                            sadconnectors = sadMA.Connectors.Count;     //Store Instance for 'Staging Area Database MA'
                            if (connectors == 0)
                            {   //NLW mini Release Change
                                if (sadconnectors == 1)
                                {
                                   // if (mventry["employeeType"].IsPresent && mventry["employeeType"].Value.ToUpper() == "D"
                                     //        && mventry["emply_sub_grp_cd"].IsPresent
                                     //        && (mventry["emply_sub_grp_cd"].Value == "61" || mventry["emply_sub_grp_cd"].Value == "62" || mventry["emply_sub_grp_cd"].Value == "63" || mventry["emply_sub_grp_cd"].Value == "64" || mventry["emply_sub_grp_cd"].Value == "65")
                                     //        && mventry["System_Access_Flag"].IsPresent && mventry["System_Access_Flag"].Value.ToUpper() == "N")
									 
									 //HCM - Phase 1 - Modified the code to remvoe the Employee Sub Group Code attribute
									if(mventry["employeeType"].IsPresent && mventry["employeeType"].Value.ToUpper() == "D" 
									 && mventry["System_Access_Flag"].IsPresent && mventry["System_Access_Flag"].Value.ToUpper() == "N")
                                    {
                                        //Do nothing
                                    }
                                    else
                                    {
                                        csentry = SADPubMA.Connectors.StartNewConnector("person");
                                        csentry["PRSNL_NBR"].Value = mventry["employeeID"].Value;
                                        csentry.CommitNewConnector();
                                    }
                                }
                            }
                            else if (connectors == 1)
                            {
                                //Do nothing
                            }
                            else
                            {
                                throw (new UnexpectedDataException("Multiple connectors in AUDF PUB MA:" + SADPubMA.Connectors.Count.ToString()));
                            }
                            if(audfconnectors == 0)
                            {
                                //if record is deleted from AUDF and SAD then should be permanently deleted
                                if (sadconnectors == 0)
                                {
                                    if (connectors == 1)
                                    {
                                        csentry = SADPubMA.Connectors.ByIndex[0];
                                        csentry.Deprovision();
                                    }
                                }
                            }
                            break;
                        }
                }
            }
            // Handle any exceptions
            catch (ObjectAlreadyExistsException)
            {
                // Ignore if the object already exists, join rules will join the existing object later
            }
            catch (AttributeNotPresentException)
            {
                // Ignore if the attribute on the mventry object is not available at this time
                // For example if employee type isn't present then for those users will be ignored, 
                // This exception is used instead of using isPresent() method.
            }
            catch (Exception ex)
            {
                // All other exceptions re-throw to rollback the object synchronization transaction and 
                // report the error to the run history
                throw ex;
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
