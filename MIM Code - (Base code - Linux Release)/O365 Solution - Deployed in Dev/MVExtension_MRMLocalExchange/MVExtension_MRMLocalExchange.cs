
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

        void IMVSynchronization.Provision(MVEntry mventry)
        {

            const string PSAgentName = "MRM LocalExchange MA";
            const string SADAgentName = "Staging Area Database MA";
            const string RFAgentName = "Resource Forest AD MA";



            ConnectedMA managementAgent = mventry.ConnectedMAs[PSAgentName];
            ConnectedMA SADmanagementAgent = mventry.ConnectedMAs[SADAgentName];
            ConnectedMA RFmanagementAgent = mventry.ConnectedMAs[RFAgentName];



            int connectors = managementAgent.Connectors.Count;
            int SADconnectors = SADmanagementAgent.Connectors.Count;
            int RFconnectors = RFmanagementAgent.Connectors.Count;



            if (connectors == 0 && SADconnectors == 1 && RFconnectors == 1)
            {
                if (mventry["GetUserType"].IsPresent && mventry["GetUserType"].Value == "LocalMailbox")
                {
                    CSEntry csentry = managementAgent.Connectors.StartNewConnector("User");

                    csentry.DN = managementAgent.CreateDN("USER=" + mventry["sAMAccountName"].StringValue);

                    // set any additional mandatory attributes
                    if (mventry["sAMAccountName"].IsPresent)
                        csentry["sAMAccountName"].StringValue = mventry["sAMAccountName"].StringValue;
                    if (mventry["personnel_area_cd"].IsPresent)
                        csentry["PACCode"].StringValue = mventry["personnel_area_cd"].StringValue;


                    csentry.CommitNewConnector();
                }
            }

            else if (connectors == 1 && SADconnectors == 1)
            {
                CSEntry csentry = managementAgent.Connectors.ByIndex[0];

                csentry.DN = managementAgent.CreateDN("USER=" + mventry["sAMAccountName"].StringValue);

                csentry.CommitNewConnector();

            }
            else if (connectors > 1 && SADconnectors > 1)
            {
                string error = "Multiple connectors on the management agent";
                throw new UnexpectedDataException(error);
            }

            //throw new EntryPointNotImplementedException();
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
