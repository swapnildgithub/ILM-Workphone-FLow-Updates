
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

            const string UPNAgentName = "UserPrincipalName MA Export Only";
            const string SADAgentName = "Staging Area Database MA";
            const string RFAgentName = "Resource Forest AD MA";



            ConnectedMA managementAgent = mventry.ConnectedMAs[UPNAgentName];
            ConnectedMA SADmanagementAgent = mventry.ConnectedMAs[SADAgentName];
            ConnectedMA RFmanagementAgent = mventry.ConnectedMAs[RFAgentName];



            int connectors = managementAgent.Connectors.Count;
            int SADconnectors = SADmanagementAgent.Connectors.Count;
            int RFconnectors = RFmanagementAgent.Connectors.Count;



            if (connectors == 0 && SADconnectors == 1 && RFconnectors == 1)
            {
                if (mventry["GetUserType"].IsPresent && mventry["GetUserType"].Value != "None")
                {
                    CSEntry csentry = managementAgent.Connectors.StartNewConnector("User");
                    csentry.DN = managementAgent.CreateDN("USER=" + mventry["sAMAccountName"].StringValue);
                    if (mventry["mail"].IsPresent)
                        csentry["mail"].StringValue = mventry["mail"].StringValue;
                    if (mventry["sAMAccountName"].IsPresent)
                        csentry["saMAccountName"].StringValue = mventry["sAMAccountName"].StringValue;
                    if (mventry["GetUserType"].IsPresent)
                        csentry["userType"].StringValue = mventry["GetUserType"].StringValue;
                    if (mventry["domain"].IsPresent)
                        csentry["domain"].StringValue = mventry["domain"].StringValue;
                    if (mventry["Business_Area_CD"].IsPresent)
                        csentry["Business_Area_CD"].StringValue = mventry["Business_Area_CD"].StringValue;
                    if (mventry["employeeType"].IsPresent)
                        csentry["employeeType"].StringValue = mventry["employeeType"].StringValue;
                    if (mventry["org_unit_cd"].IsPresent)
                        csentry["org_unit_cd"].StringValue = mventry["org_unit_cd"].StringValue;
                    if (mventry["userPrincipalName"].IsPresent)
                        csentry["userPrincipalName"].StringValue = mventry["userPrincipalName"].StringValue;
                    if (mventry["updateRecipientFlag"].IsPresent)
                        csentry["OldUPNDetails"].StringValue = mventry["updateRecipientFlag"].StringValue;




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

        bool IMVSynchronization.ShouldDeleteFromMV(CSEntry csentry, MVEntry mventry)
        {
            //
            // TODO: Add MV deletion logic here
            //
            throw new EntryPointNotImplementedException();
        }
    }
}
