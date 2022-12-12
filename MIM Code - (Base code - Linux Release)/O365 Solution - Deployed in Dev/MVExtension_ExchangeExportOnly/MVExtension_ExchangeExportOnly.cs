
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

            const string PSAgentName = "Exchange MA Export Only";
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
                CSEntry csentry = managementAgent.Connectors.StartNewConnector("User");

                csentry.DN = managementAgent.CreateDN("USER=" + mventry["sAMAccountName"].StringValue);

                // set any additional mandatory attributes
                if (mventry["Exch_Trans_CD"].IsPresent)
                    csentry["exchangeTransitionCode"].StringValue = mventry["Exch_Trans_CD"].StringValue;
                if (mventry["WhereToProvision"].IsPresent)
                    csentry["whereToProvision"].StringValue = mventry["WhereToProvision"].StringValue;
                if (mventry["deprovisionedDate"].IsPresent)
                    csentry["deprovisionedDate"].StringValue = mventry["deprovisionedDate"].StringValue;
                else
                    csentry["deprovisionedDate"].StringValue = "NA";
                if (mventry["Anglczd_first_nm"].IsPresent)
                    csentry["anglczd_first_nm"].StringValue = mventry["Anglczd_first_nm"].StringValue;
                if (mventry["Anglczd_last_nm"].IsPresent)
                    csentry["anglczd_last_nm"].StringValue = mventry["Anglczd_last_nm"].StringValue;
                if (mventry["GetUserType"].IsPresent)
                    csentry["getUserType"].StringValue = mventry["GetUserType"].StringValue;
                if (mventry["sAMAccountName"].IsPresent)
                    csentry["sAMAccountName"].StringValue = mventry["sAMAccountName"].StringValue;
                if (mventry["employeeID"].IsPresent)
                    csentry["employeeID"].StringValue = mventry["employeeID"].StringValue;
                if (mventry["msExchHomeServerName"].IsPresent)
                    csentry["msExchHomeServerName"].StringValue = mventry["msExchHomeServerName"].StringValue;
                if (mventry["homeMDB"].IsPresent)
                    csentry["homeMDB"].StringValue = mventry["homeMDB"].StringValue;
                if (mventry["mailNickname"].IsPresent)
                    csentry["mailNickname"].StringValue = mventry["mailNickname"].StringValue;
                if (mventry["domain"].IsPresent)
                    csentry["domain"].StringValue = mventry["domain"].StringValue;
                if (mventry["TargetAddress"].IsPresent)
                    csentry["targetAddress"].StringValue = mventry["TargetAddress"].StringValue;
                if (mventry["Business_Area_CD"].IsPresent)
                    csentry["businessAreaCode"].StringValue = mventry["Business_Area_CD"].StringValue;
                if (mventry["employeeType"].IsPresent)
                    csentry["employeeType"].StringValue = mventry["employeeType"].StringValue;
                if (mventry["org_unit_cd"].IsPresent)
                    csentry["org_unit_cd"].StringValue = mventry["org_unit_cd"].StringValue;
                if (mventry["System_Access_Flag"].IsPresent)
                    csentry["System_Access_Flag"].StringValue = mventry["System_Access_Flag"].StringValue;
                if (mventry["supervisor_flag"].IsPresent)
                    csentry["supervisor_flag"].StringValue = mventry["supervisor_flag"].StringValue;
                if (mventry["employeeStatus"].IsPresent)
                    csentry["employeeStatus"].StringValue = mventry["employeeStatus"].StringValue;
                if (mventry["personnel_area_cd"].IsPresent)
                    csentry["personnel_area_cd"].StringValue = mventry["personnel_area_cd"].StringValue;
                if (mventry["cost_center_cd"].IsPresent)
                    csentry["cost_center_cd"].StringValue = mventry["cost_center_cd"].StringValue;
                if (mventry["company_code"].IsPresent)
                    csentry["company_code"].StringValue = mventry["company_code"].StringValue;




                csentry.CommitNewConnector();

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
