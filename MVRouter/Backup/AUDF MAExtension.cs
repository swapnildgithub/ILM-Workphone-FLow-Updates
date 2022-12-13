
using System;
using Microsoft.MetadirectoryServices;

namespace Mms_ManagementAgent_AUDF_MAExtension
{
    /// <summary>
    /// Summary description for MAExtensionObject.
	/// </summary>
	public class MAExtensionObject : IMASynchronization
	{
		public MAExtensionObject()
		{
            //
            // TODO: Add constructor logic here
            //
        }
		void IMASynchronization.Initialize ()
		{
            //
            // TODO: write initialization code
            //
        }

        void IMASynchronization.Terminate ()
        {
            //
            // TODO: write termination code
            //
        }

        bool IMASynchronization.ShouldProjectToMV (CSEntry csentry, out string MVObjectType)
        {
			//
			// TODO: Remove this throw statement if you implement this method
			//
			throw new EntryPointNotImplementedException();
		}

        DeprovisionAction IMASynchronization.Deprovision (CSEntry csentry)
        {
			//
			// TODO: Remove this throw statement if you implement this method
			//
			throw new EntryPointNotImplementedException();
        }	

        bool IMASynchronization.FilterForDisconnection (CSEntry csentry)
        {
            //
            // TODO: write connector filter code
            //
            throw new EntryPointNotImplementedException();
		}

		void IMASynchronization.MapAttributesForJoin (string FlowRuleName, CSEntry csentry, ref ValueCollection values)
        {
            //
            // TODO: write join mapping code
            //
            throw new EntryPointNotImplementedException();
        }

        bool IMASynchronization.ResolveJoinSearch (string joinCriteriaName, CSEntry csentry, MVEntry[] rgmventry, out int imventry, ref string MVObjectType)
        {
            //
            // TODO: write join resolution code
            //
            throw new EntryPointNotImplementedException();
		}

        void IMASynchronization.MapAttributesForImport( string FlowRuleName, CSEntry csentry, MVEntry mventry)
        {
            //
            // TODO: write your import attribute flow code
            //
            switch (FlowRuleName)
            {
                case "cd.person:CNTRY_CD,EMPLY_GRP_CD,EMPLY_SUB_GRP_CD,PRSNL_NBR,STATUS_CD,SYSTEM_ID->mv.person:miis_mod_dt":
                    {
                        mventry["miis_mod_dt"].Value = System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"); //Stores Updated Modification date
                    }
                    break;

                default:
                    throw new EntryPointNotImplementedException();
            }
        }

        void IMASynchronization.MapAttributesForExport (string FlowRuleName, MVEntry mventry, CSEntry csentry)
        {
            //
			// TODO: write your export attribute flow code
			//
            switch (FlowRuleName)
			{
                case "cd.person:ALT_CNTCT_PRSNL_NBR<-mv.person:EDS_Alt_Cntct_Prsnl_Nbr,sAMAccountName":
                   
					if (!(mventry["EDS_Alt_Cntct_Prsnl_Nbr"].IsPresent))
                    {
                        csentry["ALT_CNTCT_PRSNL_NBR"].Value = "0";
                    }
                    break;

				default:
					// TODO: remove the following statement and add your default script here
					throw new EntryPointNotImplementedException();
			}
        }
	}
}
