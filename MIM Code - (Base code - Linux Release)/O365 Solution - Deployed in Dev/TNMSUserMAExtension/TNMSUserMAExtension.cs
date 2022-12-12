
using System;
using Microsoft.MetadirectoryServices;

namespace Mms_ManagementAgent_TNMSUserMAExtension
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
            switch (FlowRuleName)
            { 
                case "cd.person:GLOBAL_ID->mv.person:TNMSTransCode":
                    ConnectedMA tnmsDataMA = mventry.ConnectedMAs["TNMS Data MA"];
                    bool justDeleted = false;
                    int tnmsDataconnectors = tnmsDataMA.Connectors.Count;
                    if (tnmsDataconnectors == 1)
                    {
                        mventry["TNMSTransCode"].Value = "1";
                    
                    }
                    else
                    {
                        if (mventry["TNMSTransCode"].IsPresent && mventry["TNMSTransCode"].Value == "0")
                        {
                            mventry["TNMSTransCode"].Delete();
                            justDeleted = true;

                        }
                        if(justDeleted)
                        mventry["TNMSTransCode"].Value = "0";
                    
                    }


                    
                    break;
            }
        }

        void IMASynchronization.MapAttributesForExport (string FlowRuleName, MVEntry mventry, CSEntry csentry)
        {
            switch (FlowRuleName)
            {
               
                case "cd.person:TELEPHONE_VERFC<-mv.person:tnmsVerificationFlags":
                    
                    string strTelephoneNbr = getProperties(mventry["tnmsVerificationFlags"].Value, "LineURI");
                    if(!string.IsNullOrEmpty(strTelephoneNbr))
                    csentry["TELEPHONE_VERFC"].Value = strTelephoneNbr;
                    break;

                case "cd.person:POLICIES_VERFC<-mv.person:tnmsVerificationFlags":

                    string strUserPolicy = getProperties(mventry["tnmsVerificationFlags"].Value, "UserPolicy");
                    if (!string.IsNullOrEmpty(strUserPolicy))
                    csentry["POLICIES_VERFC"].Value = strUserPolicy;
                    break;

                case "cd.person:EVFLAG_VERFC<-mv.person:tnmsVerificationFlags":
                    string strEVFlag = getProperties(mventry["tnmsVerificationFlags"].Value, "EVFlag");
                    if(!string.IsNullOrEmpty(strEVFlag))
                    csentry["EVFLAG_VERFC"].Value = strEVFlag;
                    break;

             

                case "cd.person:UM<-mv.person:sAMAccountName,UC_Flag":

                    if (mventry["UC_Flag"].IsPresent &&
                       mventry["UC_Flag"].Value.ToUpper() == "Y"
                       )
                    {
                        csentry["UM"].BooleanValue = true;
                    }
                    else
                        csentry["UM"].BooleanValue = false;
                    break;

                case "cd.person:TNMS_TEL_USER_DTLS_STS<-mv.person:sAMAccountName":

                    ConnectedMA sadMA = mventry.ConnectedMAs["Staging Area Database MA"];
                    int sadconnectors = sadMA.Connectors.Count;

                    if(sadconnectors==0)
                        csentry["TNMS_TEL_USER_DTLS_STS"].BooleanValue = false;
                    else
                        csentry["TNMS_TEL_USER_DTLS_STS"].BooleanValue = true;

            break;

                default:
                    throw new EntryPointNotImplementedException();

            }
        }

        private string getProperties(string tnmsVerificationFlags, string propertyName)
        {
            string[] arrVerificationProperties = tnmsVerificationFlags.Split('&');
            string requiredProperty = "";
            foreach (string verificationProperty in arrVerificationProperties)
            {
                if (verificationProperty.Contains(propertyName))
                {
                    requiredProperty = verificationProperty;
                }
            }
            if(requiredProperty.Length>1)
            requiredProperty = requiredProperty.Split('&')[0].Substring(propertyName.Length+1);
            return requiredProperty;
        }

	}
}
