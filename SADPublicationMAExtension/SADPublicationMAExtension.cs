
using System;
using Microsoft.MetadirectoryServices;
using CommonLayer_NameSpace;

namespace Mms_ManagementAgent_SADPublicationMAExtension
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
            throw new EntryPointNotImplementedException();
        }

        void IMASynchronization.MapAttributesForExport (string FlowRuleName, MVEntry mventry, CSEntry csentry)
        {
            /// Export attribute flow 

            switch (FlowRuleName)
            {

                case "cd.person:PROXY_ADDRESS<-mv.person:proxyAddresses,sAMAccountName":

                    ValueCollection finalValues = Utils.ValueCollection("initialValue");
                    string proxyaddresslist = null;
                    int intTempCount = 1;
                    finalValues.Clear();


                    foreach (Value addrElement in mventry["proxyAddresses"].Values)
                    {
                        if (intTempCount < mventry["proxyAddresses"].Values.Count)
                        {
                            proxyaddresslist = proxyaddresslist + addrElement.ToString() + ";";
                        }
                        else
                        {
                            proxyaddresslist = proxyaddresslist + addrElement.ToString();
                        }
                        intTempCount++;
                    }

                    if ((proxyaddresslist == null) || (proxyaddresslist == ""))
                        csentry["PROXY_ADDRESS"].Delete();
                    else
                        csentry["PROXY_ADDRESS"].Value = proxyaddresslist;
                    
                    break;

                case "cd.person:INTERNET_STYLE_EMAIL_ADRS<-mv.person:Ext_Auth_Flag,mail,msExchHideFromAddressList,msExchMailboxGuid,msExchRecipientTypeDetails,sAMAccountName":

                    if (mventry["Ext_Auth_Flag"].IsPresent)
                    {
                        if (mventry["Ext_Auth_Flag"].Value.ToString().ToLower().Equals("n") || CommonLayer.IsMigrationOverride(mventry,csentry,"SADPUB"))
                        {
                            if (mventry["mail"].IsPresent)
                                if (!(mventry["msExchHideFromAddressList"].IsPresent && mventry["msExchHideFromAddressList"].BooleanValue == true))
                                    csentry["INTERNET_STYLE_EMAIL_ADRS"].Value = mventry["mail"].Value;
                                else
                                    csentry["INTERNET_STYLE_EMAIL_ADRS"].Delete();
                            else
                                csentry["INTERNET_STYLE_EMAIL_ADRS"].Delete();
                        }
                        else
                        {
                            csentry["INTERNET_STYLE_EMAIL_ADRS"].Delete();
                        }
                    }
                    else
                    {
                        if (mventry["mail"].IsPresent)
                            if (!(mventry["msExchHideFromAddressList"].IsPresent && mventry["msExchHideFromAddressList"].BooleanValue == true))
                                csentry["INTERNET_STYLE_EMAIL_ADRS"].Value = mventry["mail"].Value;
                            else
                                csentry["INTERNET_STYLE_EMAIL_ADRS"].Delete();
                    }
                    break;
                
                default:
                    throw new EntryPointNotImplementedException();
            }
        }
	}
}
