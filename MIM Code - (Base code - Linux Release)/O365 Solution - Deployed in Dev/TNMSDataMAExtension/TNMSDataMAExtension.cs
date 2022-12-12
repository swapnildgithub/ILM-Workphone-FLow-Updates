
using System;
using System.Xml;
using Microsoft.MetadirectoryServices;
using System.Data;
using System.Collections;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;


namespace Mms_ManagementAgent_TNMSDataMAExtension
{
    /// <summary>
    /// Summary description for MAExtensionObject.
    /// </summary>
    public class MAExtensionObject : IMASynchronization
    {
        string TNTypes, LEVType;
        XmlNode rnode;
        XmlNode node;
        public MAExtensionObject()
        {
            //
            // TODO: Add constructor logic here
            //
        }
        void IMASynchronization.Initialize()
        {
            //
            // Initialize config details from xml file
            //
            try
            {
                const string XML_CONFIG_FILE = @"\rules-config.xml";
                XmlDocument config = new XmlDocument();
                string dir = Utils.ExtensionsDirectory;
                config.Load(dir + XML_CONFIG_FILE);


                rnode = config.SelectSingleNode("rules-extension-properties");
                node = rnode.SelectSingleNode("environment");
                string env = node.InnerText;
                rnode = config.SelectSingleNode
                    ("rules-extension-properties/management-agents/" + env + "/ad-ma");

                node = rnode.SelectSingleNode("TNTypes");
                TNTypes = node.InnerText;
                node = rnode.SelectSingleNode("LEVType");
                LEVType = node.InnerText;

            }
            catch (NullReferenceException nre)
            {
                // If a tag does not exist in the xml, then the stopped-extension-dll 
                // error will be thrown.
                throw nre;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        void IMASynchronization.Terminate()
        {
            //
            // TODO: write termination code
            //
        }

        bool IMASynchronization.ShouldProjectToMV(CSEntry csentry, out string MVObjectType)
        {
            //
            // TODO: Remove this throw statement if you implement this method
            //
            throw new EntryPointNotImplementedException();
        }

        DeprovisionAction IMASynchronization.Deprovision(CSEntry csentry)
        {
            //
            // TODO: Remove this throw statement if you implement this method
            //
            throw new EntryPointNotImplementedException();
        }

        bool IMASynchronization.FilterForDisconnection(CSEntry csentry)
        {
            //
            // TODO: write connector filter code
            //
            throw new EntryPointNotImplementedException();
        }

        void IMASynchronization.MapAttributesForJoin(string FlowRuleName, CSEntry csentry, ref ValueCollection values)
        {
            if (FlowRuleName == "cd.person#1:SYSTEM_ID->sAMAccountName")
            {
                if (csentry["SYSTEM_ID"].Value.ToString().ToUpper() != "N/A")
                {
                    String strsAMAccountName = csentry["SYSTEM_ID"].Value.ToString();
                    values.Add(strsAMAccountName);
                }
            }
        }

        bool IMASynchronization.ResolveJoinSearch(string joinCriteriaName, CSEntry csentry, MVEntry[] rgmventry, out int imventry, ref string MVObjectType)
        {

            if (joinCriteriaName == "cd.person#1")
            {
                if (csentry["SYSTEM_ID"].Value.ToString().ToUpper() == rgmventry[0]["sAMAccountName"].Value.ToString().ToUpper())
                {
                    imventry = 0;
                    return true;
                }
                else
                {
                    imventry = 0;
                    return false;
                }
            }
            imventry = 0;
            return true;
        }

        void IMASynchronization.MapAttributesForImport(string FlowRuleName, CSEntry csentry, MVEntry mventry)
        {
            switch (FlowRuleName)
            {
                case "cd.person:CNTRY_PHN_CD,SYSTEM_ID->mv.person:tnmsCountryCode":
                    if (csentry["CNTRY_PHN_CD"].IsPresent)
                    {
                        mventry["tnmsCountryCode"].Value = csentry["CNTRY_PHN_CD"].Value.ToString();
                    }
                    break;

                case "cd.person:CNTRY_PHN_CD,SYSTEM_ID,TELEPHONE_NBR->mv.person:tnmsReadableFormatTelephone":
                    if (mventry["tnmsCountryCode"].IsPresent && mventry["tnmsTelephoneNumber"].IsPresent)
                    {
                        string tnmsCountryCode = "+" + mventry["tnmsCountryCode"].Value;
                        string tnmsTelephoneNumber = mventry["tnmsTelephoneNumber"].Value;
                        string tnmsReadableFormatTelephone = tnmsTelephoneNumber.Insert(tnmsCountryCode.Length, "-");
                        mventry["tnmsReadableFormatTelephone"].Value = tnmsReadableFormatTelephone;
                    }
                    break;


                case "cd.person:CLIENT_POLICY,DIAL_PLAN,EV_FLAG,EVFLAG_VERIFICATION,LINE_URI,MAILBOX_POLICY,POLICY_VERIFICATION,REGISTRAR_SERVER,TELEPHONE_TYPE,TELEPHONE_VERIFICATION,UMENABLE_VERIFICATION,VOICE_POLICY->mv.person:tnmsChangedFlags":

                    string strEA14 = string.Empty;


                    if (mventry["EA14"].IsPresent)
                    {
                        strEA14 = mventry["EA14"].Value;
                    }

                    string[] EA14Values;
                    string existingEVFlag = string.Empty,
                           existingDialPlan = string.Empty,
                           existingVoicePolicy = string.Empty,
                           existingRegistrarServer = string.Empty,
                           existingClientPolicy = string.Empty,
                           existingLineURI = string.Empty,
                           existingOVAValue = string.Empty,
                           existingMBPolicyValue = string.Empty,
                           existingLyncValue = string.Empty,
                           existingUMValue = string.Empty;

                    if (strEA14.Length > 2)
                    {
                        EA14Values = strEA14.Split(',')[1].Split(':');

                        if (EA14Values.Length > 6)
                        {
                            existingEVFlag = (EA14Values[0] == "1") ? "True" : "False";
                            existingDialPlan = EA14Values[1];
                            existingVoicePolicy = EA14Values[2];
                            existingRegistrarServer = EA14Values[3];
                            existingLineURI = EA14Values[4];
                            existingClientPolicy = EA14Values[5];
                            existingOVAValue = EA14Values[6];
                            existingMBPolicyValue = EA14Values[7];

                        }
                    }
                    string strmsExchUMEnabledFlags = "0";
                    if (mventry["msExchUMEnabledFlags"].IsPresent)
                        strmsExchUMEnabledFlags = mventry["msExchUMEnabledFlags"].Value.ToString();
                    else
                        strmsExchUMEnabledFlags = "0";


                    string strTelephonyType = "0";
                    if (csentry["TELEPHONE_TYPE"].IsPresent)
                        strTelephonyType = csentry["TELEPHONE_TYPE"].Value.ToString();
                    else
                        strTelephonyType = "0";
                    string dialPlanFlag, evFlag, registrarServerFlag, voicePolicyFlag, lineURIFlag, clientPolicyFlag, ovaFlag, tnmsMBPolicyFlag, tnmsLyncFlag = "0", tnmsUMFlag;

                    if (mventry["msRTCSIP-UserEnabled"].IsPresent)
                    {
                        tnmsLyncFlag = (mventry["msRTCSIP-UserEnabled"].BooleanValue == true) ? "1" : "0";
                    }

                    if ((csentry["DIAL_PLAN"].IsPresent && existingDialPlan == csentry["DIAL_PLAN"].Value) || (string.IsNullOrEmpty(existingDialPlan) && !(isValidTNType(strTelephonyType))))
                        dialPlanFlag = " 0 ";
                    else
                        dialPlanFlag = " 1 ";

                    if (((csentry["EVFLAG_VERIFICATION"].IsPresent && !(isExternallyChanged(strTelephonyType, csentry["EVFLAG_VERIFICATION"].Value))) && (csentry["EV_FLAG"].IsPresent && existingEVFlag == csentry["EV_FLAG"].Value)) || (string.IsNullOrEmpty(existingEVFlag) && !(isValidTNType(strTelephonyType))))
                        evFlag = " 0 ";
                    else
                        evFlag = " 1 ";
                    if (csentry["REGISTRAR_SERVER"].IsPresent && existingRegistrarServer == csentry["REGISTRAR_SERVER"].Value || (string.IsNullOrEmpty(existingRegistrarServer) && !(isValidTNType(strTelephonyType))))
                        registrarServerFlag = " 0 ";
                    else
                        registrarServerFlag = " 1 ";
                    if (csentry["VOICE_POLICY"].IsPresent && existingVoicePolicy == csentry["VOICE_POLICY"].Value || (string.IsNullOrEmpty(existingVoicePolicy) && !(isValidTNType(strTelephonyType))))
                        voicePolicyFlag = " 0 ";
                    else
                        voicePolicyFlag = " 1 ";
                    //Telephone Number from TNMS for Lineuri
                    if (csentry["LINE_URI"].IsPresent && existingLineURI == csentry["LINE_URI"].Value || (string.IsNullOrEmpty(existingLineURI) && !(isValidTNType(strTelephonyType))))
                        lineURIFlag = " 0 ";
                    else
                        lineURIFlag = " 1 ";
                    //Client Policy
                    if (csentry["CLIENT_POLICY"].IsPresent && existingClientPolicy == csentry["CLIENT_POLICY"].Value || (string.IsNullOrEmpty(existingClientPolicy) && !(isValidTNType(strTelephonyType))))
                        clientPolicyFlag = " 0 ";
                    else
                        clientPolicyFlag = " 1 ";
                    //OVA Flag. This is set from Telephone Type field in TNMS. The value LEV,PCXTU or OVA will set ovaFlag to 1
                    //if (csentry["TELEPHONE_TYPE"].IsPresent && (strTelephonyType.ToUpper() != existingOVAValue.ToUpper() || isUMExternallyChanged(strTelephonyType, strmsExchUMEnabledFlags)))
                    
                    //Commenting out below if block code (TFS) as this is not latest as per PRD TNMS Data MA dll, Feb-16-2018 
                    //if (csentry["TELEPHONE_TYPE"].IsPresent && (strTelephonyType.ToUpper() != existingOVAValue.ToUpper() ))
                    //ovaFlag = " 1 ";
                    //else
                    //ovaFlag = " 0 ";
                    
                    // Below if block code decompiled with the latest TNMS Data MA dll from PRD, Feb-16-2018
                    //if (!(csentry["TELEPHONE_TYPE"].IsPresent) && (!(strTelephonyType.ToUpper() != existingOVAValue.ToUpper() || isUMExternallyChanged(strTelephonyType, strmsExchUMEnabledFlags))))
                    //ovaFlag = " 0 ";
                    //else
                    //ovaFlag = " 1 ";
                    
                    //Start-Feb-16-2016, CHG1223670 Below if block updated because it is always getting true in PRD so ovaFlag value setting 1 again and again for records those have TNMS data connector.
                    if (csentry["TELEPHONE_TYPE"].IsPresent && (isUMExternallyChanged(strTelephonyType, strmsExchUMEnabledFlags)))
                        ovaFlag = " 1 ";
                    else
                        ovaFlag = " 0 ";
                    //End-Feb-16-2016, CHG1223670

                    if (csentry["MAILBOX_POLICY"].IsPresent && existingMBPolicyValue.ToUpper() == csentry["MAILBOX_POLICY"].Value.ToUpper() || (string.IsNullOrEmpty(existingMBPolicyValue) && !(isValidTNType(strTelephonyType))))
                        tnmsMBPolicyFlag = " 0 ";
                    else
                        tnmsMBPolicyFlag = " 1 ";



                    mventry["tnmsChangedFlags"].Value = evFlag + dialPlanFlag + voicePolicyFlag + registrarServerFlag + lineURIFlag + clientPolicyFlag + ovaFlag + tnmsMBPolicyFlag;

                    break;

                default:
                    throw new EntryPointNotImplementedException();
            }
        }

        private bool isValidTNType(string TNType)
        {
            bool blnTNTypes = false;
            string[] arrTNTypes = TNTypes.Split(',');
            foreach (string strTNType in arrTNTypes)
            {
                if (strTNType.ToUpper() == TNType.ToUpper())
                {
                    blnTNTypes = true;
                    break;
                }
            }
            return blnTNTypes;
        }
        private bool isExternallyChanged(string telephonyType, string optionFlags)
        {
            bool blnExternallyChanged = false;
            //if (telephonyType == LEVType && optionFlags != "385")
            //    blnExternallyChanged = true;
            //if (telephonyType != LEVType && optionFlags == "385")
            //    blnExternallyChanged = true;
            if (LEVType.Contains(telephonyType) && optionFlags != "385")
                blnExternallyChanged = true;
            if (!LEVType.Contains(telephonyType) && optionFlags == "385")
                blnExternallyChanged = true;
            return blnExternallyChanged;
        }
        //Start-Feb-16-2016, CHG1223670. below block of code is added with msExchUMEnabledFlags value to -1 or 831 as it was there for on-prem users.
        //Intially msExchUMEnabledFlags value was 831 only, that was only intended to on-prem users. 
        private bool isUMExternallyChanged(string telephonyType, string msExchUMEnabledFlags)
        {
            bool blnExternallyChanged = false;
            if (isValidTNType(telephonyType) && msExchUMEnabledFlags != "-1")

            {
                if (msExchUMEnabledFlags != "831")
                {
                    blnExternallyChanged = true;
                }
                return blnExternallyChanged;
            }
            if ((!(isValidTNType(telephonyType))) && (msExchUMEnabledFlags == "-1" || msExchUMEnabledFlags == "831"))

            {
                blnExternallyChanged = true;
            }
            return blnExternallyChanged;
        }
        //End-Feb-16-2016, CHG1223670
        void IMASynchronization.MapAttributesForExport(string FlowRuleName, MVEntry mventry, CSEntry csentry)
        {
            throw new EntryPointNotImplementedException();
        }
    }
}
