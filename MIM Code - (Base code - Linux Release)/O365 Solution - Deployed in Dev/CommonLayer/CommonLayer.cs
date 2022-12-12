using System;
using System.Collections.Generic;
using System.Text;
using System.Collections;
using System.IO;
using System.Xml;
using System.Diagnostics;
using System.DirectoryServices;
using Microsoft.MetadirectoryServices;
using System.Text.RegularExpressions;

namespace CommonLayer_NameSpace
{
    public static class CommonLayer
    {

        //Used for list of secure email domains
        static string strSecureEmailDomain, input, strWTPGlobal, strRFDCPath,strPACWhereToprovision;
        static Hashtable afvalhash;

        //Secure email array and valuecollection
        static string[] arrSecureDomain;
        static ValueCollection vcSecureDomainWildCard = Utils.ValueCollection("initialvalue");
        static ValueCollection vcSecureDomainWithoutWildCard = Utils.ValueCollection("initialvalue");

        static StreamReader objStreamReader;
        static XmlNode rnode;
        static XmlNode node;
        const string XML_CONFIG_FILE = @"\rules-config.xml";
        static string stremployeeType = string.Empty;

        //Constant values strings
        public const string USER_MAILBOX = "1";
        public const string LINKED_MAILBOX = "2";
        public const string MAIL_USER = "128";
        public const string MAIL_CONTACT = "32768";
        public const string REMOTE_MAILBOX = "2147483648";

        public const string GUT_MAIL_CONTACT = "MailContact";
        public const string GUT_MADS = "MADS";
        public const string GUT_LOCAL_MBX = "LocalMailbox";
        public const string GUT_REMOTE_MBX = "RemoteMailbox";
        public const string GUT_None = "None";

        static string strRFServer = string.Empty; 

        static CommonLayer()
        {
            objStreamReader = File.OpenText("C:\\Program Files\\Microsoft Forefront Identity Manager\\2010\\Synchronization Service\\Extensions\\SecureDomains.txt");

            XmlDocument config = new XmlDocument();
            string dir = Utils.ExtensionsDirectory;
            config.Load(dir + XML_CONFIG_FILE);

            rnode = config.SelectSingleNode("rules-extension-properties");
            node = rnode.SelectSingleNode("environment");
            string env = node.InnerText;
            rnode = config.SelectSingleNode
                ("rules-extension-properties/management-agents/" + env + "/ad-ma");
            XmlNode confignode = rnode.SelectSingleNode("siteconfigfile");
            afvalhash = LoadAFConfig(confignode.InnerText);

            //NLW mini Release
            node = rnode.SelectSingleNode("emplyeecode");
            stremployeeType = node.InnerText;

            //Get RF DC path for mailnickname search
            node = rnode.SelectSingleNode("RFDCLDAPString");
            strRFDCPath = node.InnerText;

            //Get RF server for LDAP query search
            node = rnode.SelectSingleNode("RFDC");
            strRFServer = node.InnerText;

            rnode = config.SelectSingleNode
                ("rules-extension-properties/management-agents/" + env + "/Exchange");
            node = rnode.SelectSingleNode("WhereToProvision");
            strWTPGlobal = node.InnerText;


            while ((input = objStreamReader.ReadLine()) != null)
            {
                strSecureEmailDomain = strSecureEmailDomain + input;
            }

            vcSecureDomainWildCard.Remove("initialvalue");
            vcSecureDomainWithoutWildCard.Remove("initialvalue");

            if ((strSecureEmailDomain == "") || (strSecureEmailDomain == null))
            {
                throw new Exception("Secure Email Domain is empty");
            }
            else
            {
                arrSecureDomain = strSecureEmailDomain.Split(';');
            }
            //Add secure domains to the value collection
            for (int i = 0; i < arrSecureDomain.Length; i++)
            {
                if (arrSecureDomain[i].ToString().TrimStart().StartsWith("."))
                {
                    vcSecureDomainWildCard.Add(arrSecureDomain[i].ToString());
                }
                else
                {
                    vcSecureDomainWithoutWildCard.Add(arrSecureDomain[i].ToString());
                }
            }
        }
      
       

        /// <summary>
        /// This is the common function to determine the SA Flag value 
        /// </summary>
        /// <param name="htUserDetails">Hashtable with input details to determine the user type</param>
        /// <returns>returns user type</returns>
        public static bool ISSysAccessFlagSet(MVEntry mventry)
        {
            if (mventry["System_Access_Flag"].IsPresent && mventry["System_Access_Flag"].Value.ToUpper() == "N")
            {
                string[] arrstrEmplyeeTypes = stremployeeType.Split(',');
                for (int i = 0; i < arrstrEmplyeeTypes.Length; i++)
                {
                    if (mventry["employeeType"].IsPresent)
                    {
                        if (Convert.ToString(mventry["employeeType"].Value).ToUpper() == arrstrEmplyeeTypes[i])
                        {

                            return false;
                        }
                    }
                }
            }
            return true;
        }

        /// <summary>
        /// This is the common function to determine the user type in the Exchange migration
        /// </summary>
        /// <param name="htUserDetails">Hashtable with input details to determine the user type</param>
        /// <returns>returns user type (RemoteMailbox,LocalMailbox,MailEnabledUser,MailContact,SADAuthoritative,None)</returns>
        public static string GetUserType(MVEntry mvObject, CSEntry csObject, string CallingMA, string skipMail)
        {
            #region variable declaration
            string strType;
            string strExternal_Email_Address = string.Empty;
            string strExt_Auth_Flag = string.Empty;
            string strInternet_Style_Adrs = string.Empty;
            string strUC_Flag = string.Empty;
            string strmsHomeMTA = string.Empty;
            string strWhereToProvision = string.Empty;
            string strmsRecTypeDetails = string.Empty;
            string strWTP = string.Empty;
            #endregion

            #region Set the values
            if (mvObject["homeMTA"].IsPresent)
                strmsHomeMTA = mvObject["homeMTA"].Value.ToString();
            if (mvObject["msExchRecipientTypeDetails"].IsPresent)
                strmsRecTypeDetails = mvObject["msExchRecipientTypeDetails"].Value.ToString();

            //Get where to provision information
            strWhereToProvision = GetWhereToProvisionMailbox(mvObject, csObject, strWTPGlobal, CallingMA);

            if (CallingMA == "SAD")
            {
                if ((csObject["BUSINESS_EMAIL_TXT"].IsPresent))
                    strExternal_Email_Address = csObject["BUSINESS_EMAIL_TXT"].Value.ToString();
                else if (mvObject["IsBusinessEmailFromSAD"].IsPresent)
                    strExternal_Email_Address = string.Empty;
                else if (mvObject["External_Email_Address"].IsPresent)
                    strExternal_Email_Address = mvObject["External_Email_Address"].Value.ToString();
                if ((csObject["INTERNET_STYLE_ADRS"].IsPresent))
                    strInternet_Style_Adrs = csObject["INTERNET_STYLE_ADRS"].Value.ToString();
                //BOC- Novartis Release
                //Override logic for MADS project
                if (!csObject["NOVARTIS_FLG"].IsPresent || (csObject["NOVARTIS_FLG"].IsPresent && csObject["NOVARTIS_FLG"].Value.ToString() == "N"))
                {
                    if (IsMigrationOverride(mvObject, csObject, CallingMA))
                    {
                        strUC_Flag = "N";
                        strExt_Auth_Flag = "Y";
                    }
                    else
                    {
                        strUC_Flag = csObject["UNIFIED_COMM_FLG"].Value.ToUpper();
                        strExt_Auth_Flag = csObject["EXT_AUTH_FLG"].Value.ToUpper();
                    }
                }
                else
                {
                    if (IsMigrationAcquisition(mvObject, csObject, CallingMA))
                    {
                        strUC_Flag = "Y";
                        strExt_Auth_Flag = "N";
                    }
                    else
                    {
                        strUC_Flag = csObject["UNIFIED_COMM_FLG"].Value.ToUpper();
                        strExt_Auth_Flag = csObject["EXT_AUTH_FLG"].Value.ToUpper();
                    }

                }
                ////EOC- Novartis Release

            }
            else
            {
                if ((mvObject["External_Email_Address"].IsPresent))
                {
                    strExternal_Email_Address = mvObject["External_Email_Address"].Value.ToString();
                }
                if ((mvObject["Internet_Style_Adrs"].IsPresent))
                {
                    strInternet_Style_Adrs = mvObject["Internet_Style_Adrs"].Value.ToString();
                }
                strUC_Flag = mvObject["UC_Flag"].Value.ToUpper();
                strExt_Auth_Flag = mvObject["Ext_Auth_Flag"].Value.ToUpper();
            }
            #endregion

            try
            {
                //Logic to determine user type
                
                #region UC Flag = Y check
                if (strUC_Flag == "Y")
                {
                    if (strmsRecTypeDetails == REMOTE_MAILBOX)
                    {//for existing remote mailbox
                        strType = CommonLayer.GUT_REMOTE_MBX;
                    }
                    else if (strWhereToProvision == "Remote")
                    {//for new user remote mailbox
                        strType = CommonLayer.GUT_REMOTE_MBX;
                    }
                    else if (strmsRecTypeDetails == LINKED_MAILBOX)
                    {//for existing local mailbox
                        strType = CommonLayer.GUT_LOCAL_MBX;
                    }
                    else if (strWhereToProvision == "Local")
                    {//for new local mailbox
                        strType = CommonLayer.GUT_LOCAL_MBX;
                    }
                    else
                    {
                        strType = CommonLayer.GUT_None;
                    }
                }
                #endregion
                #region Business Email check
                else if (strExternal_Email_Address != string.Empty)
                {
                    if (IsDomainSecure(strExternal_Email_Address) == true)
                    {//Secure domain
                        strType = CommonLayer.GUT_MAIL_CONTACT;
                    }
                    else
                    {
                        strType = CommonLayer.GUT_None;
                    }
                }
                # endregion
                #region Ext. Auth. check
                else if (strExt_Auth_Flag == "Y")
                {
                    if (strInternet_Style_Adrs != string.Empty)
                    {
                        strType = CommonLayer.GUT_MADS;
                    }
                    else
                    {
                        strType = CommonLayer.GUT_None;
                    }
                }
                #endregion
                #region None
                else
                {//None of the above
                    strType = CommonLayer.GUT_None;
                }
                #endregion None
                return strType;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// This method is used to check if the email domain is secure (from CIT source data)
        /// </summary>
        /// <param name="BusinessEmail"></param>
        /// <returns></returns>
        public static bool IsDomainSecure(string BusinessEmail)
        {
            string[] EmailDomainstr;
            bool secureEmailFlag = false;
            EmailDomainstr = BusinessEmail.ToLower().Split('@');

            //Secure email check
            //Wildcard check
            foreach (Value addrElement in vcSecureDomainWildCard)
            {
                string strWildCardDomain = "." + EmailDomainstr[1].ToString().ToLower();
                if (strWildCardDomain.EndsWith(addrElement.ToString().ToLower().TrimEnd()))
                {
                    secureEmailFlag = true;
                    break;
                }
            }
            //Without wildcard - exact match
            foreach (Value addrElement in vcSecureDomainWithoutWildCard)
            {
                if (EmailDomainstr[1].ToLower() == addrElement.ToString().ToLower().TrimStart().TrimEnd())
                {
                    secureEmailFlag = true;
                    break;
                }
            }
            return secureEmailFlag;
        }

        /// <summary>
        /// This method is used to check if the UC flag and EA flag will be overridden 
        /// </summary>
        /// <param name="mvObject"></param>
        /// <param name="csObject"></param>
        /// <param name="CallingMA"></param>
        /// <returns></returns>
        public static Boolean IsMigrationOverride(MVEntry mvObject, CSEntry csObject, string CallingMA)
        {

            if (
                    (
                        CallingMA != "SAD"
                            ||
                        ((csObject["UNIFIED_COMM_FLG"].Value.ToUpper() == "Y") && (csObject["EXT_AUTH_FLG"].Value.ToUpper() == "N"))
                    )
                    &&
                    mvObject["msExchRecipientTypeDetails"].IsPresent
                    &&
                    (mvObject["msExchRecipientTypeDetails"].Value.ToString() == MAIL_USER || mvObject["msExchRecipientTypeDetails"].Value.ToString() == MAIL_CONTACT)
                    &&
                    mvObject["msExchMailboxGuid"].IsPresent
               )
            {
                return true;
            }
            else
            {
                return false;
            }

        }


        /// <summary>
        /// This method is used to check if the migration of email domain is done
        /// </summary>
        /// <param name="mvObject"></param>
        /// <param name="csObject"></param>
        /// <param name="CallingMA"></param>
        /// <returns></returns>
        public static Boolean IsMigrationAcquisition(MVEntry mvObject, CSEntry csObject, string CallingMA)
        {
            bool boolIsMigrationAcquisition = false;
            if (csObject["NOVARTIS_FLG"].IsPresent && csObject["NOVARTIS_FLG"].Value.ToString() == "Y")
            {
                if (mvObject["msExchMailboxGuid"].IsPresent)
                {
                    if (mvObject["msExchRecipientTypeDetails"].IsPresent)
                    {
                        if (mvObject["msExchRecipientTypeDetails"].Value.ToString() != LINKED_MAILBOX)
                            boolIsMigrationAcquisition = true;
                        else
                            boolIsMigrationAcquisition = true;
                    }
                    else
                    {
                        boolIsMigrationAcquisition = true;
                    }
                }
                else
                {
                    boolIsMigrationAcquisition = false;
                }
            }
            if (mvObject["IsAcquisitionMigrate"].IsPresent && mvObject["IsAcquisitionMigrate"].BooleanValue == true)
            {
                boolIsMigrationAcquisition = true;
            }

            return boolIsMigrationAcquisition;
        }

        /// <summary>
        /// This method is used to check if mailnickname exists in RF
        /// </summary> 
        /// <param name="mvObject"></param>
        /// <param name="csObject"></param>
        /// <returns></returns>
        public static string GetUniqueMailNickname(string strdefaultmNickname, string strSamAccountName, MVEntry mvObject)
        {
            string strFinalMailnickName = strdefaultmNickname;

            // #CHG1324643 - Modifying ldap filter for UPN uniqueness fix            
            string strLDAPFilterForMailNickname = "(&(|(mailNickName=" + strdefaultmNickname + "*)(userPrincipalName=" + strdefaultmNickname + "*))(!samAccountName=" + strSamAccountName + "))";
            
            DirectoryEntry mailNNameDirectoryEntry = new DirectoryEntry("LDAP://" + strRFServer + ":389/" + strRFDCPath, null, null, AuthenticationTypes.Secure);
            DirectorySearcher dSearcher = new DirectorySearcher(mailNNameDirectoryEntry);

            //filter just user objects
            dSearcher.Filter = strLDAPFilterForMailNickname;
            dSearcher.PageSize = 1000;
            dSearcher.SearchScope = SearchScope.Subtree;
            dSearcher.Sort.Direction = System.DirectoryServices.SortDirection.Ascending;
            dSearcher.Sort.PropertyName = "mailNickName";
            bool IsUniqueMailnickname = false;
            string strUniqueMailnickname = strdefaultmNickname;
            string[] srchUpn = {};

            try
            {
                // Get collection of all users (other then strSamAccountName) with strdefaultmNickname in mail, proxy, mailnickname or upn
                SearchResultCollection resultCollection = dSearcher.FindAll();

                if (resultCollection.Count > 0)
                {
                    //search new mailnickname for all users found by incrementing by the search count
                    foreach (SearchResult userResults in resultCollection)
                    {

                        if (userResults.Properties["userPrincipalName"].Count > 0)
                        {
                            srchUpn = userResults.Properties["userPrincipalName"][0].ToString().Split('@');
                        }

                        strFinalMailnickName = strdefaultmNickname;
                        for (int i = 1; i <= resultCollection.Count; i++)
                        {
                            //increment the mailnickname since the strdefaultmNickname is already present based on the search results. search the new mailnickname in current userresult properties.
                            strFinalMailnickName = strdefaultmNickname + i.ToString();

                            // #CHG1324643 - Updaing code to search for unique mailnickname among existing mailnicknames and UPNs.
                            //If new mailnickname is not found in mailnicknames and upn then assign it to final mailnickname
                            if (((userResults.Properties["mailNickName"].Count > 0) && (strFinalMailnickName.ToLower() != userResults.Properties["mailNickName"][0].ToString().ToLower())) &&
                              ((userResults.Properties["userPrincipalName"].Count > 0) && (strFinalMailnickName.ToLower() != srchUpn[0].ToString().ToLower())))
                            {
                                IsUniqueMailnickname = true;
                                strUniqueMailnickname = strFinalMailnickName;
                            }
                        }
                    }
                    if (IsUniqueMailnickname == true)
					{
						//HCM - Adding additional check in MV for multiple users being onboarded with same names in single job run
						strUniqueMailnickname = GetCheckedMailNickName(strUniqueMailnickname, mvObject);
                        return strUniqueMailnickname;
					}
                    else
					{
						strdefaultmNickname = GetCheckedMailNickName(strdefaultmNickname, mvObject);
                        return strdefaultmNickname;
					}
                }
                else
                {
					strFinalMailnickName = GetCheckedMailNickName(strFinalMailnickName, mvObject);
                    return strFinalMailnickName;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

		
        /// <summary>
        /// This method is used to build mailnickname from angliziced names
        /// </summary>
        /// <param name="mvObject"></param>
        /// <returns></returns>
        public static string BuildMailNickName(MVEntry mvObject)
        {
            string pattern = "[^a-z0-9_\\-.]";
            Regex rgx = new Regex(pattern, RegexOptions.IgnoreCase);

            string defaultmNickname = string.Empty;
            if (mvObject["Anglczd_last_nm"].IsPresent)
                defaultmNickname = mvObject["Anglczd_last_nm"].Value.ToString();
            if (mvObject["Anglczd_first_nm"].IsPresent)
                defaultmNickname = defaultmNickname.Trim() + "_" + mvObject["Anglczd_first_nm"].Value.ToString();
            if (mvObject["Anglczd_middle_nm"].IsPresent)
                defaultmNickname = defaultmNickname.Trim() + "_" + mvObject["Anglczd_middle_nm"].Value.ToString();
            //"Remove "_" if last nm is null
            if (defaultmNickname.StartsWith("_"))
                defaultmNickname = defaultmNickname.Remove(0, 1);
            //Replace white spaces with "_"
            if (defaultmNickname.Contains(" "))
                defaultmNickname = defaultmNickname.Replace(" ", "_");

            defaultmNickname = rgx.Replace(defaultmNickname,"");

            return defaultmNickname;

        }
		
		/// <summary>
        /// HCM - This method is used to build mailnickname from angliziced names in SAD
        /// </summary>
        /// <param name="mvObject"></param>
        /// <returns></returns>
        public static string BuildMailNickName(CSEntry csObject)
        {
            string pattern = "[^a-z0-9_\\-.]";
            Regex rgx = new Regex(pattern, RegexOptions.IgnoreCase);

            string defaultmNickname = string.Empty;
            if (csObject["ANGLCZD_LAST_NM"].IsPresent)
                defaultmNickname = csObject["ANGLCZD_LAST_NM"].Value.ToString();
            if (csObject["ANGLCZD_FIRST_NM"].IsPresent)
                defaultmNickname = defaultmNickname.Trim() + "_" + csObject["ANGLCZD_FIRST_NM"].Value.ToString();
            if (csObject["ANGLCZD_MDL_NM"].IsPresent)
                defaultmNickname = defaultmNickname.Trim() + "_" + csObject["ANGLCZD_MDL_NM"].Value.ToString();
            //"Remove "_" if last nm is null
            if (defaultmNickname.StartsWith("_"))
                defaultmNickname = defaultmNickname.Remove(0, 1);
            //Replace white spaces with "_"
            if (defaultmNickname.Contains(" "))
                defaultmNickname = defaultmNickname.Replace(" ", "_");

            defaultmNickname = rgx.Replace(defaultmNickname,"");

            return defaultmNickname;

        }
		
		// This function creates a unique mailNickname for use in a metaverse entry.
		public static string GetCheckedMailNickName(string mailNickname, MVEntry mventry)
		{
			MVEntry[] findResultList = null;
			string checkedMailNickname = mailNickname;

			// Create a unique naming attribute by adding a number to
			// the existing mailNickname value.
			for (int nameSuffix = 1; nameSuffix < 100; nameSuffix++)
			{
				// Check if the mailNickname value exists in the metaverse by 
				// using the Utils.FindMVEntries method.
				findResultList = Utils.FindMVEntries("mailNickname", checkedMailNickname, 1);
				if (findResultList.Length == 0)
				{
					// The current mailNickname is not in use.
					return(checkedMailNickname);
				}

				//Check if a metaverse entry was found with the specified mailNickname, 
				MVEntry mvEntryFound = findResultList[0];
				if (mvEntryFound.Equals(mventry))
				{
					return(checkedMailNickname);
				}

				// If the passed nickname is already in use by another metaverse 
				// entry, add the counter number to the passed value and 
				// verify this new value exists. Repeat this step until a unique 
				// value is created.
				checkedMailNickname = mailNickname + nameSuffix.ToString();
			}

			// Return an empty string if no unique nickname could be created.
			return mailNickname;
		}

        /// <summary>
        /// This method is used to find the location for Exchange mailbox provisioning. It can have two values- Local and Remote
        /// </summary>
        /// <param name="mvObject"></param>
        /// <param name="csObject"></param>
        /// <returns>string</returns>
        public static string GetWhereToProvisionMailbox(MVEntry mvobject, CSEntry csobject, string strGlobalPAC, string strCallingMA)
        {
            if (strGlobalPAC == "PAC")
            {
                if (mvobject["msExchRecipientTypeDetails"].IsPresent && mvobject["msExchRecipientTypeDetails"].Value.ToString() == CommonLayer.REMOTE_MAILBOX)
                {
                    return "Remote";
                }
                // O365 bug fix - to cover transition scenario from mail contact with PAC 1000 to a mail box User - When a Mail User (128) type transitions to a Mailbox the Remote PAC should return WTP as Remote
                else if(mvobject["msExchRecipientTypeDetails"].IsPresent && ((mvobject["msExchRecipientTypeDetails"].Value.ToString() == CommonLayer.MAIL_CONTACT) || (mvobject["msExchRecipientTypeDetails"].Value.ToString() == CommonLayer.MAIL_USER)))
                {
                    setUserVariables(mvobject["personnel_area_cd"].Value, mvobject["employeeType"].Value);
                    return strPACWhereToprovision;
                }
                else if (!mvobject["msExchRecipientTypeDetails"].IsPresent && strCallingMA == "SAD")
                {
                    setUserVariables(csobject["PRSNL_AREA_CD"].Value, csobject["EMPLY_GRP_CD"].Value);
                    return strPACWhereToprovision;
                }
                else
                {
                    return "Local";
                }
            }
            else
                return "Remote";
        }
        

        public static Hashtable LoadAFConfig(string filename)
        {
            //
            // Initializing AD config file, any new PArent Tag in the config means code change / addition to this section 
            //
            XmlTextReader xmlReader = new XmlTextReader("C:\\Program Files\\Microsoft Identity Integration Server\\Extensions\\" + filename);
            // Read the line of the xml file
            Hashtable primaryhash = new Hashtable();
            while (xmlReader.ReadToFollowing("PARENT"))
            {
                Hashtable valhash = new Hashtable();
                if (xmlReader.HasAttributes)
                {
                    for (int i = 0; i < xmlReader.AttributeCount; i++)
                    {
                        xmlReader.MoveToAttribute(i);
                        primaryhash.Add(xmlReader.Value, valhash);
                    }
                }
                if (xmlReader.ReadToFollowing("PERSONNELAREA"))
                {
                    ArrayList parray = new ArrayList();
                    XmlReader tempxmlReader = (XmlReader)xmlReader.ReadSubtree();
                    while (tempxmlReader.ReadToFollowing("CODE"))
                    {
                        tempxmlReader.Read();
                        string pcode = tempxmlReader.Value;
                        parray.Add(pcode);
                    }
                    valhash.Add("PERSONNELAREA", parray);
                }
                if (xmlReader.ReadToFollowing("CONSTITUENT"))
                {
                    ArrayList etypearray = new ArrayList();
                    XmlReader tempxmlReader = (XmlReader)xmlReader.ReadSubtree();
                    while (tempxmlReader.ReadToFollowing("EMPLY_GRP"))
                    {
                        tempxmlReader.Read();
                        string etype = tempxmlReader.Value;
                        etypearray.Add(etype.ToUpper());
                    }
                    valhash.Add("CONSTITUENT", etypearray);
                }
                if (xmlReader.ReadToFollowing("WhereToProvision"))//
                {
                    xmlReader.Read();
                    string WhereToProv = xmlReader.Value;
                    valhash.Add("WhereToProvision", WhereToProv);
                }
            }
            return primaryhash;
        }

        public static void setUserVariables(string paccode, string employeeType)
        {
            //
            // Initializing the variable necessary to build a AD account from the already initialized AD config file in a Hashtable
            //

            foreach (DictionaryEntry de in afvalhash)
            {
                Hashtable afconfig = (Hashtable)de.Value;
                ArrayList listarea = (ArrayList)afconfig["PERSONNELAREA"];
                ArrayList listtype = (ArrayList)afconfig["CONSTITUENT"];
                if (listarea.Contains(paccode) && listtype.Contains(employeeType.ToUpper()))
                {
                    strPACWhereToprovision = afconfig["WhereToProvision"].ToString();
                    break;
                }
            }
        }
    }
}
