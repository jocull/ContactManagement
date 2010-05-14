using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OleDb;

namespace ContactManagement
{
    public class clsContact
    {
        public int ContactID;
        public string ContactLastName;
        public string ContactFirstName;
        public string ContactMiddleName;
        public string ContactCompany;
        public string ContactJobTitle;
        public string ContactPhoneBusiness;
        public string ContactPhoneHome;
        public string ContactPhoneCell;
        public string ContactPhoneFax;
        public string ContactPhoneOther;
        public string ContactAddressHome;
        public string ContactAddressWork;
        public string ContactAddressVacation;
        public string ContactAddressOther;
        public string ContactEmail1;
        public string ContactEmail2;
        public string ContactEmail3;
        public string ContactNotes; //Memo field
        public string ContactDepartment;
        public string ContactManagersName;
        public string ContactAssistantsName;
        public string ContactOffice;
        public string ContactProfession;
        public string ContactNickname;
        public string ContactSpousesName;
        public string ContactTitle;
        public string ContactSuffix;

        public DateTime ContactBirthday;
        public DateTime ContactAnniversary;
        public DateTime ContactDateAdded;
        public DateTime ContactModified;

        public void createObject(OleDbDataReader rdr)
        {
            ContactID = (int)rdr["ContactID"];
            ContactLastName = (string)rdr["ContactLastName"].ToString();
            ContactFirstName = (string)rdr["ContactFirstName"].ToString();
            ContactMiddleName = (string)rdr["ContactMiddleName"].ToString();
            ContactJobTitle = (string)rdr["ContactJobTitle"].ToString();
            ContactCompany = (string)rdr["ContactCompany"].ToString();
            ContactPhoneBusiness = (string)rdr["ContactPhoneBusiness"].ToString();
            ContactPhoneHome = (string)rdr["ContactPhoneHome"].ToString();
            ContactPhoneCell = (string)rdr["ContactPhoneCell"].ToString();
            ContactPhoneFax = (string)rdr["ContactPhoneFax"].ToString();
            ContactPhoneOther = (string)rdr["ContactPhoneOther"].ToString();
            ContactAddressHome = (string)rdr["ContactAddressHome"].ToString();
            ContactAddressWork = (string)rdr["ContactAddressWork"].ToString();
            ContactAddressVacation = (string)rdr["ContactAddressVacation"].ToString();
            ContactAddressOther = (string)rdr["ContactAddressOther"].ToString();
            ContactEmail1 = (string)rdr["ContactEmail1"].ToString();
            ContactEmail2 = (string)rdr["ContactEmail2"].ToString();
            ContactEmail3 = (string)rdr["ContactEmail3"].ToString();
            ContactNotes = (string)rdr["ContactNotes"].ToString();
            ContactDepartment = (string)rdr["ContactDepartment"].ToString();
            ContactManagersName = (string)rdr["ContactManagerName"].ToString();
            ContactAssistantsName = (string)rdr["ContactAssistantName"].ToString();
            ContactOffice = (string)rdr["ContactOffice"].ToString();
            ContactProfession = (string)rdr["ContactProfession"].ToString();
            ContactNickname = (string)rdr["ContactNickname"].ToString();
            ContactSpousesName = (string)rdr["ContactSpousesName"].ToString();
            ContactTitle = (string)rdr["ContactTitle"].ToString();
            ContactSuffix = (string)rdr["ContactSuffix"].ToString();

            //DateTime variables must be in try/catch blocks?
            try
            {
                ContactBirthday = (DateTime)rdr["ContactBirthday"];
            }
            catch
            {
                //Leave value as null
            }

            try
            {
                ContactAnniversary = (DateTime)rdr["ContactAnniversary"];
            }
            catch
            {
                //Leave value as null
            }

            try
            {
                ContactDateAdded = (DateTime)rdr["ContactAdded"];
            }
            catch
            {
                //Leave value as null
            }

            try
            {
                ContactModified = (DateTime)rdr["ContactModified"];
            }
            catch
            {
                //Leave value as null
            }
        }
    }
}