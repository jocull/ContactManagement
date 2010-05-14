using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;

namespace ContactManagement
{
    public partial class UpdateContact : Form
    {

        //Global Values for Database, Mode, and Contact object
        public string mFileName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Contacts.ctm");
        private string mFormMode;

        clsContact mContact = new clsContact();

        public UpdateContact(string title, string mode, clsContact contact)
        {
            InitializeComponent();

            Text = title;

            mFormMode = mode;

            clsContact mContact = contact;

            if (mode == "Update")
            {
                //Populate the form with existing data
                populate(contact);
            }

        }

        private void UpdateContact_Load(object sender, EventArgs e)
        {

        }

        private void populate(clsContact contact)
        {
            //Build values into the form here.
            txtContactFirstName.Text = contact.ContactFirstName;
            txtContactMiddleName.Text = contact.ContactMiddleName;
            txtContactLastName.Text = contact.ContactLastName;
            txtContactJobTitle.Text = contact.ContactJobTitle;
            txtContactCompany.Text = contact.ContactCompany;

            //Email addresses
            txtContactEmail1.Text = contact.ContactEmail1;
            txtContactEmail2.Text = contact.ContactEmail2;
            txtContactEmail3.Text = contact.ContactEmail3;

            //Addresses
            txtContactAddressHome.Text = contact.ContactAddressHome;
            txtContactAddressWork.Text = contact.ContactAddressWork;
            txtContactAddressVacation.Text = contact.ContactAddressVacation;
            txtContactAddressOther.Text = contact.ContactAddressOther;

            //Phone Numbers
            txtContactPhoneWork.Text = contact.ContactPhoneBusiness;
            txtContactPhoneHome.Text = contact.ContactPhoneHome;
            txtContactPhoneCell.Text = contact.ContactPhoneCell;
            txtContactPhoneFax.Text = contact.ContactPhoneFax;
            txtContactPhoneOther.Text = contact.ContactPhoneOther;

            //Notes
            txtContactNotes.Text = contact.ContactNotes;

            //Dept, Office, Prof.
            txtContactDepartment.Text = contact.ContactDepartment;
            txtContactOffice.Text = contact.ContactOffice;
            txtContactProfession.Text = contact.ContactProfession;

            //Nickname, Title, Suffix
            txtContactNickname.Text = contact.ContactNickname;
            txtContactTitle.Text = contact.ContactTitle;
            txtContactSuffix.Text = contact.ContactSuffix;

            //Manager, Assistant, Spouse names
            txtContactManagerName.Text = contact.ContactManagersName;
            txtContactAssistantName.Text = contact.ContactAssistantsName;
            txtContactSpouseName.Text = contact.ContactSpousesName;

            //Birthday and Anniversary
            if (contact.ContactBirthday.Year < 1000)
            {
                txtContactBirthday.Text = "";
            }
            else
            {
                txtContactBirthday.Text = contact.ContactBirthday.ToShortDateString();
            }

            if (contact.ContactAnniversary.Year < 1000)
            {
                txtContactAnniversary.Text = "";
            }
            else
            {
                txtContactAnniversary.Text = contact.ContactAnniversary.ToShortDateString();
            }

            //Set Hidden ID Field to Contact ID
            lblHiddenID.Text = contact.ContactID.ToString();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (validateInput() == false)
            {
                return;
            }

            if (mFormMode == "Add")
            {
                addToDatabase(buildInsertSQL());
            }

            if (mFormMode == "Update")
            {
                addToDatabase(buildUpdateSQL());
            }

            //Close form when done saving.
            Close();
        }

        private void addToDatabase(string sql)
        {
            OleDbConnection db = null;

            try
            {
                db = new OleDbConnection();
                db.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data source=" + mFileName;
                db.Open();

                OleDbCommand command = new OleDbCommand(sql, db);
                command.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                showErrorMessage(ex.Message);
            }
            finally
            {
                if (db != null)
                {
                    db.Close();
                }
            }
        }

        private bool validateInput()
        {
            //Most data is user specific, so only validate Dates for now
                //Birthdate, anniversary

            DateTime x;

            if (txtContactBirthday.Text.Length > 0)
            {
                if (DateTime.TryParse(txtContactBirthday.Text, out x) == false)
                {
                    showErrorMessage("Birthdate must be in a valid format.");
                    txtContactBirthday.Focus();
                    return false;
                }
                if (x > DateTime.Today)
                {
                    showErrorMessage("Birthdate must be less than today's date.");
                    return false;
                }
            }

            if (txtContactAnniversary.Text.Length > 0)
            {
                if (DateTime.TryParse(txtContactAnniversary.Text, out x) == false)
                {
                    showErrorMessage("Anniversary must be in a valid format.");
                    txtContactAnniversary.Focus();
                    return false;
                }
            }

            return true;
        }

        private void showErrorMessage(string msg)
        {
            MessageBox.Show(msg, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private string buildInsertSQL()
        {
            string sql = "INSERT INTO Contacts ("
                //Columns
                + "ContactLastName, "
                + "ContactFirstName, "
                + "ContactMiddleName, "
                + "ContactCompany, "
                + "ContactJobTitle, "
                + "ContactPhoneBusiness, "
                + "ContactPhoneHome, "
                + "ContactPhoneCell, "
                + "ContactPhoneFax, "
                + "ContactPhoneOther, "
                + "ContactAddressHome, "
                + "ContactAddressWork, "
                + "ContactAddressVacation, "
                + "ContactAddressOther, "
                + "ContactEmail1, "
                + "ContactEmail2, "
                + "ContactEmail3, "
                + "ContactNotes, "
                + "ContactDepartment, "
                + "ContactManagerName, "
                + "ContactAssistantName, "
                + "ContactOffice, "
                + "ContactProfession, "
                + "ContactNickname, "
                + "ContactSpousesName, "
                + "ContactTitle, "
                + "ContactSuffix, ";
                    //Dates
                if(txtContactBirthday.Text.Length > 0)
                {
                    sql += "ContactBirthday, ";
                }
                if(txtContactAnniversary.Text.Length >0)
                {
                    sql += "ContactAnniversary, ";
                }
                sql += "ContactAdded, "
                    + "ContactModified) ";
            
                //Values
                    sql += "VALUES ("
                    + toSQL(txtContactLastName.Text) + ", "
                    + toSQL(txtContactFirstName.Text) + ", "
                    + toSQL(txtContactMiddleName.Text) + ", "
                    + toSQL(txtContactCompany.Text) + ", "
                    + toSQL(txtContactJobTitle.Text) + ", "
                    + toSQL(txtContactPhoneWork.Text) + ", "
                    + toSQL(txtContactPhoneHome.Text) + ", "
                    + toSQL(txtContactPhoneCell.Text) + ", "
                    + toSQL(txtContactPhoneFax.Text) + ", "
                    + toSQL(txtContactPhoneOther.Text) + ", "
                    + toSQL(txtContactAddressHome.Text) + ", "
                    + toSQL(txtContactAddressWork.Text) + ", "
                    + toSQL(txtContactAddressVacation.Text) + ", "
                    + toSQL(txtContactAddressOther.Text) + ", "
                    + toSQL(txtContactEmail1.Text) + ", "
                    + toSQL(txtContactEmail2.Text) + ", "
                    + toSQL(txtContactEmail3.Text) + ", "
                    + toSQL(txtContactNotes.Text) + ", "
                    + toSQL(txtContactDepartment.Text) + ", "
                    + toSQL(txtContactManagerName.Text) + ", "
                    + toSQL(txtContactAssistantName.Text) + ", "
                    + toSQL(txtContactOffice.Text) + ", "
                    + toSQL(txtContactProfession.Text) + ", "
                    + toSQL(txtContactNickname.Text) + ", "
                    + toSQL(txtContactSpouseName.Text) + ", "
                    + toSQL(txtContactTitle.Text) + ", "
                    + toSQL(txtContactSuffix.Text) + ", ";

                    //Dates
                    if (txtContactBirthday.Text.Length > 0)
                    {
                        sql += toSQL(DateTime.Parse(txtContactBirthday.Text)) + ", ";
                    }
                    if (txtContactAnniversary.Text.Length > 0)
                    {
                        sql += toSQL(DateTime.Parse(txtContactAnniversary.Text)) + ", ";
                    }

                    //Added and Moddified
                    sql += toSQL(DateTime.Now) + ", "
                    + toSQL(DateTime.Now) + ")";

                    return sql;
        }

        private string buildUpdateSQL()
        {
            string sql = "UPDATE Contacts "
                //Columns
            + "SET ContactLastName = " + toSQL(txtContactLastName.Text)
            + ", ContactFirstName = " + toSQL(txtContactFirstName.Text)
            + ", ContactMiddleName = " + toSQL(txtContactMiddleName.Text)
            + ", ContactCompany = " + toSQL(txtContactCompany.Text)
            + ", ContactJobTitle = " + toSQL(txtContactJobTitle.Text)
            + ", ContactPhoneBusiness = " + toSQL(txtContactPhoneWork.Text)
            + ", ContactPhoneHome = " + toSQL(txtContactPhoneHome.Text)
            + ", ContactPhoneCell = " + toSQL(txtContactPhoneCell.Text)
            + ", ContactPhoneFax = " + toSQL(txtContactPhoneFax.Text)
            + ", ContactPhoneOther = " + toSQL(txtContactPhoneOther.Text)
            + ", ContactAddressHome = " + toSQL(txtContactAddressHome.Text)
            + ", ContactAddressWork = " + toSQL(txtContactAddressWork.Text)
            + ", ContactAddressVacation = " + toSQL(txtContactAddressVacation.Text)
            + ", ContactAddressOther = " + toSQL(txtContactAddressOther.Text)
            + ", ContactEmail1 = " + toSQL(txtContactEmail1.Text)
            + ", ContactEmail2 = " + toSQL(txtContactEmail2.Text)
            + ", ContactEmail3 = " + toSQL(txtContactEmail3.Text)
            + ", ContactNotes = " + toSQL(txtContactNotes.Text)
            + ", ContactDepartment = " + toSQL(txtContactDepartment.Text)
            + ", ContactManagerName = " + toSQL(txtContactManagerName.Text)
            + ", ContactAssistantName = " + toSQL(txtContactAssistantName.Text)
            + ", ContactOffice = " + toSQL(txtContactOffice.Text)
            + ", ContactProfession = " + toSQL(txtContactProfession.Text)
            + ", ContactNickname = " + toSQL(txtContactNickname.Text)
            + ", ContactSpousesName = " + toSQL(txtContactSpouseName.Text)
            + ", ContactTitle = " + toSQL(txtContactTitle.Text)
            + ", ContactSuffix = " + toSQL(txtContactSuffix.Text);
            //Dates
            if (txtContactBirthday.Text.Length > 0)
            {
                sql += ", ContactBirthday = " + toSQL(DateTime.Parse(txtContactBirthday.Text));
            }
            else
            {
                sql += ", ContactBirthday = NULL";
            }
            if (txtContactAnniversary.Text.Length > 0)
            {
                sql += ", ContactAnniversary = " + toSQL(DateTime.Parse(txtContactAnniversary.Text));
            }
            else
            {
                sql += ", ContactAnniversary = NULL";
            }
            sql += ", ContactModified = " + toSQL(DateTime.Now)
                + " WHERE ContactID=" + lblHiddenID.Text;

            return sql;
        }

        private string toSQL(string str)
        {
            return "'" + str.Replace("'", "''") + "'";
        }

        private string toSQL(DateTime date)
        {
            return "#" + date.ToString() + "#";
        }
    }
}