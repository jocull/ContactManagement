using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ContactManagement
{
    public partial class ViewContact : Form
    {
        public ViewContact(string title, clsContact contact)
        {
            InitializeComponent();

            //Set caption to the person's name
            Text = title;

            buildForm(contact);
        }

        private string mContactEmail1;
        private string mContactEmail2;
        private string mContactEmail3;

        private void buildForm(clsContact contact)
        {
            //Build values into the form here.
            lblContactName.Text = contact.ContactFirstName + " " + contact.ContactMiddleName + " " + contact.ContactLastName;
            lblContactJobTitle.Text = contact.ContactJobTitle;
            lblContactCompany.Text = contact.ContactCompany;

                //Email addresses
            lblContactEmail1.Text = contact.ContactEmail1;
            mContactEmail1 = contact.ContactEmail1;
            lblContactEmail2.Text = contact.ContactEmail2;
            mContactEmail2 = contact.ContactEmail2;
            lblContactEmail3.Text = contact.ContactEmail3;
            mContactEmail3 = contact.ContactEmail3;

                //Addresses
            lblContactAddressHome.Text = contact.ContactAddressHome;
            lblContactAddressWork.Text = contact.ContactAddressWork;
            lblContactAddressVacation.Text = contact.ContactAddressVacation;
            lblContactAddressOther.Text = contact.ContactAddressOther;

                //Phone Numbers
            lblContactPhoneBusiness.Text = contact.ContactPhoneBusiness;
            lblContactPhoneHome.Text = contact.ContactPhoneHome;
            lblContactPhoneCell.Text = contact.ContactPhoneCell;
            lblContactPhoneFax.Text = contact.ContactPhoneFax;
            lblContactPhoneOther.Text = contact.ContactPhoneOther;

                //Notes
            txtContactNotes.Text = contact.ContactNotes;

                //Dept, Office, Prof.
            lblContactDepartment.Text = contact.ContactDepartment;
            lblContactOffice.Text = contact.ContactOffice;
            lblContactProfession.Text = contact.ContactProfession;

                //Nickname, Title, Suffix
            lblContactNickname.Text = contact.ContactNickname;
            lblContactTitle.Text = contact.ContactTitle;
            lblContactSuffix.Text = contact.ContactSuffix;

                //Manager, Assistant, Spouse names
            lblContactManagerName.Text = contact.ContactManagersName;
            lblContactAssistantName.Text = contact.ContactAssistantsName;
            lblContactSpouseName.Text = contact.ContactSpousesName;

                //Birthday, Anniversary
            if (contact.ContactBirthday.Year < 1000)
            {
                lblContactBirthday.Text = "";
            }
            else
            {
                lblContactBirthday.Text = contact.ContactBirthday.ToShortDateString();
            }

            if (contact.ContactAnniversary.Year < 1000)
            {
                lblContactAnniversary.Text = "";
            }
            else
            {
                lblContactAnniversary.Text = contact.ContactAnniversary.ToShortDateString();
            }

                //Added, Modified
            lblDateAdded.Text = contact.ContactDateAdded.ToShortDateString();
            lblDateModified.Text = contact.ContactModified.ToString();
        }

        private void lblContactEmail1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string connectionString = "mailto:" + mContactEmail1;
            System.Diagnostics.Process.Start("mailto:" + mContactEmail1);
        }

        private void lblContactEmail2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string connectionString = "mailto:" + mContactEmail2;
            System.Diagnostics.Process.Start("mailto:" + mContactEmail2);
        }

        private void lblContactEmail3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string connectionString = "mailto:" + mContactEmail3;
            System.Diagnostics.Process.Start("mailto:" + mContactEmail3);
        }
    }
}