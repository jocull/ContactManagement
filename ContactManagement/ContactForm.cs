using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Collections;
using System.IO;

namespace ContactManagement
{
    public partial class ContactForm : Form
    {
        //Global Values for Database and ArrayList
        public string mFileName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Contacts.ctm");
        public ArrayList mContacts = new ArrayList();
        public string mOwnerName;
        public byte mWarningDays;

        //Sort order variable
        private bool mSortOrder = true;

        public ContactForm()
        {
            InitializeComponent();
        }

        private void ContactForm_Load(object sender, EventArgs e)
        {
            //Check for database
            if (checkDatabase() == false)
            {
                Close();
            }

            try
            {
                //Read the user's settings from options
                readSettings();

                //Read the contacts in the database and show them
                readUsers();
                showUsers();

                //Apply settings and give alerts
                goStartUp();
            }
            catch (Exception ex)
            {
                //If you get an error reading the Contacts, then there is a
                //database error. Exit the program.
                showErrorMessage(ex.Message);
                Close();
            }
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private bool checkDatabase()
        {
            if (File.Exists(mFileName) == false)
            {
                showErrorMessage("\t\tDatabase could not be found!\nPlease reload the database into the application directory and restart the program.");
                return false;
            }
            return true;
        }

        public void readSettings()
        {
            OleDbConnection db = null;
            OleDbDataReader dataReader = null;

            try
            {
                // Open a connection to the DBMS.
                db = new OleDbConnection();
                db.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data source=" + mFileName;
                db.Open();

                string sql = "SELECT * FROM Settings";

                OleDbCommand command = new OleDbCommand(sql, db);
                dataReader = command.ExecuteReader();

                while (dataReader.Read() == true)
                {
                    mOwnerName = (string)dataReader["SettingsOwnerName"].ToString();
                    mWarningDays = (byte)dataReader["SettingsBirthdayWarning"];
                }
            }
            catch (Exception ex)
            {
                showErrorMessage(ex.Message);
                Close(); //Close if you get an error here, because there is a database issue.
            }
            finally
            {
                if (dataReader != null)
                {
                    dataReader.Close();
                }

                if (db != null)
                {
                    db.Close();
                }
            }
        }

        public void goStartUp()
        {
            //Set title of the form
            if (mOwnerName.Length == 0)
            {
                this.Text = "Contact Management";
            }
            else
            {
                this.Text = "Contact Management for " + mOwnerName;
            }

            //Show an alert about any upcoming birthdays / anniversarys
            string birthdays = "";
            string anniversarys = "";
            string finalAlert = "";

            double warningDays = double.Parse(mWarningDays.ToString());
            
            //Calculate Julian day for today's date
            double todayNumber = calculateJulianDay(DateTime.Today);

            foreach(clsContact contact in mContacts)
            {
                //Add years to get this birthday up to this year
                DateTime birthdayThisYear;
                if (contact.ContactBirthday.DayOfYear < DateTime.Now.DayOfYear)
                {
                    birthdayThisYear = contact.ContactBirthday.AddYears(DateTime.Now.Year - contact.ContactBirthday.Year + 1);
                }
                else
                {
                    birthdayThisYear = contact.ContactBirthday.AddYears(DateTime.Now.Year - contact.ContactBirthday.Year);
                }

                //Calculate Julian day for birthday of this year
                double birthdayNumber = calculateJulianDay(birthdayThisYear);

                if (birthdayNumber >= todayNumber && birthdayNumber <= (todayNumber + warningDays) && contact.ContactBirthday > System.DateTime.MinValue)
                {
                    int age = DateTime.Today.AddDays(birthdayNumber - todayNumber).Year - contact.ContactBirthday.Year;

                    birthdays += contact.ContactFirstName + " " + contact.ContactLastName
                        + " - " + birthdayThisYear.ToLongDateString()
                        + " (" + contact.ContactFirstName + " will be " + age + ")"
                        + "\n";
                }

                //Add years to get this anniversary up to this year
                DateTime anniveraryThisYear;
                if (contact.ContactAnniversary.DayOfYear < DateTime.Now.DayOfYear)
                {
                    anniveraryThisYear = contact.ContactAnniversary.AddYears(DateTime.Now.Year - contact.ContactAnniversary.Year + 1);
                }
                else
                {
                    anniveraryThisYear = contact.ContactAnniversary.AddYears(DateTime.Now.Year - contact.ContactAnniversary.Year);
                }

                //Calculate Julian day for anniversary of this year
                double anniversaryNumber = calculateJulianDay(anniveraryThisYear);

                //Same thing with anniversaries
                if (anniversaryNumber >= todayNumber && anniversaryNumber <= (todayNumber + warningDays) && contact.ContactAnniversary > System.DateTime.MinValue)
                {
                    int anniversaryYears = DateTime.Today.AddDays(anniversaryNumber - todayNumber).Year - contact.ContactAnniversary.Year;

                    anniversarys += contact.ContactFirstName + " " + contact.ContactLastName
                        + " - " + contact.ContactAnniversary.ToLongDateString()
                        + " (" + contact.ContactFirstName + " and " + contact.ContactSpousesName + " have been together " + anniversaryYears + " years)"
                        + "\n";
                }
            }

            //Give the user a breakdown of all upcoming birthdays / anniversarys
            if (birthdays != "")
            {
                finalAlert += "Upcoming Birthdays:\n" + birthdays + "\n";
            }
            if (anniversarys != "")
            {
                finalAlert += "Upcoming Anniversaries:\n" + anniversarys + "\n";
            }

            if (finalAlert != "")
            {
                MessageBox.Show(finalAlert, "Upcoming Dates", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void readUsers()
        {
            OleDbConnection db = null;
            OleDbDataReader dataReader = null;

            try
            {
                // Open a connection to the DBMS.
                db = new OleDbConnection();
                db.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data source=" + mFileName;
                db.Open();

                string sql = "SELECT * FROM Contacts ORDER BY ContactLastName ASC";

                OleDbCommand command = new OleDbCommand(sql, db);
                dataReader = command.ExecuteReader();

                // Format each message returned from the query.
                mContacts = new ArrayList();
                while (dataReader.Read() == true)
                {
                    //Build a class object and add it to the array
                    //Make a new clsPost object each time.
                    clsContact contact = new clsContact();
                    contact.createObject(dataReader); //Pass the reader to create the object.

                    mContacts.Add(contact); //Add the object to the ArrayList
                }
            }
            catch (Exception ex)
            {
                showErrorMessage(ex.Message);
                Close(); //Close if you get an error here, because there is a database issue.
            }
            finally
            {
                if (dataReader != null)
                {
                    dataReader.Close();
                }

                if (db != null)
                {
                    db.Close();
                }
            }
        }

        private void showUsers()
        {
            //show users to list view
            lvwContacts.Items.Clear();

            foreach (clsContact contact in mContacts)
            {
                //Add to the ListView
                ListViewItem item;
                item = new ListViewItem(contact.ContactLastName);
                item.SubItems.Add(contact.ContactFirstName);
                item.SubItems.Add(contact.ContactPhoneBusiness);
                item.SubItems.Add(contact.ContactPhoneHome);
                item.SubItems.Add(contact.ContactPhoneCell);
                item.SubItems.Add(contact.ContactCompany);

                item.Tag = contact;
                lvwContacts.Items.Add(item);
            }
        }

        private void deleteContact(clsContact contact)
        {
            OleDbConnection db = null;

            try
            {
                db = new OleDbConnection();
                db.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data source=" + mFileName;
                db.Open();

                string sql = "DELETE FROM Contacts WHERE ContactID=" + contact.ContactID;

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

        private void showErrorMessage(string msg)
        {
            MessageBox.Show(msg, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void showInformationMessage(string msg)
        {
            MessageBox.Show(msg, Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public DialogResult messageBoxYesNo(string msg) //private helper method for messageboxyesno
        {
            return MessageBox.Show(msg, Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        }

        private void lvwContacts_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            mSortOrder = !mSortOrder; //Reverse sort order each time column is clicked.

            lvwContacts.ListViewItemSorter = new clsListViewComparer(e.Column, mSortOrder);
        }

        private void lvwContacts_DoubleClick(object sender, EventArgs e)
        {
            if (lvwContacts.SelectedItems.Count == 0)
            {
                return;
            }
            else
            {
                showUserCard();
            }
        }

        private void showUserCard()
        {
            clsContact contact;

            //make sure an there is a selected item in the List Box
            if (lvwContacts.SelectedItems.Count == 0)
            {
                return;
            }

            contact = (clsContact)lvwContacts.SelectedItems[0].Tag; //tags the row to update

            string name = contact.ContactLastName + ", " + contact.ContactFirstName + " ";

            ViewContact viewContact = new ViewContact(name, contact);

            viewContact.ShowDialog();
        }

        private void addContact()
        {
            clsContact contact = new clsContact();

            UpdateContact updateContact = new UpdateContact("Add a new User", "Add", contact);

            updateContact.ShowDialog();

            //Rebuild table after adding or updating a user.
            readUsers();
            showUsers();
        }

        private void updateContact()
        {
            clsContact contact;

            //make sure an there is a selected item in the List Box
            if (lvwContacts.SelectedItems.Count == 0)
            {
                showInformationMessage("You must select a user to update.");
                return;
            }

            contact = (clsContact)lvwContacts.SelectedItems[0].Tag; //tags the row to update

            string name = contact.ContactLastName + ", " + contact.ContactFirstName + " ";

            UpdateContact updateContact = new UpdateContact("Update: " + name, "Update", contact);

            updateContact.ShowDialog();

            //Rebuild table after adding, updating, or deleting a contact.
            readUsers();
            showUsers();
        }

        private void deleteContact()
        {
            clsContact contact;

            //make sure an there is a selected item in the List Box
            if (lvwContacts.SelectedItems.Count == 0)
            {
                showInformationMessage("You must select a user to delete.");
                return;
            }

            contact = (clsContact)lvwContacts.SelectedItems[0].Tag; //tags the row to update

            string name = contact.ContactFirstName + " " + contact.ContactLastName;

            //Ask the user if they want to delete the contact
            DialogResult result;
            result = messageBoxYesNo("Are you sure you want to delete " + name + "?");
            if (result == DialogResult.Yes)
            {
                deleteContact(contact);
            }

            //Rebuild table after adding, updating, or deleting a contact.
            readUsers();
            showUsers();
        }

        private void deleteAllContacts()
        {
            OleDbConnection db = null;

            try
            {
                db = new OleDbConnection();
                db.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data source=" + mFileName;
                db.Open();

                string sql = "DELETE * FROM Contacts";

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

            //Rebuild table after adding, updating, or deleting a contact.
            readUsers();
            showUsers();
        }

        private void addUserToolStripMenuItem_Click(object sender, EventArgs e)
        {
            addContact();
        }

        private void updateUserToolStripMenuItem_Click(object sender, EventArgs e)
        {
            updateContact();
        }

        private void deleteUserToolStripMenuItem_Click(object sender, EventArgs e)
        {
            deleteContact();
        }

        private void addContactToolStripMenuItem_Click(object sender, EventArgs e)
        {
            addContact();
        }

        private void updateContactToolStripMenuItem_Click(object sender, EventArgs e)
        {
            updateContact();
        }

        private void deleteContactToolStripMenuItem_Click(object sender, EventArgs e)
        {
            deleteContact();
        }

        private void optionsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult result;
            Options options = new Options();
            result = options.ShowDialog();

            if (result != DialogResult.OK)
            {
            }
            else
            {
                //Redo the startup with new settings
                readSettings();
                goStartUp();
            }
        }

        private double calculateJulianDay(DateTime date)
        {
            //From http://quasar.as.utexas.edu/BillInfo/JulianDatesG.html

            //1) Express the date as Y M D, where Y is the year, M is the month 
            //   number (Jan = 1, Feb = 2, etc.), and D is the day in the month.
            //
            //2) If the month is January or February, subtract 1 from the year to 
            //   get a new Y, and add 12 to the month to get a new M.
            //   (Thus, we are thinking of January and February as being the 
            //   13th and 14th month of the previous year).
            //
            //3) Dropping the fractional part of all results of all 
            //   multiplications and divisions, let

            //  A = Y/100
            //  B = A/4
            //  C = 2-A+B
            //  E = 365.25x(Y+4716)
            //  F = 30.6001x(M+1)
            //  JD= C+D+E+F-1524.5

            double day, month, year;
            double a, b, c, e, f, jd;

            year = date.Year;
            month = date.Month;
            day = date.Day;

                //If Jan. or Feb.
            if (month == 1 || month == 2)
            {
                //Subtract one from year
                year--;
                //Add 12 to the month
                month += 12;
            }

            a = year / 100;
            //a = Math.Round(a);
            a = Math.Truncate(a);
            
            b = a / 4;
            //b = Math.Round(b);
            b = Math.Truncate(b);

            c = 2 - a + b;
            //c = Math.Round(c); Not a multiplication or division

            e = 365.25 * (year + 4716);
            //e = Math.Round(e);
            e = Math.Truncate(e);

            f = 30.6001 * (month + 1);
            //f = Math.Round(f);
            f = Math.Truncate(f);

            jd = c + day + e + f - 1524.5;
            jd = Math.Round(jd, 0, MidpointRounding.AwayFromZero);

            return jd;
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutBox about = new AboutBox();
            about.ShowDialog();
        }

        private void deleteAllContactsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Ask the user for confirmation
            DialogResult result;
            result = messageBoxYesNo("This will delete ALL CONTACTS in the database! Are you sure?");
            if (result == DialogResult.Yes)
            {
                // Ask the user for confirmation again
                result = messageBoxYesNo("Once again, this will delete ALL CONTACTS in the database! Are you really sure?");
                if (result == DialogResult.Yes)
                {
                    deleteAllContacts();
                }
                else
                {
                    return;
                }
            }
            else
            {
                return;
            }
        }
    }
}