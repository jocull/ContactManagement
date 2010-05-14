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
    public partial class Options : Form
    {

        string mFileName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Contacts.ctm");

        public Options()
        {
            InitializeComponent();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            byte warningDays;
            if (byte.TryParse(txtBirthdayDays.Text, out warningDays) == false)
            {
                showErrorMessage("Number of days for warning must be a number less than 255");
                txtBirthdayDays.Focus();
                return;
            }

            //Save options to database
            OleDbConnection db = null;

            try
            {
                //Database location string.
                db = new OleDbConnection();
                db.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data source=" + mFileName;
                db.Open();

                string sql = "UPDATE Settings "
                    + "SET SettingsBirthdayWarning=" + toSQL(warningDays)
                    + ", SettingsOwnerName=" + toSQL(txtYourName.Text)
                    + " WHERE SettingsKey='Update'";

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

            MessageBox.Show("You must restart the program for changes to take effect.", "Restart Required", MessageBoxButtons.OK, MessageBoxIcon.Information);

            //Close form after saving options
            Close();
        }

        private void showErrorMessage(string msg)
        {
            MessageBox.Show(msg, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private string toSQL(string str)
        {
            return "'" + str.Replace("'", "''") + "'";
        }

        private string toSQL(byte number)
        {
            return number.ToString();
        }

        private void Options_Load(object sender, EventArgs e)
        {
            //Load exisiting options from the database
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

                // Format each message returned from the query.
                while (dataReader.Read() == true)
                {
                    byte tempByte;
                    tempByte = (byte)dataReader["SettingsBirthdayWarning"];
                    txtBirthdayDays.Text = tempByte.ToString();
                    txtYourName.Text = (string)dataReader["SettingsOwnerName"].ToString();
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
    }
}