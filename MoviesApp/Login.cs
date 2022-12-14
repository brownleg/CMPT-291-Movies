using MoviesApp.SQL;
using System.Reflection.Metadata;
using System.Data.SqlClient;

namespace MoviesApp
{
    public partial class Login : Form
    {

        DBConnection connection;
        public Login(DBConnection input_connection)
        {
            InitializeComponent();
            connection = input_connection;
        }

        private void label1_Click(object sender, EventArgs e)
        {
           
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }


        private void loginButton_Click(object sender, EventArgs e)
        {

            if (accountTextInput.Text.StartsWith("E"))
            {
                string account_ID = accountTextInput.Text.Substring(1);
                string query = $"select employee_id from Employees where employee_id = {account_ID}";
                SqlDataReader? empdata = connection.GetDataReader(query); //get data reader must be initalized like line 9

                if (empdata != null && empdata.HasRows)
                {
                    empdata.Read();                                     //to see the data you must read() first
                    string ID = empdata["employee_id"].ToString();      //passing ID
                    empdata.Close();                //closes the reader after the data is read in

                    accountTextInput.Text = "";
                    passwordTextInput.Text = "";
                    new EmployeeViewForm(ID, connection).ShowDialog();

                }
                else
                {
                    MessageBox.Show("Invalid username or password");
                }

                //Console.WriteLine(empdata["employee_id"].ToString()); //call to the table, [col name]. ToString so its reable, can be looped with while for the empdata

            }
            else
            {
                string account_ID = accountTextInput.Text;
                string query = $"select account_number from Customer where account_number = {account_ID}";
                SqlDataReader? custdata = connection.GetDataReader(query); //get data reader must be initalized like line 9

                if (custdata != null && custdata.HasRows) //checks if there is data in the table as well as a table
                {
                    custdata.Read();                                     //to see the data you must read() first
                    string ID = custdata["account_number"].ToString();      //passing ID
                    custdata.Close();                //closes the reader after the data is read in

                    accountTextInput.Text = "";
                    passwordTextInput.Text = "";
                    new CustomerViewForm(ID, connection).ShowDialog();

                }
                else
                {
                    MessageBox.Show("Invalid username or password");
                }
            }
            
        }

    }
}