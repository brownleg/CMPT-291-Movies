using System.Data.SqlClient;
using System.DirectoryServices;
using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.VisualBasic.Devices;
using MoviesApp.SQL;

namespace MoviesApp
{
    public partial class EmployeeViewForm : Form
    {
        private string ID = ""; //id initalized to use

        private DBConnection connection;
        public EmployeeViewForm(string input, DBConnection input_connection) //takes the id in
        {
            InitializeComponent(); //create the forms
            connection = input_connection;
            ID = input; //call this ID to get the current id, needs to be sent to others as well when tranfering
        }

        //dataGridView2
        private void hideSearchElements()
        {
            movieNameTextBox.Visible = false;
            genreComboBox.Visible = false;
            actorsNamesTextBox.Visible = false;
            moviesSearchClearButton.Visible = false;
            moviesSearchButton.Visible = false;
        }
        private void showSearchElements()
        {
            movieNameTextBox.Visible = true;
            genreComboBox.Visible = true;
            actorsNamesTextBox.Visible = true;
            moviesSearchClearButton.Visible = true;
            moviesSearchButton.Visible = true;
        }

        private void searchAllMovies(Object sender, EventArgs e)
        {
            hideSearchElements();
            empMoviesDataGridView.Rows.Clear();

            string query = $"select M.movie_id, M.movie_name, " +

                $"case when exists (select * from Movie_type T2 where T2.movie_id = M.movie_id) " +
                $"then temp.genres else '' end as genres " +

                $"from Movie M left join (select M2.movie_id, STRING_AGG(T.type_of_movie, ', ') as genres " +
                $"from Movie M2, Movie_type T where M2.movie_id = T.movie_id group by M2.movie_id) temp on temp.movie_id = M.movie_id " +
                $"order by M.movie_name";

            SqlDataReader? empdata = connection.GetDataReader(query);

            while (empdata != null && empdata.Read())
            {
                int movie_id = Int32.Parse(empdata["movie_id"].ToString());
                empMoviesDataGridView.Rows.Add(empdata["movie_id"].ToString(), empdata["movie_name"].ToString(), empdata["genres"].ToString());
            }

            if (empdata != null)
                empdata.Close();                //closes the reader after the data is read in

        }

        private void clearMoviesSearchForm(Object sender, EventArgs e)
        {
            movieNameTextBox.Clear();
            actorsNamesTextBox.Clear();
            genreComboBox.Text = "Genre";
        }

        private void newSearchFilters(Object sender, EventArgs e)
        {
            showSearchElements();
            empMoviesDataGridView.Rows.Clear();
        }
        
        private void searchActors(Object sender, EventArgs e)
        {
            empActorsDataGridView.Rows.Clear();

            string[] actorsSearchParams = empActorSearchTextBox.Text.Split(' ');

            string query = $"select distinct actor_id, first_name, last_name, gender, age " +
                $"from Actors ";


            if (actorsSearchParams.Count() > 0) {

                query += "where ";

                for (int i = 0; i < actorsSearchParams.Count(); i++) {
                    query += $"(first_name like '%{actorsSearchParams[i]}%' or last_name like '%{actorsSearchParams[i]}%') ";

                    if (i < actorsSearchParams.Count() - 1)
                        query += "and ";
                }
            }

            query += "order by first_name";

            SqlDataReader? empdata = connection.GetDataReader(query);

            while (empdata != null && empdata.Read())
            {
                int actor_id = Int32.Parse(empdata["actor_id"].ToString()!);
                empActorsDataGridView.Rows.Add(empdata["actor_id"].ToString(), empdata["first_name"].ToString(), empdata["last_name"].ToString(), empdata["gender"].ToString(), empdata["age"].ToString());
            }

            if (empdata != null)
                empdata.Close();                //closes the reader after the data is read in
        }

        private void searchMoviesWithFilters(Object sender, EventArgs e)
        {
            empMoviesDataGridView.Rows.Clear();

            // If actors are specified in the search bar then get a comma-separated list
            string[] actors = actorsNamesTextBox.Text.Split(',');

            string query = $"select distinct M.movie_id, M.movie_name, " +

                $"case when exists (select * from Movie_type T2 where T2.movie_id = M.movie_id) " +
                $"then temp.genres else '' end as genres, " +

                $"case when exists (select * from Acts_in AI2 where AI2.movie_id = M.movie_id) " +
                $"then temp2.actors else '' end as actors " +

                $"from Movie_type T2, Movie M " +

                $"left join (select M2.movie_id, STRING_AGG(T.type_of_movie, ', ') as genres " +
                $"from Movie M2, Movie_type T where M2.movie_id = T.movie_id group by M2.movie_id) temp on temp.movie_id = M.movie_id " +

                $"left join (select M3.movie_id, STRING_AGG(A2.first_name + ' ' + A2.last_name, ', ') as actors " +
                $"from Movie M3, Acts_in AI, Actors A2 where M3.movie_id = AI.movie_id and A2.actor_id = AI.actor_id " +
                $"group by M3.movie_id) temp2 on temp2.movie_id = M.movie_id " +

                $"where M.movie_name like '%{movieNameTextBox.Text}%' ";

            if (actors.Length > 0)
                query += $"and temp2.actors like '%" + String.Join("%' and temp2.actors like '%", actors) + "%' ";

            query += $"and ('{genreComboBox.Text}' = 'Genre' or genres like '%{genreComboBox.Text}%') " +
                $"order by M.movie_name";

            SqlDataReader? empdata = connection.GetDataReader(query);

            while (empdata != null && empdata.Read())
            {
                int movie_id = Int32.Parse(empdata["movie_id"].ToString()!);
                empMoviesDataGridView.Rows.Add(empdata["movie_id"].ToString(), empdata["movie_name"].ToString(), empdata["genres"].ToString());
            }

            if (empdata != null)
                empdata.Close();                //closes the reader after the data is read in
        }

        private void mostActivelyRented(object sender, EventArgs e)
        {
            hideSearchElements();
            empMoviesDataGridView.Rows.Clear();

            string query = $"select M.movie_id, M.movie_name, " +

                $"case when exists (select * from Movie_type T2 where T2.movie_id = M.movie_id) " +
                $"then temp.genres else '' end as genres, " +

                $"isnull(temp2.count, 0) as count " +

                $"from Movie M left join (select M2.movie_id, STRING_AGG(T.type_of_movie, ', ') as genres from Movie M2, Movie_type T " +
                $"where M2.movie_id = T.movie_id group by M2.movie_id) temp on temp.movie_id = M.movie_id " +

                $"left join (select O.movie_id, count(*) as count from Orders O group by O.movie_id) temp2 on temp2.movie_id = M.movie_id " +

                $"order by temp2.count desc, M.movie_name asc";

            SqlDataReader? empdata = connection.GetDataReader(query);

            while (empdata != null && empdata.Read())
            {
                empMoviesDataGridView.Rows.Add(empdata["movie_id"].ToString(), empdata["movie_name"].ToString(), empdata["genres"].ToString());
            }

            if (empdata != null)
                empdata.Close();

        }
        private void empLogoutButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }
            private void launchMovieButton_Click(object sender, EventArgs e)
        {
            new MovieForm(ID, connection, -1).ShowDialog();
        }
        private void actorsAddButton_Click(object sender, EventArgs e)
        {
            new ActorForm(ID, connection, -1).ShowDialog();

        }
        private void launchCustomerButton_Click(object sender, EventArgs e)
        {
            new CustomerForm().ShowDialog();

        }
        private void launchEmployeeButton_Click(object sender, EventArgs e)
        {
            new EmployeeForm(ID).ShowDialog();

        }
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (tabControl1.SelectedIndex)
            {
                case 0: this.AcceptButton = moviesSearchButton;break;
                case 1: this.AcceptButton = actorsSearchButton;break;
                case 2: this.AcceptButton = searchCustomersButton;break;
                case 3: this.AcceptButton = searchEmployeesButton;break;
                case 4: this.AcceptButton = transactionsSearchButton; break;
                case 5: this.AcceptButton = reportsGenerateButton;break;
                default: this.AcceptButton = moviesSearchButton;break;
            }
        }

        private void EmployeeViewForm_Load(object sender, EventArgs e)
        {
            empMoviesDataGridView.AllowUserToAddRows = false;
            empActorsDataGridView.AllowUserToAddRows = false;

            // Set default search button on page load
            tabControl1.TabIndex = 0;
            this.AcceptButton = moviesSearchButton;

            // Populate combo box on movies page
            string query = $"select * from Genre order by type_of_movie";
            SqlDataReader? empdata = connection.GetDataReader(query);

            while (empdata.Read())
            {
                genreComboBox.Items.Add(empdata["type_of_movie"].ToString());
            }

            // Close the reader after data is read in
            empdata.Close();
        }

        private void empMovies_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0)
                return;

            // Get clicked movie id
            int clickedMovieId = Int32.Parse(empMoviesDataGridView.Rows[e.RowIndex].Cells[0].Value.ToString());
            
            // Pass movie id of clicked-on row into MovieForm
            MovieForm f2 = new MovieForm(ID, connection, clickedMovieId);

            // Open the window
            f2.ShowDialog(); //showing form after creation

        }

        private void empActors_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
           
            if (e.RowIndex < 0)
                return;

            // Get clicked movie id
            int clickedActorId = Int32.Parse(empActorsDataGridView.Rows[e.RowIndex].Cells[0].Value.ToString());

            // Pass movie id of clicked-on row into MovieForm
            ActorForm f2 = new ActorForm(ID, connection, clickedActorId);

            // Open the window
            f2.ShowDialog(); //showing form after creation

        }





        /************** Transactions Page ********************/

        private void empNewTransactionButton_Click(object sender, EventArgs e)
        {
            // Pass movie id of clicked-on row into MovieForm
            TransactionForm f = new TransactionForm(ID, connection);

            // Open the window
            f.ShowDialog(); //showing form after creation
        }

        
        /******************* reports **************************/
        public int RBCN = 1; //report button click number, picks the report to generate when a report button is chosen
        public string Month_picked = "0"; //basic starting month chosen, full names
        public int Year_picked = 2022; //starting year
        public int Quarter_picked = 0; //quarter picked, First Quarter, Second Quarter, Third Quarter, Forth Quarter, full year is 5

        //when clicking a button, on of the grid views will display
        private void report2Button_Click_1(object sender, EventArgs e)
        {
            reportsDataGridView.Visible = false;
            dataGridView2.Visible = true;
            dataGridView3.Visible = false;
            dataGridView6.Visible = false;
            dataGridView5.Visible = false;
            reportsDescriptionTextBox.Text = "Report 2: Best Movie Genre";
            dataGridView2.BringToFront();
            RBCN = 2;
        }
        private void report3Button_Click_1(object sender, EventArgs e)
        {
            reportsDataGridView.Visible = false;
            dataGridView2.Visible = false;
            dataGridView3.Visible = true;
            dataGridView6.Visible = false;
            dataGridView5.Visible = false;
            reportsDescriptionTextBox.Text = "Report 3: top 5 movies per genre";
            dataGridView3.BringToFront();
            RBCN = 3;
        }
        private void report4Button_Click_1(object sender, EventArgs e)
        {
            reportsDataGridView.Visible = false;
            dataGridView2.Visible = false;
            dataGridView3.Visible = false;
            dataGridView6.Visible = true;
            dataGridView5.Visible = false;
            reportsDescriptionTextBox.Text = "Report 4: profit earnings";
            dataGridView6.BringToFront();
            RBCN = 4;
        }
        private void report5Button_Click_1(object sender, EventArgs e)
        {
            reportsDataGridView.Visible = false;
            dataGridView2.Visible = false;
            dataGridView3.Visible = false;
            dataGridView6.Visible = false;
            dataGridView5.Visible = true;
            reportsDescriptionTextBox.Text = "Report 5: people to recomend to upgrade based on usage";
            dataGridView5.BringToFront();
            RBCN = 5;
        }

        private void report1Button_Click1(object sender, EventArgs e)
        {
            reportsDataGridView.Visible = true;
            dataGridView2.Visible = false;
            dataGridView3.Visible = false;
            dataGridView6.Visible = false;
            dataGridView5.Visible = false;
            reportsDescriptionTextBox.Text = "Report 1: profit per month";
            reportsDataGridView.BringToFront();
            RBCN = 1;
        }

        private void QuarterComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            MonthComboBox.Text = "Month";
            Month_picked = "0";
            MonthComboBox.Refresh();
            string Quarter_picked_str = QuarterComboBox.Text;
            switch (Quarter_picked_str)
            {
                case "First Quarter":
                    Quarter_picked = 1;
                    break;
                case "Second Quarter":
                    Quarter_picked = 2;
                    break;
                case "Third Quarter":
                    Quarter_picked = 3;
                    break;
                case "Forth Quarter":
                    Quarter_picked = 4;
                    break;
                case "Full Year":
                    Quarter_picked = 5;
                    break;
                
            }
        }

        private void MonthComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            QuarterComboBox.Text = "Quarter";
            Quarter_picked = 0;
            QuarterComboBox.Refresh();
            string Month_picked_str = MonthComboBox.Text;
            int Month_picked_int = DateTime.ParseExact(Month_picked_str, "MMMM", CultureInfo.CreateSpecificCulture("en-GB")).Month;
            if (Month_picked_int < 10)
            {
                string b = "0";
                Month_picked = b + Month_picked_int.ToString();
            }
        }

        private void reportsGenerateButton_Click(object sender, EventArgs e)
        {
            //needs to know all the info from each value to call functions and send over info
            switch (RBCN) //based on the value of RBCN it will display the info of the report
            {
                case 1:
                    string full_date;
                    reportsDataGridView.Rows.Clear();
                    if (Quarter_picked == 0) //short, just the month
                    {
                        string revenue = "0";
                        string costs = "0";
                        full_date = Year_picked.ToString() + Month_picked + "01";


                        string query = $"select sum(temp2.revenue) as total_revenue " +
                                       $"from(select count(*) as plan_count, temp.cost, (count(*) * temp.cost) as revenue " +
                                       $"from(select P.plan_number, C.account_number, P.cost " +
                                       $"from Customer C, Plans P " +
                                       $"where P.plan_number = C.plan_number and C.start_date < '{full_date}' and C.end_date >= '{full_date}') as temp " +
                                       $"group by temp.plan_number, temp.cost) as temp2";

                        SqlDataReader? revdata = connection.GetDataReader(query);
                        if (revdata != null && revdata.HasRows)
                        {
                            revdata.Read();                                     //to see the data you must read() first
                            revenue = revdata["total_revenue"].ToString();      //passing ID
                            revdata.Close();                //closes the reader after the data is read in
                        }
                        
                        query = $"select sum(M.distribution_fee) as total_movie_cost " +
                                $"from Movie M;";
                        SqlDataReader? costdata = connection.GetDataReader(query);
                        if (costdata != null && costdata.HasRows)
                        {
                            costdata.Read();                                     //to see the data you must read() first
                            costs = costdata["total_movie_cost"].ToString();      //passing ID
                            costdata.Close();                //closes the reader after the data is read in
                        }

                        float profits = float.Parse(revenue) - float.Parse(costs);
                        reportsDataGridView.Rows.Add(revenue, costs, profits);

                    }
                    else //long, looping search and add
                    {
                        switch (Quarter_picked)
                        {
                            case 1:
                                double revenue1 = 0;
                                string revenue_1 = "0";
                                string costs1 = "0";
                                string query1;
                                var Q1 = new string[3];
                                Q1[0] = Year_picked.ToString() + "0101";
                                Q1[1] = Year_picked.ToString() + "0201";
                                Q1[2] = Year_picked.ToString() + "0301";
                                foreach (string full_date1 in Q1) {
                                    
                                    query1 = $"select sum(temp2.revenue) as total_revenue " +
                                                   $"from(select count(*) as plan_count, temp.cost, (count(*) * temp.cost) as revenue " +
                                                   $"from(select P.plan_number, C.account_number, P.cost " +
                                                   $"from Customer C, Plans P " +
                                                   $"where P.plan_number = C.plan_number and C.start_date < '{full_date1}' and C.end_date >= '{full_date1}') as temp " +
                                                   $"group by temp.plan_number, temp.cost) as temp2";

                                    SqlDataReader? revdata1 = connection.GetDataReader(query1);
                                    if (revdata1 != null && revdata1.HasRows)
                                    {
                                        revdata1.Read();                  //to see the data you must read() first
                                        revenue_1 = revdata1["total_revenue"].ToString(); //.ToString();     
                                        revdata1.Close();                //closes the reader after the data is read in
                                        revenue1 = revenue1 + Convert.ToDouble(revenue_1);
                                    }

                                }
                                query1 = $"select sum(M.distribution_fee) as total_movie_cost " +
                                        $"from Movie M;";
                                SqlDataReader? costdata1 = connection.GetDataReader(query1);
                                if (costdata1 != null && costdata1.HasRows)
                                {
                                    costdata1.Read();                                     //to see the data you must read() first
                                    costs1 = costdata1["total_movie_cost"].ToString();
                                    costdata1.Close();                //closes the reader after the data is read in
                                }

                                double profits1 = revenue1 - (float.Parse(costs1) * 3);
                                reportsDataGridView.Rows.Add(revenue1, costs1, profits1);

                                return;

                            case 2:

                                var Q2 = new string[3];
                                Q2[0] = Year_picked.ToString() + "0401";
                                Q2[1] = Year_picked.ToString() + "0501";
                                Q2[2] = Year_picked.ToString() + "0601";
                                double revenue2 = 0;
                                string revenue_2 = "0";
                                string costs2 = "0";
                                string query2;
                                foreach (string full_date2 in Q2)
                                {

                                    query2 = $"select sum(temp2.revenue) as total_revenue " +
                                                   $"from(select count(*) as plan_count, temp.cost, (count(*) * temp.cost) as revenue " +
                                                   $"from(select P.plan_number, C.account_number, P.cost " +
                                                   $"from Customer C, Plans P " +
                                                   $"where P.plan_number = C.plan_number and C.start_date < '{full_date2}' and C.end_date >= '{full_date2}') as temp " +
                                                   $"group by temp.plan_number, temp.cost) as temp2";

                                    SqlDataReader? revdata2 = connection.GetDataReader(query2);
                                    if (revdata2 != null && revdata2.HasRows)
                                    {
                                        revdata2.Read();                  //to see the data you must read() first
                                        revenue_2 = revdata2["total_revenue"].ToString(); //.ToString();     
                                        revdata2.Close();                //closes the reader after the data is read in
                                        revenue2 = revenue2 + Convert.ToDouble(revenue_2);
                                    }

                                }
                                query2 = $"select sum(M.distribution_fee) as total_movie_cost " +
                                        $"from Movie M;";
                                SqlDataReader? costdata2 = connection.GetDataReader(query2);
                                if (costdata2 != null && costdata2.HasRows)
                                {
                                    costdata2.Read();                                     //to see the data you must read() first
                                    costs2 = costdata2["total_movie_cost"].ToString();
                                    costdata2.Close();                //closes the reader after the data is read in
                                }

                                double profits2 = revenue2 - (float.Parse(costs2) * 3);
                                reportsDataGridView.Rows.Add(revenue2, costs2, profits2);
                                return;

                            case 3:
                                var Q3 = new string[3];
                                Q3[0] = Year_picked.ToString() + "0701";
                                Q3[1] = Year_picked.ToString() + "0801";
                                Q3[2] = Year_picked.ToString() + "0901";
                                double revenue3 = 0;
                                string revenue_3 = "0";
                                string costs3 = "0";
                                string query3;
                                foreach (string full_date3 in Q3)
                                {

                                    query3 = $"select sum(temp2.revenue) as total_revenue " +
                                                   $"from(select count(*) as plan_count, temp.cost, (count(*) * temp.cost) as revenue " +
                                                   $"from(select P.plan_number, C.account_number, P.cost " +
                                                   $"from Customer C, Plans P " +
                                                   $"where P.plan_number = C.plan_number and C.start_date < '{full_date3}' and C.end_date >= '{full_date3}') as temp " +
                                                   $"group by temp.plan_number, temp.cost) as temp2";

                                    SqlDataReader? revdata3 = connection.GetDataReader(query3);
                                    if (revdata3 != null && revdata3.HasRows)
                                    {
                                        revdata3.Read();                  //to see the data you must read() first
                                        revenue_3 = revdata3["total_revenue"].ToString(); //.ToString();     
                                        revdata3.Close();                //closes the reader after the data is read in
                                        revenue3 = revenue3 + Convert.ToDouble(revenue_3);
                                    }

                                }
                                query3 = $"select sum(M.distribution_fee) as total_movie_cost " +
                                        $"from Movie M;";
                                SqlDataReader? costdata3 = connection.GetDataReader(query3);
                                if (costdata3 != null && costdata3.HasRows)
                                {
                                    costdata3.Read();                                     //to see the data you must read() first
                                    costs3 = costdata3["total_movie_cost"].ToString();
                                    costdata3.Close();                //closes the reader after the data is read in
                                }

                                double profits3 = revenue3 - (float.Parse(costs3) * 3);
                                reportsDataGridView.Rows.Add(revenue3, costs3, profits3);
                                return;
                                

                            case 4:
                                var Q4 = new string[3];
                                Q4[0] = Year_picked.ToString() + "1001";
                                Q4[1] = Year_picked.ToString() + "1101";
                                Q4[2] = Year_picked.ToString() + "1201";
                                double revenue4 = 0;
                                string revenue_4 = "0";
                                string costs4 = "0";
                                string query4;
                                foreach (string full_date4 in Q4)
                                {

                                    query4 = $"select sum(temp2.revenue) as total_revenue " +
                                                   $"from(select count(*) as plan_count, temp.cost, (count(*) * temp.cost) as revenue " +
                                                   $"from(select P.plan_number, C.account_number, P.cost " +
                                                   $"from Customer C, Plans P " +
                                                   $"where P.plan_number = C.plan_number and C.start_date < '{full_date4}' and C.end_date >= '{full_date4}') as temp " +
                                                   $"group by temp.plan_number, temp.cost) as temp2";

                                    SqlDataReader? revdata4 = connection.GetDataReader(query4);
                                    if (revdata4 != null && revdata4.HasRows)
                                    {
                                        revdata4.Read();                  //to see the data you must read() first
                                        revenue_4 = revdata4["total_revenue"].ToString(); //.ToString();     
                                        revdata4.Close();                //closes the reader after the data is read in
                                        revenue4 = revenue4 + Convert.ToDouble(revenue_4);
                                    }

                                }
                                query4 = $"select sum(M.distribution_fee) as total_movie_cost " +
                                        $"from Movie M;";
                                SqlDataReader? costdata4 = connection.GetDataReader(query4);
                                if (costdata4 != null && costdata4.HasRows)
                                {
                                    costdata4.Read();                                     //to see the data you must read() first
                                    costs4 = costdata4["total_movie_cost"].ToString();
                                    costdata4.Close();                //closes the reader after the data is read in
                                }

                                double profits4 = revenue4 - (float.Parse(costs4) * 3);
                                reportsDataGridView.Rows.Add(revenue4, costs4, profits4);
                                return;

                            case 5:


                                return;

                        }
                    }

                    break;

                case 2:

                    break;

                case 3:

                    break;
                case 4:

                    break;

                case 5:

                    break;

            }
        }

        private void YearComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            Year_picked = Int32.Parse(YearComboBox.Text);
            YearComboBox.Text = Year_picked.ToString();
        }
    }
}