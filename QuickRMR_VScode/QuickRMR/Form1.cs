using System;
using System.IO;
using System.Diagnostics;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Threading;

namespace QuickRMR
{
    public partial class Form1 : Form
    {
        SqlConnection con = new SqlConnection(@"Data Source = (LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\RMR.mdf; Integrated Security = True;Connect Timeout=30");
        double brmr = 0, rmr = 0, ucs, rucs = 0, plsi, rplsi = 0, rqd, rrqd = 0, spd, rspd = 0, gwi, rgwi = 0, jwp, rjwp = 0, rgw = 0, dl, rdl = 0, da, rda = 0, jrc, rjrc = 0, 
            hi, rhi = 0, si, rsi = 0,rni = 0, sw, Ru5 = 0, Rw1 = 0, rsw = 0, rdc= 0, dar = 0, rdd=0, ralt = 0, Fe = 1, Fs = 1;
        int selectedRow;
        public Form1() // initialize form with display data
        {
            Thread splash = new Thread(new ThreadStart(StartForm));
            splash.Start();
            Thread.Sleep(5000);
            InitializeComponent();
            splash.Abort();
        }

        public void StartForm()
        {
            Application.Run(new frmSplashScreen());
        }

        private void button1_Click(object sender, EventArgs e) // opens RMR89
        {
            panel1.Visible = true;
            panel2.Visible = false;
            disp_data();
        }

        private void button2_Click(object sender, EventArgs e) // opens RMR14
        {
            panel2.Visible = true;
            panel1.Visible = false;
            disp_data();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e) // gets row index of a selected cell in datagrid  view
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                selectedRow = e.RowIndex;
                DataGridViewRow row = dataGridView1.Rows[selectedRow];
            }

        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                selectedRow = e.RowIndex;
                DataGridViewRow row = dataGridView2.Rows[selectedRow];
            }
        }

        //RMR 89

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e) // chnages unit according to groundwater options
        {
            string optioniii = comboBox2.Text;
            string unit;

            switch (optioniii)
            {
                case "GW Inflow":
                    unit = "l/min/m";
                    label17.Text = unit;
                    break;
                case "JW Pressure":
                    unit = "MPa";
                    label17.Text = unit;
                    break;
            }
        }

        //RMR 14

        private void comboBox9_SelectedIndexChanged(object sender, EventArgs e)
        {
            string optioniii = comboBox9.Text;
            string unit;

            switch (optioniii)
            {
                case "GW Inflow":
                    unit = "l/min/m";
                    label48.Text = unit;
                    break;
                case "JW Pressure":
                    unit = "MPa";
                    label48.Text = unit;
                    break;
            }
        }

        //DELETE all data from selected database table, clear text boxes and reset comboboxes

        private void newToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult dialog = MessageBox.Show("All Record and entries will be deleted\n Do you want to Proceed?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (dialog == DialogResult.Yes)
            {
                con.Open();
                if (panel1.Visible) //RMR89
                {
                    string sqlTrunc = "TRUNCATE TABLE " + "rmr89";
                    SqlCommand cmd = new SqlCommand(sqlTrunc, con);
                    cmd.CommandType = CommandType.Text;
                    cmd.ExecuteNonQuery();
                }

                if (panel2.Visible) //RMR14
                {
                    string sqlTrunc = "TRUNCATE TABLE " + "rmr14";
                    SqlCommand cmd = new SqlCommand(sqlTrunc, con);
                    cmd.CommandType = CommandType.Text;
                    cmd.ExecuteNonQuery();
                }

                con.Close();
                disp_data();
                errorProvider1.Clear();
                ClearTextBoxes(this);
                clearlistbox();
            }
        }
        
        //CHECK

        private void exportToolStripMenuItem_Click(object sender, EventArgs e) // EXPORT to excel call
        {
                export();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e) // EXIT from menu item
        {
            Environment.Exit(0);
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e) // EXIT application with close button prompt
        {
            DialogResult dialog = MessageBox.Show("Do you really want to exit?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (dialog == DialogResult.Yes)
            {
                Application.ExitThread();
            }
            else if (dialog == DialogResult.No)
            {
                e.Cancel = true;
            }
        }

        private void rMR89ToolStripMenuItem_Click(object sender, EventArgs e) // RMR89 visible
        {
            panel1.Visible = true;
            panel2.Visible = false;
            disp_data();
        }

        private void rMR14ToolStripMenuItem_Click(object sender, EventArgs e) // RMR14 visible
        {
            panel2.Visible = true;
            panel1.Visible = false;
            disp_data();
        }

        private void comboBox4_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            string optioniv = comboBox4.Text;
            string description;

            switch (optioniv)
            {
                case "Excellent":
                    description = "Very rough surfaces\nNot continuous\nNo separation\nUnweathered wall rock";
                    label32.Text = description;
                    break;
                case "Good":
                    description = "Slightly rough surfaces\nSeparation <1mm\nSlightly weathered walls";
                    label32.Text = description;
                    break;
                case "Fair":
                    description = "Slightly rough surfaces\nSeparation <1mm\nHighly weathered walls";
                    label32.Text = description;
                    break;
                case "Bad":
                    description = "Slickensided surfaces\nor\nGouge<5mm thick\nor\nSeparation 1-5mm\nContinuous";
                    label32.Text = description;
                    break;
                case "Very Bad":
                    description = "Soft gouge>5mm thick\nor\nSeparation >5mm\nContinuous";
                    label32.Text = description;
                    break;
            }
        }

        private void radioButton1_CheckedChanged_1(object sender, EventArgs e) //RMR 89 UCS/PLSI radiobutton
        {
            comboBox1.Visible = true;
            numericUpDown1.Visible = true;
            label14.Visible = true;
            numericUpDown2.Visible = false;
            numericUpDown3.Visible = false;
            label66.Visible = false;
            label67.Visible = false;
            label68.Visible = false;
        }

        private void radioButton2_CheckedChanged_1(object sender, EventArgs e) //RMR 89 R value radiobutton
        {
            numericUpDown2.Visible = true;
            numericUpDown3.Visible = true;
            label66.Visible = true;
            label67.Visible = true;
            label68.Visible = true;
            comboBox1.Visible = false;
            numericUpDown1.Visible = false;
            label14.Visible = false;
        }

        private void radioButton3_CheckedChanged_1(object sender, EventArgs e) // GENERAL Ground water condition options
        {
            comboBox3.Visible = true;
            numericUpDown7.Visible = true;
            label25.Visible = true;
            comboBox2.Visible = false;
            numericUpDown6.Visible = false;
            label17.Visible = false;
        }

        private void radioButton4_CheckedChanged_1(object sender, EventArgs e) // MEASURED Ground water condition options
        {
            comboBox2.Visible = true;
            numericUpDown6.Visible = true;
            label17.Visible = true;
            comboBox3.Visible = false;
            numericUpDown7.Visible = false;
            label25.Visible = false;
        }

        private void radioButton5_CheckedChanged(object sender, EventArgs e) // OVERALL discontinuity conditions options
        {
            panel7.Visible = true;
            comboBox4.Visible = true;
            numericUpDown9.Visible = true;
            label31.Visible = true;
            panel3.Visible = false;
        }

        private void radioButton6_CheckedChanged(object sender, EventArgs e) // QUANTIFIED PARAMETER BASED discontinuity conditions options
        {
            panel3.Visible = true;
            comboBox4.Visible = false;
            numericUpDown9.Visible = false;
            label31.Visible = false;
            panel7.Visible = false;
        }

        private void radioButton7_CheckedChanged(object sender, EventArgs e)
        {
            numericUpDown17.Visible = true;
            numericUpDown18.Visible = true;
            label61.Visible = true;
            label64.Visible = true;
            label65.Visible = true;
            comboBox8.Visible = false;
            numericUpDown16.Visible = false;
            label45.Visible = false;
        }

        private void radioButton8_CheckedChanged(object sender, EventArgs e)
        {
            comboBox8.Visible = true;
            numericUpDown16.Visible = true;
            label45.Visible = true;
            numericUpDown17.Visible = false;
            numericUpDown18.Visible = false;
            label61.Visible = false;
            label64.Visible = false;
            label65.Visible = false;
        }

        private void radioButton9_CheckedChanged(object sender, EventArgs e)
        {
            comboBox10.Visible = true;
            numericUpDown21.Visible = true;
            label54.Visible = true;
            comboBox9.Visible = false;
            numericUpDown20.Visible = false;
            label48.Visible = false;
        }

        private void radioButton10_CheckedChanged(object sender, EventArgs e)
        {
            comboBox9.Visible = true;
            numericUpDown20.Visible = true;
            label48.Visible = true;
            comboBox10.Visible = false;
            numericUpDown21.Visible = false;
            label54.Visible = false;
        }

        //RMR89 Calculate

        private void button3_Click(object sender, EventArgs e)
        {
            rmr = 0; brmr = 0; rucs = 0; rplsi = 0; rrqd = 0; rspd = 0; rgwi = 0; rjwp = 0; rgw = 0; rdl = 0; rda = 0; rjrc = 0; rhi = 0; rsi = 0; rni = 0; rsw = 0; rdc = 0; dar = 0;
            errorProvider1.Clear();
            clearlistbox();

            STRENGTH();  // Strength of Intact rock
            RQD();  //Rock qualty Designation
            SPD(); //Spacing of Discontinuities
            GWC(); //Ground water condition

            if (radioButton6.Checked) // QUANTITATIVE discontinuity condition active
            {
                DL(); //Discontinuity length (persistence)
                DA(); //Discontinuity Aperture
                JRC(); //Roughness (Joint Roughness Coefficient)
                INFL(); //Infilling of the discontinuities
                SW(); // discontinuity Surface weathering
            }
            else if (radioButton5.Checked) // GENERAL discontinuity condition active
            {
                DC(); //General Discontinuity conditions
            }
            DAR(); // Adjustment rating for discontinuity orientation
            RMR(); //Basic RMR and RMR

            if (radioButton6.Checked && jrc<=20 && rqd<=100 && dar>=-100 && Rw1<=Ru5) // INSERT to database table
            {
                con.Open();
                SqlCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "insert into [rmr89] values('" + textBox1.Text + "'," + rucs + "," + rplsi + "," + rrqd + "," + rspd + "," + rgw + "," + rgwi + "," + rjwp + "," + rdc+ "," + rdl + "," + rda + "," + rjrc + "," + rhi + "," + rsi + "," + rni + "," + rsw + "," + brmr + "," + dar + "," + rmr + ")";
                cmd.ExecuteNonQuery();
                con.Close();
            }

            else if (radioButton5.Checked && rdc <=30) // INSERT to database table
            {
                con.Open();
                SqlCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "insert into rmr89 values('" + textBox1.Text + "'," + rucs + "," + rplsi + "," + rrqd + "," + rspd + "," + rgw + "," + rgwi + "," + rjwp + "," + rdc + "," + rdl + "," + rda + "," + rjrc + "," + rhi + "," + rsi + "," + rni + "," + rsw + "," + brmr + "," + dar + "," + rmr + ")";
                cmd.ExecuteNonQuery();
                con.Close();
            }
            disp_data();
        }

        //RMR14 CALCULATE

        private void button9_Click(object sender, EventArgs e)
        {
             rmr = 0; brmr = 0; rucs = 0; rplsi = 0; dar = 0; rgwi = 0; rjwp = 0; rgw = 0; rdl = 0;
                rjrc = 0; rhi = 0; rsi = 0; rni = 0; rsw = 0; rdd = 0; ralt = 0; Fe = 1; Fs = 1;
             errorProvider1.Clear();
             clearlistbox();
             
             // CALLING ALL PARAMETERS' FUNCTIONS;

             STRENGTH14();  // Strength of Intact rock
             DD14(); //Number of Discontinuities per meter
             GWC14(); //Ground water condition
             listBox1.Items.Add(" ");
             listBox1.Items.Add("Joint Condition: ");
             DL14(); //Discontinuity length (persistence)
             JRC14(); //Roughness (Joint Roughness Coefficient)
             INFL14(); //Infilling of the discontinuities
             SW14(); // discontinuity Surface weathering
             ALT14(); // Alterability
             DAR14(); // Rating for discontinuity orientation adjustment
             FE(); //  Adjustment Rating for Exacavation method
             FS(); //  Adjustment Rating forStress-Strain behaviour along tunnel face
             RMR14(); //Basic RMR and RMR
                con.Open();
                SqlCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "insert into rmr14 values('" + textBox16.Text + "'," + rucs + "," + rplsi + "," + rdd + "," + rgw + "," + rgwi + "," + rjwp + "," + rdl + ","  + rjrc + "," + rhi + "," + rsi + "," + rni + "," + rsw + ","+ ralt + "," + brmr + "," + dar + "," + Fe + "," + Fs + "," + rmr + ")";
                cmd.ExecuteNonQuery();
                con.Close();    
            disp_data();
        }

        // RMR14 UPDATE (CALLING ALL PARAMETERS' FUNCTIONS)

        private void button8_Click(object sender, EventArgs e)
        {
            rmr = 0; brmr = 0; rucs = 0; rplsi = 0; dar = 0; rgwi = 0; rjwp = 0; rgw = 0; rdl = 0;
            rjrc = 0; rhi = 0; rsi = 0; rni = 0; rsw = 0; rdd = 0; ralt = 0; Fe = 1; Fs = 1;
            errorProvider1.Clear();
            clearlistbox();

            // CALLING ALL PARAMETERS' FUNCTIONS;

            STRENGTH14();  // Strength of Intact rock
            DD14(); //Number of Discontinuities per meter
            GWC14(); //Ground water condition
            listBox1.Items.Add(" ");
            listBox1.Items.Add("Joint Condition: ");
            DL14(); //Discontinuity length (persistence)
            JRC14(); //Roughness (Joint Roughness Coefficient)
            INFL14(); //Infilling of the discontinuities
            SW14(); // discontinuity Surface weathering
            ALT14(); // Alterability
            DAR14(); // Rating for discontinuity orientation adjustment
            FE(); //  Adjustment Rating for Exacavation method
            FS(); //  Adjustment Rating forStress-Strain behaviour along tunnel face
            RMR14(); //Basic RMR and RMR

            update14(); 
        }

        // RMR89 UPDATE (CALLING ALL PARAMETERS' FUNCTIONS)

        private void button4_Click(object sender, EventArgs e)
        {
            rmr = 0; brmr = 0; rucs = 0; rplsi = 0; rrqd = 0; rspd = 0; rgwi = 0; rjwp = 0; rgw = 0; rdl = 0; rda = 0; rjrc = 0; rhi = 0; rsi = 0; rni = 0; rsw = 0; rdc = 0; dar = 0;
            errorProvider1.Clear();
            clearlistbox();

            STRENGTH();  // Strength of Intact rock
            RQD();  //Rock qualty Designation
            SPD(); //Spacing of Discontinuities
            GWC(); //Ground water condition
            Parameters.Items.Add(" ");
            Parameters.Items.Add("Joint Condition: ");
            if (radioButton6.Checked)
            {
                DL(); //Discontinuity length (persistence)
                DA(); //Discontinuity Aperture
                JRC(); //Roughness (Joint Roughness Coefficient)
                INFL(); //Infilling of the discontinuities
                SW(); // discontinuity Surface weathering
            }
            else if (radioButton5.Checked)
            {
                DC();//General Discontinuity conditions
            }
            DAR(); // Rating for discontinuity adjustment
            RMR(); //Basic RMR and RMR

            update();
        }

        private void button5_Click(object sender, EventArgs e) // CLEAR RECORD button
        {
            clearrecord();
        }

        private void button6_Click(object sender, EventArgs e) // CLEAR button
        {
            errorProvider1.Clear();
            ClearTextBoxes(this);
            clearlistbox();
            numericUpDown13.Visible = false;
            label21.Visible = false;
            label17.Text = "...";
        }

        private void button7_Click(object sender, EventArgs e) // EXPORT button
        {
            export();
        }

        private void comboBox12_SelectedIndexChanged(object sender, EventArgs e)
        {
            string optioniii = comboBox12.Text;
            switch (optioniii)
            {
                case "None":
                    numericUpDown25.Visible = false;
                    label47.Visible = false;
                    break;
                case "Hard":
                    numericUpDown25.Visible = true;
                    label47.Visible = true;
                    break;
                case "Soft":
                    numericUpDown25.Visible = true;
                    label47.Visible = true;
                    break;
            }
        }

        // Opening pdf from resources
        private void tutorialToolStripMenuItem_Click(object sender, EventArgs e)
        {
            String openPDFFile = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\User_guide.pdf";//PDF DOc name
            System.IO.File.WriteAllBytes(openPDFFile, global::QuickRMR.Properties.Resources.User_Guide);//the resource automatically creates            
            System.Diagnostics.Process.Start(openPDFFile);   
        }

        // Opens about form
        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmAbout frm = new frmAbout();
            frm.Show();
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            string optioniii = comboBox5.Text;
            switch (optioniii)
            {
                case "None":
                    numericUpDown13.Visible = false;
                    label21.Visible = false;
                    break;
                case "Hard":
                    numericUpDown13.Visible = true;
                    label21.Visible = true;
                    break;
                case "Soft":
                    numericUpDown13.Visible = true;
                    label21.Visible = true;
                    break;
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            errorProvider1.Clear();
            ClearTextBoxes(this);
            clearlistbox();
            label48.Text = "...";
        }

        private void button11_Click(object sender, EventArgs e)
        {
            clearrecord();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            export();
        }

        // RMR89 STRENGTH function

        double STRENGTH()
        {
            string option1 = comboBox1.Text;    //defining variable for combobox string.

            if (radioButton1.Checked) // When UCS/PLSI radiobutton actve
            {
                    if (numericUpDown1.Text == "")
                    {
                        numericUpDown1.Text = "0";
                    }
                    if (option1 == "Select")
                    {
                        MessageBox.Show("You have not selected any of the Strength options.\nRating has been ignored.", "Attention!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }

                    switch (option1)
                    {
                        //Uniaxial Compressive Strength (UCS)
                        case "UCS":
                            ucs = Convert.ToDouble(numericUpDown1.Text);
                            rucs = 0.42 * Math.Pow(Math.Abs(ucs - 0.01), 0.65) - 0.1;
                            if (rucs < 0.1)
                                rucs = 0;
                            else if (rucs > 15)
                                rucs = 15;
                            rucs = Math.Round(rucs, 1);
                            Parameters.Items.Add("UCS               " + rucs);
                            break;

                        //Point load Strength Index (PLSI)

                        case "PLSI":
                            plsi = Convert.ToDouble(numericUpDown1.Text);

                            if (plsi < 1)
                            {
                                MessageBox.Show("PLSI value is too low to be considered. UCS value is Prefered.", "Attention!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                rplsi = -13 + (28 / (1 + Math.Pow(10, (0.02 - 0.15 * plsi))));
                                if (rplsi > 15)
                                    rplsi = 15;
                                rplsi = Math.Round(rplsi, 1);
                                Parameters.Items.Add("PLSI              " + rplsi);
                            }
                            break;
                    }
                
            }

            else if (radioButton2.Checked) // When R value Radio button is active
            {
                if (numericUpDown2.Text == "")
                {
                    numericUpDown2.Text = "0";
                }
                if (numericUpDown3.Text == "")
                {
                    numericUpDown3.Text = "0";
                }
                double rn = Convert.ToDouble(numericUpDown2.Text);
                    double rl = (rn - 6.3673) / 1.0646;
                    double rho = Convert.ToDouble(numericUpDown3.Text);
                    ucs = 6.9 * Math.Pow(10, (0.0087 * rl * rho) + 0.16);
                    rucs = 0.42 * Math.Pow(Math.Abs(ucs - 0.01), 0.65) - 0.1;
                    if (rucs < 0.1)
                        rucs = 0;
                    else if (rucs > 15)
                        rucs = 15;
                    rucs = Math.Round(rucs, 1);
                    Parameters.Items.Add("UCS               " + rucs);            
            }
            return rucs;
        }

        //RMR14 STRENGTH function

        double STRENGTH14()
        {
            string option1 = comboBox8.Text; //defining variable for combobox string.

            if (radioButton8.Checked) // When UCS/PLSI radiobutton active
            {
                    if (numericUpDown16.Text == "")
                    {
                        numericUpDown16.Text = "0";
                    }
                    if (option1 == "Select")
                    {
                        MessageBox.Show("You have not selected any of the Strength options.\nRating has been ignored.", "Attention!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }

                    switch (option1)
                    {
                        //Uniaxial Compressive Strength (UCS)
                        case "UCS":
                            ucs = Convert.ToDouble(numericUpDown16.Text);
                            rucs = 0.42 * Math.Pow(Math.Abs(ucs - 0.01), 0.65) - 0.1;
                            if (rucs < 0.1)
                                rucs = 0;
                            else if (rucs > 15)
                                rucs = 15;
                            rucs = Math.Round(rucs, 1);
                            listBox1.Items.Add("UCS               " + rucs);
                            break;

                        //Point load Strength Index (PLSI)

                        case "PLSI":
                            plsi = Convert.ToDouble(numericUpDown16.Text);

                            if (plsi < 1)
                            {
                                MessageBox.Show("PLSI value is too low to be considered. UCS value is Prefered.", "Attention!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                rplsi = -13 + (28 / (1 + Math.Pow(10, (0.02 - 0.15 * plsi))));
                                if (rplsi > 15)
                                    rplsi = 15;
                                rplsi = Math.Round(rplsi, 1);
                                listBox1.Items.Add("PLSI              " + rplsi);
                            }
                            break;
                    }
                
            }

            else if (radioButton7.Checked) // When R value Radio button is active
            {
                    if (numericUpDown17.Text == "")
                    {
                        numericUpDown17.Text = "0";
                    }
                    if (numericUpDown18.Text == "")
                    {
                        numericUpDown18.Text = "0";
                    }
                double rn = Convert.ToDouble(numericUpDown17.Text);
                    double rl = (rn - 6.3673) / 1.0646;
                    double rho = Convert.ToDouble(numericUpDown18.Text);
                    ucs = 6.9 * Math.Pow(10,(0.0087*rl*rho)+0.16);
                    rucs = 0.42 * Math.Pow(Math.Abs(ucs - 0.01), 0.65) - 0.1;
                    if (rucs < 0.1)
                        rucs = 0;
                    else if (rucs > 15)
                        rucs = 15;
                    rucs = Math.Round(rucs, 1);
                    listBox1.Items.Add("UCS               " + rucs);
            }
            return rucs;
        }

        // ROCK QUALITY DESIGNATION function

        double RQD()
        {
                if (numericUpDown4.Text == "")
                {
                    numericUpDown4.Text = "0";
                }
            
                rqd = Convert.ToDouble(numericUpDown4.Text);

                if (rqd > 100)
                {
                    MessageBox.Show("RQD can not be more than 100.\nEnter a value between 0 and 100.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                else
                {
                    rrqd = 0.317 * Math.Pow(Math.Abs(rqd - 0.07), 0.9) - 0.02;
                    if (rrqd > 19.95)
                        rrqd = 20;
                    else if (rrqd < 0.1)
                        rrqd = 0;
                    rrqd = Math.Round(rrqd, 1);
                    Parameters.Items.Add("RQD               " + rrqd);
                }
            return rrqd;
        }

        // DISCONTINUITY SPACING function

        double SPD()
        {
            if (numericUpDown5.Text == "")
            {
                numericUpDown5.Text = "0";
            }

            spd = Convert.ToDouble(numericUpDown5.Text);
                    rspd = 1.94 * Math.Pow(spd, 0.316) - 1.5;
                    if (rspd > 19.92)
                        rspd = 20;
                    else if (rspd < 0)
                        rspd = 0;
                    rspd = Math.Round(rspd, 1);
                    Parameters.Items.Add("Joint Spacing     " + rspd);
            return (rspd);
        }

        double DD14()
        {
            if (numericUpDown19.Text == "")
            {
                numericUpDown19.Text = "0";
            }
            double dd = Convert.ToDouble(numericUpDown19.Text);
                    rdd = 42.35-7.4*Math.Pow(dd,0.442);
                    if (rdd > 39.68)
                        rdd = 40;
                    else if (rdd < 0.65)
                        rdd = 0;
                    rdd = Math.Round(rdd, 1);
                    listBox1.Items.Add("No. of joints/m     " + rdd);
            return rdd;
        }

        // RMR89 GROUND WATER CONDITION function

        double GWC()
        {
            string option2 = comboBox2.Text;
            string optioni = comboBox3.Text;

            if (radioButton4.Checked) //FOR MEASURED Ground water option
            {
                    if (numericUpDown6.Text == "")
                    {
                        numericUpDown6.Text = "0";
                    }

                    if (option2 == "Select")
                    {
                        MessageBox.Show("You have not selected any of the Ground Water options.\nRating has been ignored.", "Attention!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }

                    switch (option2)
                    {
                        //Ground Water Inflow                         

                        case "GW Inflow":
                            gwi = Convert.ToDouble(numericUpDown6.Text);

                            rgwi = 165 / (gwi + 11);
                            if (rgwi < 0.1)
                                rgwi = 0;
                            rgwi = Math.Round(rgwi, 1);
                            Parameters.Items.Add("GW Inflow         " + rgwi);
                            break;

                        //Joint water pressure

                        case "JW Pressure":
                            jwp = Convert.ToDouble(numericUpDown6.Text);
                            rjwp = 2.4 / (jwp + 0.15) - 1;
                            if (rjwp < 0)
                                rjwp = 0;
                            rjwp = Math.Round(rjwp, 1);
                            Parameters.Items.Add("JW Pressure       " + rjwp);
                            break;
                    }
                                   
            }
            else if (radioButton3.Checked) //FOR GENERAL Ground water option
            {
                if (numericUpDown7.Text == "")
                {
                    numericUpDown7.Text = "0";
                }

                if (optioni == "select")
                {
                    rgw = Convert.ToDouble(numericUpDown7.Text);
                    Parameters.Items.Add("Ground Water       " + rgw);
                }

                else
                {
                    switch (optioni)
                    {
                         
                        case "Completely dry":
                            rgw = 15;
                            Parameters.Items.Add("Ground Water       " + rgw);
                            break;
                        case "Damp":
                            rgw = 10;
                            Parameters.Items.Add("Ground Water       " + rgw);
                            break;
                        case "Wet":
                            rgw = 7;
                            Parameters.Items.Add("Ground Water       " + rgw);
                            break;
                        case "Dripping":
                            rgw = 4;
                            Parameters.Items.Add("Ground Water       " + rgw);
                            break;
                        case "Flowing":
                            rgw = 0;
                            Parameters.Items.Add("Ground Water       " + rgw);
                            break;
                    }
                }
            }
            return (rjwp);
        }

        // RMR14 GROUND WATER CONDITION function

        double GWC14()
        {
            string option2 = comboBox9.Text;
            string optioni = comboBox10.Text;

            if (radioButton10.Checked) //FOR MEASURED Ground water option
            {
                    if (numericUpDown20.Text == "")
                    {
                        numericUpDown20.Text = "0";
                    }
                    if (option2 == "Select")
                    {
                        MessageBox.Show("You have not selected any of the Ground Water options.\nRating has been ignored.", "Attention!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }

                    switch (option2)
                    {
                        //Ground Water Inflow

                        case "GW Inflow":
                            gwi = Convert.ToDouble(numericUpDown20.Text);
                            rgwi = 165 / (gwi + 11);
                            if (rgwi < 0.1)
                                rgwi = 0;
                            rgwi = Math.Round(rgwi, 1);
                            listBox1.Items.Add("GW Inflow         " + rgwi);
                            break;
                        
                        //Joint water pressure

                        case "JW Pressure":
                            jwp = Convert.ToDouble(numericUpDown20.Text);
                            rjwp = 2.4 / (jwp + 0.15) - 1;
                            if (rjwp < 0)
                                rjwp = 0;
                            rjwp = Math.Round(rjwp, 1);
                            listBox1.Items.Add("JW Pressure       " + rjwp);
                            break;
                    }    
                
            }

            else if (radioButton9.Checked) //FOR GENERAL Ground water option
            {
                if (numericUpDown21.Text == "")
                {
                    numericUpDown21.Text = "0";
                }

                if (optioni == "Select")
                {
                    rgw = Convert.ToDouble(numericUpDown21.Text);

                    listBox1.Items.Add("Ground Water       " + rgw);
                }

                else
                {
                    switch (optioni)
                    {

                        case "Completely dry":
                            rgw = 15;
                            listBox1.Items.Add("Ground Water       " + rgw);
                            break;
                        case "Damp":
                            rgw = 10;
                            listBox1.Items.Add("Ground Water       " + rgw);
                            break;
                        case "Wet":
                            rgw = 7;
                            listBox1.Items.Add("Ground Water       " + rgw);
                            break;
                        case "Dripping":
                            rgw = 4;
                            listBox1.Items.Add("Ground Water       " + rgw);
                            break;
                        case "Flowing":
                            rgw = 0;
                            listBox1.Items.Add("Ground Water       " + rgw);
                            break;
                    }
                }
            }
            return (rjwp);
        }

        // RMR89 ADJUSTMENT RATING FOR DISCONTINUITIES FUNCTION (Discontinuity Adjustment Rating)

        public double DAR()
        {
            if (numericUpDown8.Text == "")
            {
                numericUpDown8.Text = "0";
            }
            string option4 = comboBox6.Text;
            string option5 = comboBox7.Text;

            if (option4 == "Select" || option5 == "Select")
            {
                dar = Convert.ToDouble(numericUpDown8.Text);               
            }

            else
            {
                switch (option5)
                {
                    case "Very Favourable":
                        if (option4 == "Tunnels & Mines" || option4 == "Foundations" || option4 == "Slopes")
                        {
                            dar = 0;
                        }
                        break;

                    case "Favourable":
                        if (option4 == "Tunnels & Mines" || option4 == "Foundations")
                        {
                            dar = -2;
                        }
                        else if (option4 == "Slopes")
                        {
                            dar = -5;
                        }
                        break;

                    case "fair":

                        if (option4 == "Tunnels & Mines")
                        {
                            dar = -5;
                        }
                        else if (option4 == "Foundations")
                        {
                            dar = -7;
                        }
                        if (option4 == "Slopes")
                        {
                            dar = -25;
                        }
                        break;
                    case "Unfavourable":
                        if (option4 == "Tunnels & Mines")
                        {
                            dar = -10;
                        }
                        else if (option4 == "Foundations")
                        {
                            dar = -15;
                        }
                        if (option4 == "Slopes")
                        {
                            dar = -50;
                        }
                        break;

                    case "Very Unfavourable":
                        if (option4 == "Tunnels & Mines")
                        {
                            dar = -12;
                        }
                        else if (option4 == "Foundations")
                        {
                            dar = -25;
                        }
                        if (option4 == "Slopes")
                        {
                            dar = -50;
                        }
                        break;
                }
            }
            Parameters.Items.Add("Adjustment Rating " + dar);
            return dar;
        }

        // RMR14 ADJUSTMENT RATING FOR DISCONTINUITIES FUNCTION

        public double DAR14()
        {
            if (numericUpDown22.Text == "")
            {
                numericUpDown22.Text = "0";
            }
            string option4 = comboBox11.Text;
            string option5 = comboBox13.Text;

            if (option4 == "Select" || option5 == "Select")
            {
                dar = Convert.ToDouble(numericUpDown22.Text);
            }

            else
            {
                switch (option5)
                {
                    case "Very Favourable":
                        if (option4 == "Tunnels & Mines" || option4 == "Foundations" || option4 == "Slopes")
                        {
                            dar = 0;
                        }
                        break;

                    case "Favourable":
                        if (option4 == "Tunnels & Mines" || option4 == "Foundations")
                        {
                            dar = -2;
                        }
                        else if (option4 == "Slopes")
                        {
                            dar = -5;
                        }
                        break;

                    case "fair":

                        if (option4 == "Tunnels & Mines")
                        {
                            dar = -5;
                        }
                        else if (option4 == "Foundations")
                        {
                            dar = -7;
                        }
                        if (option4 == "Slopes")
                        {
                            dar = -25;
                        }
                        break;
                    case "Unfavourable":
                        if (option4 == "Tunnels & Mines")
                        {
                            dar = -10;
                        }
                        else if (option4 == "Foundations")
                        {
                            dar = -15;
                        }
                        if (option4 == "Slopes")
                        {
                            dar = -50;
                        }
                        break;

                    case "Very Unfavourable":
                        if (option4 == "Tunnels & Mines")
                        {
                            dar = -12;
                        }
                        else if (option4 == "Foundations")
                        {
                            dar = -25;
                        }
                        if (option4 == "Slopes")
                        {
                            dar = -50;
                        }
                        break;
                }
            }
            listBox1.Items.Add("Adjustment Rating " + dar);
            return dar;           
        }

        //condition of Discontinuities

        //RMR89 Discontinuity length (persistence)

        double DL()
        {

            if (numericUpDown10.Text == "")
            {
                numericUpDown10.Text = "0";
            }

            dl = Convert.ToDouble(numericUpDown10.Text);
               
                {
                    rdl = 1 / (0.166 + dl * 0.05);
                    if (rdl < 0.1)
                        rdl = 0;
                    else if (rdl > 6)
                        rdl = 6;
                    rdl = Math.Round(rdl, 1);
                    Parameters.Items.Add("Persistence       " + rdl);
                }
            return rdl;
        }

        //RMR14 Discontinuity length (persistence)

        double DL14()
        {
            if (numericUpDown23.Text == "")
            {
                numericUpDown23.Text = "0";
            }
            dl = Convert.ToDouble(numericUpDown23.Text);
                    rdl = 5-0.5*dl;
                    if (rdl < 0.1)
                        rdl = 0;
                    else if (rdl > 5)
                        rdl = 5;
                    rdl = Math.Round(rdl, 1);
                    listBox1.Items.Add("Persistence       " + rdl);
            return rdl;
        }

        //Discontinuity Aperture

        double DA()
        {
            if (numericUpDown11.Text == "")
            {
                numericUpDown11.Text = "0";
            }
            da = Convert.ToDouble(numericUpDown11.Text);
                rda = 1 / (0.166 + da * 0.25);
                if (rda < 0.1)
                    rda = 0;
                else if (rda > 6)
                    rda = 6;
                rda = Math.Round(rda, 1);
                Parameters.Items.Add("Aperture          " + rda);

            return rda;
        }

        //RMR89 Roughness (Joint Roughness Coefficient)

        double JRC()
        {
            if (numericUpDown12.Text == "")
            {
                numericUpDown12.Text = "0";
            }
            jrc = Convert.ToDouble(numericUpDown12.Text);
                if (jrc > 20)
                {
                    MessageBox.Show("JRC can not be more than 20.\nEnter a value between 0 and 20", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                else
                {
                    rjrc = 0.3 * jrc;
                    rjrc = Math.Round(rjrc, 1);
                    Parameters.Items.Add("Roughness         " + rjrc);
                }
            return rjrc;
        }

        //RMR14 Roughness (Joint Roughness Coefficient)

        double JRC14()
        {
            if (numericUpDown24.Text == "")
            {
                numericUpDown24.Text = "0";
            }
            jrc = Convert.ToDouble(numericUpDown24.Text);
                if (jrc > 20)
                {
                    MessageBox.Show("JRC can not be more than 20.\nEnter a value between 0 and 20", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                else
                {
                    rjrc = 0.25 * jrc;
                    rjrc = Math.Round(rjrc, 1);
                    listBox1.Items.Add("Roughness         " + rjrc);
                }
            return rjrc;
        }

        //RMR89 Infilling of the discontinuities

        double INFL()
        {
            if (numericUpDown13.Text == "")
            {
                numericUpDown13.Text = "0";
            }
            string option3 = comboBox5.Text;

            if (option3 == "Select")
                    MessageBox.Show("You have not selected any of the Infilling options.\nRating has been ignored.", "Attention!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            else
            {
                switch (option3)
                {
                    //Hard Infilling

                    case "Hard":
                            hi = Convert.ToDouble(numericUpDown13.Text);
                            rhi = 6 - 4 * Math.Pow(hi, 4) / (256 + Math.Pow(hi, 4));
                            if (rhi < 2)
                                rhi = 2;
                            else if (rhi > 6)
                                rhi = 6;
                            rhi = Math.Round(rhi, 1);
                            Parameters.Items.Add("Hard Infilling    " + rhi);                       
                        break;

                    //Soft Infilling

                    case "Soft":

                            si = Convert.ToDouble(numericUpDown13.Text);
                            rsi = 162 / (27 + Math.Pow(si, 3));
                            if (rsi < 0)
                                rsi = 0;
                            else if (rsi > 5.98)
                                rsi = 6;
                            rsi = Math.Round(rsi, 1);
                            Parameters.Items.Add("Soft Infilling    " + rsi);

                        break;
                    case "None":
                        rni = 5;
                        Parameters.Items.Add("Infilling None    " + rni);
                        break;
                }

            }

            return rsi;
        }

        //RMR14 Infilling of the discontinuities

        double INFL14()
        {
            if (numericUpDown25.Text == "")
            {
                numericUpDown25.Text = "0";
            }
            string option3 = comboBox12.Text;

            if (option3 == "Select")
             
                        MessageBox.Show("You have not selected any of the Infilling options.\nRating has been ignored.", "Attention!", MessageBoxButtons.OK, MessageBoxIcon.Information);
              else
              { 
                    switch (option3)
                    {
                        //Hard Infilling

                        case "Hard":

                                hi = Convert.ToDouble(numericUpDown25.Text);
                                rhi = 5 - 3 * Math.Pow(hi, 2) / (10 + Math.Pow(hi, 2));
                                if (rhi < 2)
                                    rhi = 2;
                                else if (rhi > 5)
                                    rhi = 5;
                                rhi = Math.Round(rhi, 1);
                                listBox1.Items.Add("Hard Infilling    " + rhi);                            
                           
                        break;

                    //Soft Infilling

                        case "Soft":

                            si = Convert.ToDouble(numericUpDown25.Text);
                            rsi = 1 / (0.2 + 0.2 * si);
                            if (rsi < 0)
                                rsi = 0;
                            else if (rsi > 4.98)
                                rsi = 5;
                            rsi = Math.Round(rsi, 1);
                            listBox1.Items.Add("Soft Infilling    " + rsi);

                        break;

                        case "None":
                            rni = 5;
                            listBox1.Items.Add("Infilling None    " + rni);
                        break;
                    }
              }
            return rsi;
        }

        //RMR89 Joint surface weathering.

        double SW()
        {
            if (numericUpDown14.Text == "")
            {
                numericUpDown14.Text = "10";
            }
            if (numericUpDown15.Text == "")
            {
                numericUpDown15.Text = "10";
            }

            Ru5 = Convert.ToDouble(numericUpDown14.Text);
            Rw1 = Convert.ToDouble(numericUpDown15.Text);
                if (Rw1>Ru5)
                {
                    MessageBox.Show("Fresh surface R value (Ru5) must be >= weathered surface R value (Ru1).\nCorrect the value and calculate again to include surface weathering rating.", "Attention!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    sw = 100 * (Ru5 - Rw1) / Ru5;
                    rsw = 6.3 / (1 + Math.Pow(10, (0.032 * sw - 1.3)));
                    if (rsw < 0)
                        rsw = 0;
                    else if (rsw > 5.99)
                        rsw = 6;
                    rsw = Math.Round(rsw, 1);
                    Parameters.Items.Add("Weathering        " + rsw);
                }
            return rsw;
        }

        //RMR89 Joint surface weathering.

        double SW14()
        {
            if (numericUpDown26.Text == "")
            {
                numericUpDown26.Text = "10";
            }
            if (numericUpDown27.Text == "")
            {
                numericUpDown27.Text = "10";
            }
            Ru5 = Convert.ToDouble(numericUpDown26.Text);
            Rw1 = Convert.ToDouble(numericUpDown27.Text);
                if (Rw1 > Ru5)
                {
                    MessageBox.Show("Fresh surface R value (Ru5) must be >= weathered surface R value (Ru1).\nCorrect the value and calculate again to include surface weathering rating.", "Attention!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    sw = 100 * (Ru5 - Rw1) / Ru5;
                    rsw = 5.3 / (1 + Math.Pow(10, (0.03 * sw - 1.22)));
                    if (rsw < 0)
                        rsw = 0;
                    else if (rsw > 4.99)
                        rsw = 5;
                    rsw = Math.Round(rsw, 1);
                    listBox1.Items.Add("Weathering        " + rsw);
                }
            
            return rsw;
        }

        //GENERAL DISCCONTINUITY CONDITION FUNCTION

        double DC()
        {
            if (numericUpDown9.Text == "")
            {
                numericUpDown9.Text = "0";
            }

            string optionii = comboBox4.Text;
    
            if (optionii == "Select")
            {
                rdc = Convert.ToDouble(numericUpDown9.Text);
                Parameters.Items.Add("Joint Condition    " + rdc);                
            }

            else
            {
                switch (optionii)
                {
                    case "Excellent":
                        rdc = 30;
                        Parameters.Items.Add("Joint Condition    " + rdc);                        
                        break;
                    case "Good":
                        rdc = 25;
                        Parameters.Items.Add("discontinuity      " + rdc);
                        break;
                    case "Fair":
                        rdc = 20;
                        Parameters.Items.Add("Joint Condition    " + rdc);
                        break;
                    case "Bad":
                        rdc = 10;
                        Parameters.Items.Add("Joint Condition    " + rdc);
                        break;
                    case "Very Bad":
                        rdc = 0;
                        Parameters.Items.Add("Joint Condition    " + rdc);
                        break;
                }
            }
            return rdc;
        }

        //RMR14 ALTERABIITY

        double ALT14()
        {
            if (numericUpDown28.Text == "")
            {
                numericUpDown28.Text = "0";
            }

            double  Id2 = Convert.ToDouble(numericUpDown28.Text);

                    ralt = 14.3 / (1 + Math.Pow( 10, (0.78-0.015 * Id2))) - 2;
                    if (ralt < 0.1)
                        ralt = 0;
                    else if (ralt > 5)
                        ralt = 5;
                    ralt = Math.Round(ralt, 1);
                    listBox1.Items.Add("Persistence       " + ralt);
            
            return ralt;
        }

        double FE()
        {
            if (numericUpDown29.Text == "")
            {
                numericUpDown29.Text = "0";
            }
            Fe = Convert.ToDouble(numericUpDown29.Text);
            return Fe;
        }

        double FS()
        {
            if (numericUpDown30.Text == "")
            {
                numericUpDown30.Text = "0";
            }
            Fs = Convert.ToDouble(numericUpDown30.Text);
            return Fs;
        }

        //ROCK MASS RATING

        void RMR()
        {
                brmr = rucs + rplsi + rrqd + rspd + rgwi + rjwp + rgw + rdc + rdl + rda + rjrc + rhi + rsi + rsw;
                rmr = brmr + dar;
                brmr = Math.Round(brmr);
                rmr = Math.Round(rmr);
                if (brmr < 0)
                    brmr = 0;
                if (brmr > 100)
                    brmr = 100;
                if (rmr < 0)
                    rmr = 0;
                if (rmr > 100)
                    rmr = 100;
                label29.Text = Convert.ToString(brmr);
                label28.Text = Convert.ToString(rmr);
        }

        void RMR14()
        {
            brmr = rucs + rplsi + rdd + rgwi + rjwp + rgw + rdl + rjrc + rhi + rsi + rsw + ralt;
            rmr = (brmr + dar) * Fe * Fs;
            brmr = Math.Round(brmr);
            rmr = Math.Round(rmr);
            if (brmr < 0)
                brmr = 0;
            if (brmr > 100)
                brmr = 100;
            if (rmr < 0)
                rmr = 0;
            if (rmr > 100)
                rmr = 100;
            label49.Text = Convert.ToString(brmr);
            label51.Text = Convert.ToString(rmr);
        }

        void clearlistbox() // clears listbox
        {
            if (panel1.Visible)
            {
                Parameters.Items.Clear();
                Parameters.Items.Add("Parameters       Rating");
                Parameters.Items.Add("");
            }

            if (panel2.Visible)
            {
                listBox1.Items.Clear();
                listBox1.Items.Add("Parameters      Rating");
                listBox1.Items.Add("");
            }
        }


        void ClearTextBoxes(Control control) // Clears all text boxes and resets comboboxes
        {
            foreach (Control c in control.Controls)
            {
                if (c is TextBox)
                {
                    if (!(c.Parent is NumericUpDown))
                    {
                        ((TextBox)c).Clear();
                    }
                }

                if (c is ComboBox)
                {
                    ((ComboBox)c).SelectedIndex = 0;
                    ((ComboBox)c).Text = "Select";
                }

                if (c.HasChildren)
                {
                    ClearTextBoxes(c);
                }

                label17.Text = "...";
                label28.Text = "....";
                label29.Text = "....";
                label48.Text = "...";
            }
        }

        public void disp_data() // Display data in gridbox
        {
            if (panel1.Visible)
            {
                con.Open();
                SqlCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "select * from rmr89";
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                con.Close();
            }

            if (panel2.Visible)
            {
                con.Open();
                SqlCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "select * from rmr14";
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt);
                dataGridView2.DataSource = dt;
                con.Close();
            }
        }

       //RMR89 UPDATE

        void update()
        {
            DataGridViewRow newDataRow = dataGridView1.Rows[selectedRow];
            newDataRow.Cells[0].Value = textBox1.Text;
            newDataRow.Cells[1].Value = rucs;
            newDataRow.Cells[2].Value = rplsi;
            newDataRow.Cells[3].Value = rrqd;
            newDataRow.Cells[4].Value = rspd;
            newDataRow.Cells[5].Value = rgw;
            newDataRow.Cells[6].Value = rgwi;
            newDataRow.Cells[7].Value = rjwp;
            newDataRow.Cells[8].Value = rdc;
            newDataRow.Cells[9].Value = rdl;
            newDataRow.Cells[10].Value = rda;
            newDataRow.Cells[11].Value = rjrc;
            newDataRow.Cells[12].Value = rhi;
            newDataRow.Cells[13].Value = rsi;
            newDataRow.Cells[14].Value = rni;
            newDataRow.Cells[15].Value = rsw;
            newDataRow.Cells[16].Value = brmr;
            newDataRow.Cells[17].Value = dar;
            newDataRow.Cells[18].Value = rmr;                        
        }

        //RMR14 UPDATE

        void update14()
        {
            DataGridViewRow newDataRow = dataGridView2.Rows[selectedRow];
            newDataRow.Cells[0].Value = textBox16.Text;
            newDataRow.Cells[1].Value = rucs;
            newDataRow.Cells[2].Value = rplsi;
            newDataRow.Cells[3].Value = rdd;
            newDataRow.Cells[4].Value = rgw;
            newDataRow.Cells[5].Value = rgwi;
            newDataRow.Cells[6].Value = rjwp;
            newDataRow.Cells[7].Value = rdl;
            newDataRow.Cells[8].Value = rjrc;
            newDataRow.Cells[9].Value = rhi;
            newDataRow.Cells[10].Value = rsi;
            newDataRow.Cells[11].Value = rni;
            newDataRow.Cells[12].Value = rsw;
            newDataRow.Cells[13].Value = ralt;
            newDataRow.Cells[14].Value = brmr;
            newDataRow.Cells[15].Value = Fe;
            newDataRow.Cells[16].Value = Fs;
            newDataRow.Cells[17].Value = dar;
            newDataRow.Cells[18].Value = rmr;
        }

        // Deletes selected table record

        void clearrecord()
        {
            DialogResult dialog = MessageBox.Show("All Record will be deleted\n Do you want to Proceed?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (dialog == DialogResult.Yes)
            {
                con.Open();

                if (panel1.Visible)
                {
                    string sqlTrunc = "TRUNCATE TABLE " + "rmr89";
                    SqlCommand cmd = new SqlCommand(sqlTrunc, con);
                    cmd.CommandType = CommandType.Text;
                    cmd.ExecuteNonQuery();
                }

                if(panel2.Visible)
                {
                    string sqlTrunc = "TRUNCATE TABLE " + "rmr14";
                    SqlCommand cmd = new SqlCommand(sqlTrunc, con);
                    cmd.CommandType = CommandType.Text;
                    cmd.ExecuteNonQuery();
                }
                con.Close();
                disp_data();
            }
        }

        //Export

        void export()
        {
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
            worksheet = workbook.Sheets["Sheet1"];
            worksheet = workbook.ActiveSheet;
            worksheet.Name = "RMRDetail";

            if (panel1.Visible)
            {
                for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
                {
                    worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                }

                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                    }
                }

                var SavefileDialog = new SaveFileDialog();
                SavefileDialog.FileName = "RMR89_";
                SavefileDialog.DefaultExt = ".xlsx";

                if (SavefileDialog.ShowDialog() == DialogResult.OK)
                {
                    workbook.SaveAs(SavefileDialog.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    app.Quit();
                }
            }

            if (panel2.Visible)
            {
                for (int i = 1; i < dataGridView2.Columns.Count + 1; i++)
                {
                    worksheet.Cells[1, i] = dataGridView2.Columns[i - 1].HeaderText;
                }

                for (int i = 0; i < dataGridView2.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView2.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1] = dataGridView2.Rows[i].Cells[j].Value.ToString();
                    }
                }

                var SavefileDialog = new SaveFileDialog();
                SavefileDialog.FileName = "RMR14_";
                SavefileDialog.DefaultExt = ".xlsx";

                if (SavefileDialog.ShowDialog() == DialogResult.OK)
                {
                    workbook.SaveAs(SavefileDialog.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    app.Quit();
                }
            }
        }
    }
}
