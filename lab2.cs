using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.DirectoryServices;

namespace Lab2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
	
	    //Task 2	
	    //create two organizational units in Active Directory
        private void button1_Click(object sender, EventArgs e)
        {
            using (DirectoryEntry AD = new DirectoryEntry("LDAP://DC=HRYSHCHUK,DC=UA"))
            {
                using (DirectoryEntry u = AD.Children.Add("OU=Support", "organizationalUnit"))
                {
                    u.CommitChanges();
                    u.Children.Add("OU=Desktops", "organizationalUnit").CommitChanges();
                    u.Children.Add("OU=Laptops", "organizationalUnit").CommitChanges();
                    
                    MessageBox.Show("Task 2 completed");
                }
            }
        }


        //Task 3
    	//set password for user
        private static void SetPassword(DirectoryEntry UE, string password)
        {
            object[] oPassword = new object[] { password };
            object ret = UE.Invoke("SetPassword", oPassword);
            UE.CommitChanges();
        }


	    //Create user in Active Directory
        private void button2_Click(object sender, EventArgs e)
        {
            using (DirectoryEntry AD = new DirectoryEntry("LDAP://OU=Desktops,OU=Support,DC=HRYSHCHUK,DC=UA"))
            {
                string globalName = "user";
                string userName;
                for (int i = 0; i < 50; i++)
                {
                    userName =globalName+ i;
                    using (DirectoryEntry u = AD.Children.Add("CN="+userName, "user"))
                    {
                        u.Properties["displayName"].Add("User");
                        u.Properties["userPrincipalName"].Add(userName+"@HRYSHCHUK.UA");
                        u.Properties["sAMAccountName"].Add(userName);
                        u.Properties["department"].Add("Desktop office");
                        u.Properties["description"].Add("Very Good User");

                        u.Properties["businessCategory"].Add("Secretary");
                        u.Properties["businessCategory"].Add("Manager");
                        u.Properties["businessCategory"].Add("HelpDesk");  

                        u.CommitChanges();

                        SetPassword(u, "P@ssw0rd");
                        u.Properties["userAccountControl"].Value = "544";
                        u.CommitChanges();

                        //using (DirectoryEntry gdn = new DirectoryEntry("LDAP://CN=G1,OU=RR2,DC=sc,DC=univ,DC=net,DC=ua"))
                        //{
                        //    gdn.Properties["member"].Add(u.Properties["distinguishedName"].Value);
                        //    gdn.CommitChanges();
                        //}
                    }
                }
            }

            MessageBox.Show("Task 3 finished");

        }

	    //Task 4
	    //Add parameters for exiting users
        private void button3_Click(object sender, EventArgs e)
        {
            using (DirectoryEntry AD = new DirectoryEntry("LDAP://OU=Desktops,OU=Support,DC=HRYSHCHUK,DC=UA"))
            {
                
                string globalName = "user";
                string userName;
                for (int i = 0; i < 50; i++)
                {
                    userName ="CN="+ globalName + i;

                    using (DirectoryEntry u = AD.Children.Find(userName, "user"))
                    {
                        u.Properties["company"].Add("Kiev National University");
                        u.Properties["telephoneNumber"].Add("044-526052");
                        u.Properties["otherMobile"].Add("067-2334383");
                        u.Properties["otherMobile"].Add("050-2206789");
                        u.Properties["otherMobile"].Add("063-2184545");
                        u.CommitChanges();
                    }
                }

                MessageBox.Show("Task 4 finished");
            }
        }

        //Task 5
        //Create security group in Active Directory
        private void button4_Click(object sender, EventArgs e)
        {
            using (DirectoryEntry AD = new DirectoryEntry("LDAP://OU=Desktops,OU=Support,DC=HRYSHCHUK,DC=UA"))
            {
                AD.Children.Add("CN=AdvancedSpecialist", "Group").CommitChanges();
                using (DirectoryEntry u = AD.Children.Add("CN=AdvancedWorkers", "Group"))
                {
                    u.CommitChanges();
                    u.Properties["member"].Add("CN=AdvancedSpecialist,OU=Desktops,OU=Support,DC=HRYSHCHUK,DC=UA");
                    u.CommitChanges();
                }
            }
            MessageBox.Show("Task 5 finished");
        }

        //Task 6
        //Move users to another folder in AD  
        private void button5_Click(object sender, EventArgs e)
        {
            using (DirectoryEntry AD = new DirectoryEntry("LDAP://OU=Desktops,OU=Support,DC=HRYSHCHUK,DC=UA"))
            {

                string globalName = "user";
                string userName;
                for (int i = 0; i < 50; i++)
                {
                    userName = "CN=" + globalName + i;

                    using (DirectoryEntry u = AD.Children.Find(userName, "user"))
                    {

                        u.MoveTo(new DirectoryEntry("LDAP://OU=Laptops,OU=Support,DC=HRYSHCHUK,DC=UA"), userName);
                        u.CommitChanges();
                    }

                    using (DirectoryEntry gdn = new DirectoryEntry("LDAP://CN=AdvancedSpecialist,OU=Desktops,OU=Support,DC=HRYSHCHUK,DC=UA")) 
                    {
                        gdn.Properties["member"].Add(userName+",OU=Laptops,OU=Support,DC=HRYSHCHUK,DC=UA");
                        gdn.CommitChanges();
                    }
                }

                MessageBox.Show("Task 6 finished");
            }
        }

        //Task 7
        //Get all users from security group
        private void button6_Click(object sender, EventArgs e)
        {
            richTextBox1.Text = "";
            using (DirectoryEntry G = new DirectoryEntry("LDAP://CN=AdvancedSpecialist,OU=Desktops,OU=Support,DC=HRYSHCHUK,DC=UA"))
            {
                foreach (Object vc2 in G.Properties["member"])
                {
                                 
                    using (DirectoryEntry AD = new DirectoryEntry("LDAP://"+vc2.ToString()))
                    {
                        if (AD.Properties["objectClass"].Contains("user"))
                            richTextBox1.Text += AD.Name.Substring(3) + "\t\t|\t" + vc2.ToString() + "\n";
                    }
                }
            }

        }

        //Task 8
        //remove user from security group
        private void button7_Click(object sender, EventArgs e)
        {
            using (DirectoryEntry gdn = new DirectoryEntry("LDAP://CN=AdvancedSpecialist,OU=Desktops,OU=Support,DC=HRYSHCHUK,DC=UA"))
            {
                gdn.Properties["member"].Remove("CN="+textBox1.Text + ",OU=Laptops,OU=Support,DC=HRYSHCHUK,DC=UA");
                gdn.Properties["member"].Remove("CN="+textBox2.Text + ",OU=Laptops,OU=Support,DC=HRYSHCHUK,DC=UA");
                
                gdn.CommitChanges();
            }

            MessageBox.Show("Task 8 finished");
        }


        //Task 9
        //Login user using credentials in Active Directory
        private void button8_Click(object sender, EventArgs e)
        {
            string message;
            if (IsADAuthenticated(textBox3.Text, textBox4.Text))
                message = textBox3.Text + " logined";
            else
                message = "Authentication for " + textBox3.Text + " failed";
            MessageBox.Show(message);
        }

        public bool IsADAuthenticated(string L, string P)    //  L - Login, P - Password
        {
            try
            {
                //string domainAndUsername = DomainShortName + "\\" + L;
                using (DirectoryEntry AD = new DirectoryEntry("LDAP://OU=Laptops,OU=Support,DC=HRYSHCHUK,DC=UA", L, P))
                {
                    DirectorySearcher S = new DirectorySearcher(AD);
                    S.Filter = "(SAMAccountName=" + L + ")"; S.PropertiesToLoad.Add("cn");
                    SearchResult R = S.FindOne();
                    if (R == null)
                    {
                        return false;
                    }
                    else
                    {
                        return true;
                    }
                }
            }
            catch
            {
                return false;
            }
        }



    }
}

