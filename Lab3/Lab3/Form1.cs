using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.DirectoryServices;
using System.IO;



namespace Lab3
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (DirectoryEntry OU = new DirectoryEntry("LDAP://DC=HRYSHCHUK,DC=UA"))
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    using (DirectoryEntry u = OU.Children.Add("MPModel=" + textBox1.Text, "MilitaryPlaneModel"))
                    {
                        u.Properties["MilitaryPlaneModel-Attribute05"].Add(textBox1.Text);
                        using (FileStream F = new FileStream(openFileDialog1.FileName, FileMode.Open, FileAccess.Read))
                        {
                            byte[] B = new byte[F.Length];
                            F.Read(B, 0, (int)F.Length);
                            u.Properties["MilitaryPlaneModel-Attribute10"].Clear();
                            u.Properties["MilitaryPlaneModel-Attribute10"].Value = B;
                        }
                        u.CommitChanges();


                        //pictureBox1.Image = Image.FromFile(@"C:\Work\WF1\WF1\Aspnet.jpg");
                        //pictureBox1.Image = V((byte[])u.Properties["MilitaryPlaneModel-Attribute10"].Value);

                        MessageBox.Show("OK!");
                    }
                }
            }

        }


        public static Bitmap V(byte[] k)
        {
            using (MemoryStream S = new MemoryStream())
            {
                S.Write(k, 0, (int)k.Length);
                Bitmap m = new Bitmap(S, false);
                return m;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();
            dataGridView1.Columns.Add("Name", "Name");
            //dataGridView1.Columns.Add("Picture", "Picture");
            DataGridViewImageColumn imageCol = new DataGridViewImageColumn();
            dataGridView1.Columns.Add(imageCol);
            imageCol.ImageLayout = DataGridViewImageCellLayout.Zoom;
            dataGridView1.RowTemplate.Height = 60;
            

            string name = "";
            using (DirectoryEntry G = new DirectoryEntry("LDAP://DC=HRYSHCHUK,DC=UA"))
            {
                DirectorySearcher searchObject = new DirectorySearcher(G);
                searchObject.SearchScope = SearchScope.Subtree;
                searchObject.Filter="(objectClass=MilitaryPlaneModel)";

                foreach (SearchResult result in searchObject.FindAll()) 
                {
                    
                    DirectoryEntry user = result.GetDirectoryEntry();
                    name= user.Properties["MilitaryPlaneModel-Attribute05"].Value.ToString();



                    Image image = V((byte[])user.Properties["MilitaryPlaneModel-Attribute10"].Value);
                    
                    
                    dataGridView1.Rows.Add(name, image);
                }

               // MessageBox.Show(str);
            }


        }

    
    }

}
