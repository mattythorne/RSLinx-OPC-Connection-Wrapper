using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;



namespace RSLinxClient
{
    public partial class Form1 : Form
    {
       
        public RSLinxConnection RSLinx = new RSLinxConnection();

        public Form1()
        {
            InitializeComponent();
            
        }


        private void btnConnect_Click(object sender, EventArgs e)
        {
            RSLinx.connect();
        }

        private void btnDisconnect_Click(object sender, EventArgs e)
        {
            RSLinx.disconnect();
        }

        private void btnAddGroup_Click(object sender, EventArgs e)
        {
            // Add a new group to linx
            RSLinx.addGroup("MyNewGroup");
            
        }

        private void btnAddTags_Click(object sender, EventArgs e)
        {
            // Add tags to the group using the group name
            RSLinx.addTag("MyNewGroup","[SP3_FU_MTO_PL01_PEEL]One_Sec_ONS");
            RSLinx.addTag("MyNewGroup", "[SP3_FU_MTO_PL01_PEEL]EM001_STATION1_DipTank_Comp");
            // Add tag to the group using the group ID
            RSLinx.addTag(0, "[SP3_FU_MTO_PL01_PEEL]DateAndTime");
        }

        private void btnReadGroup_Click(object sender, EventArgs e)
        {
            RSLinxConnection.TagGroupValues tagValues = RSLinx.readAll("MyNewGroup");
            

            String entry;
            listBox1.Items.Clear();

            if (tagValues != null)
            {
                for (int i = 0; i < tagValues.getLength(); i++)
                {
                    entry = tagValues.tagNames.GetValue(i) + " " + tagValues.tagValues.GetValue(i);
                    listBox1.Items.Add(entry);
                }

            }
        }

        private void btnGetSingleTag_Click(object sender, EventArgs e)
        {
           MessageBox.Show(RSLinx.readTag("[SP3_FU_MTO_PL01_PEEL]EM001_STATION1_DipTank_Comp"));
            
        }

    }
}
