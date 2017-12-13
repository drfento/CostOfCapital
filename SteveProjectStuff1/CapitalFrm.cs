using System;
using System.Windows.Forms;

namespace CostOfCapital
{
    public partial class CapitalFrm : Form
    {
        GetDisplay getdisplay = new GetDisplay();
        CreateSCF createscf = new CreateSCF();

        //````````````````Initialize````````````````//
        public CapitalFrm()
        {
            InitializeComponent();
            getdisplay.AppInit(this);
        }

        //````````````````Buttons````````````````//
        private void GetCoCBtn_Click(object sender, EventArgs e)
        {
            getdisplay.DisplayCapitalView(this);
            getdisplay.DisplayFeedView(this);
            getdisplay.MergeFeedView(this);
        }

        private void GetFeedBtn_Click(object sender, EventArgs e)
        {
            if (this.dgvFeedView.RowCount > 0)
            { getdisplay.CreateFinalFeedView(this); }
            else
            { MessageBox.Show("No Data found on Feed tab."); }
        }

        private void CreateScfBtn_Click(object sender, EventArgs e)
        {
            if (this.dgvFinalView.RowCount > 0)
            {   createscf.GenerateFile(this); }
            else
            {   MessageBox.Show("No Data found on Final Feed tab."); }
        }

        //````````````````Events````````````````//
        private void DgvFeedView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            //Force dirty changes to commit
            if (this.dgvFeedView.IsCurrentCellDirty)
            { this.dgvFeedView.CommitEdit(DataGridViewDataErrorContexts.Commit); }
        }

    }
}
