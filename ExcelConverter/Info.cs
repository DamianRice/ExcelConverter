using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExcelConverter
{
    public partial class Info : Form
    {
        public DataTable dt = new DataTable();
        private DataColumn dtPath = new DataColumn("ColumnPath");
        private DataColumn dtResult = new DataColumn("ColumnResult");
        private DataColumn dtReason = new DataColumn("ColumnReason");
        private BindingSource bs = new BindingSource();

        public Info()
        {
            InitializeComponent();
            dt.Columns.Add(dtPath);
            dt.Columns.Add(dtResult);
            dt.Columns.Add(dtReason);
            
        }

        private void exitBtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        public void Update(DataTable dt)
        {
            bs.DataSource = dt;
            this.ColumnPath.DataPropertyName = @"ColumnPath";
            this.ColumnReason.DataPropertyName = @"ColumnReason";
            this.ColumnResult.DataPropertyName = @"ColumnResult";
            this.dataGridView1.DataSource = bs;
            
        }

        public void GenerateInfo(string filePath, bool info, string msg=@"检查是否有重名文件")
        {
            var dr = dt.NewRow();
            dr.ItemArray = new object[] {filePath, info ? @"成功" : @"失败", info ? @"成功" : msg };
            dt.Rows.Add(dr);
            Update(this.dt);
        }

        public void GenerateInfo(Dictionary<string, bool> fileResult, bool info, string msg = @"检查是否有重名文件")
        {
            dt.Clear();
            foreach (KeyValuePair<string, bool> b in fileResult)
            {
                var dr = dt.NewRow();
                dr.ItemArray = new object[] { b.Key, b.Value ? @"成功" : @"失败", b.Value ? @"成功" : msg };
                dt.Rows.Add(dr);
            }
            Update(this.dt);
        }


    }
}
