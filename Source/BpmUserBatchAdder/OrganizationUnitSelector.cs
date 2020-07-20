using Dapper;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Windows.Forms;

namespace BpmUserBatchAdder {
    public partial class OrganizationUnitSelector : Form {
        public string selectedOrganizationUnitId;
        public List<dynamic> organizationUnitList;

        public OrganizationUnitSelector() {
            InitializeComponent();
        }

        private void OrganizationUnitSelector_Load(object sender, EventArgs e) {
            using (var connOrganizationUnitList = new SqlConnection(Database.BpmConnStr)) {
                connOrganizationUnitList.Open();

                organizationUnitList = connOrganizationUnitList.QueryAsync<dynamic>(
                #region SQL指令:部門代號清單
                @"
                SELECT OrganizationUnit.organizationUnitName, OrganizationUnit.id FROM OrganizationUnit WITH (NOLOCK)
                INNER JOIN Organization WITH (NOLOCK) ON Organization.OID = OrganizationUnit.organizationOID
                WHERE OrganizationUnit.validType = 1
                AND Organization.id IN (100000, 200000, 300000)
                AND OrganizationUnit.id NOT IN ('FC00', 'SC00')
                ORDER BY OrganizationUnit.id
                "
                #endregion
                , null, null, 600, null).Result.ToList();

                cbxOrganizationUnit.DisplayMember = "Text";
                cbxOrganizationUnit.ValueMember = "Value";

                cbxOrganizationUnit.Items.Clear();
                cbxOrganizationUnit.Items.Add(new { Text = "--請選擇--", Value = "" });
                foreach (var organizationUnit in organizationUnitList) {
                    cbxOrganizationUnit.Items.Add(new { Text = organizationUnit.organizationUnitName + "(" + organizationUnit.id + ")", Value = organizationUnit.id });
                }
                cbxOrganizationUnit.SelectedIndex = 0;
            }
        }

        private void cbxOrganizationUnit_KeyUp(object sender, KeyEventArgs e) {
            if (e.KeyCode == Keys.Enter) {
                if (cbxOrganizationUnit.SelectedItem == null) {
                    var searchText = (sender as ComboBox).Text;
                    var filteredOrganizationUnitList = organizationUnitList.Where(o => (o.organizationUnitName as string).ToLower().Contains(searchText.ToLower()) || (o.id as string).ToLower().Contains(searchText.ToLower()));

                    cbxOrganizationUnit.Items.Clear();
                    foreach (var organizationUnit in filteredOrganizationUnitList) {
                        cbxOrganizationUnit.Items.Add(new { Text = organizationUnit.organizationUnitName + "(" + organizationUnit.id + ")", Value = organizationUnit.id });
                    }
                    cbxOrganizationUnit.SelectionStart = cbxOrganizationUnit.Text.Length;
                }
                else {
                    selectedOrganizationUnitId = (cbxOrganizationUnit.SelectedItem as dynamic).Value;

                    this.Close();
                }
            }
        }

        protected override bool ProcessDialogKey(Keys keyData) {
            if (Form.ModifierKeys == Keys.None && keyData == Keys.Escape) {
                this.Close();

                return true;
            }

            return base.ProcessDialogKey(keyData);
        }
    }
}