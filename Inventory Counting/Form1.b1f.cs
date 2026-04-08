// File: Form1.cs

using System;
using SAPbouiCOM.Framework;

namespace Inventory_Counting
{
    [FormAttribute("Inventory_Counting.Form1", "Form1.b1f")]
    class Form1 : UserFormBase
    {
        private SAPbouiCOM.Matrix Matrix0;
        private SAPbouiCOM.EditText txtDate;
        private SAPbouiCOM.EditText txtTime;

        public override void OnInitializeComponent()
        {
            Matrix0 = (SAPbouiCOM.Matrix)this.GetItem("Item_17").Specific;
            txtDate = (SAPbouiCOM.EditText)this.GetItem("Item_41").Specific;
            txtTime = (SAPbouiCOM.EditText)this.GetItem("Item_0").Specific;

            OnCustomInitialize();
        }

        public override void OnInitializeFormEvents()
        {
            Application.SBO_Application.ItemEvent += SBO_Application_ItemEvent;
        }

        private void OnCustomInitialize()
        {
            txtDate.Value = DateTime.Now.ToString("dd-MMMM-yyyy");
            txtTime.Value = DateTime.Now.ToString("HH:mm");

            SetupCountingType();
            EnsureMatrixDataTable();
            BindMatrixColumns();
            AddItemCFL();
            AddWarehouseCFL();

            NewTransaction();
        }

        // =========================
        // DATATABLE
        // =========================
        private void EnsureMatrixDataTable()
        {
            SAPbouiCOM.DataTable dt;

            try
            {
                dt = this.UIAPIRawForm.DataSources.DataTables.Item("DT_1");
            }
            catch
            {
                dt = this.UIAPIRawForm.DataSources.DataTables.Add("DT_1");
            }

            AddColumn(dt, "ItemCode");
            AddColumn(dt, "ItemDesc");
            AddColumn(dt, "Freeze");
            AddColumn(dt, "WhsCode");
            AddColumn(dt, "Counted");
            AddColumn(dt, "InWhs"); // 🔥 IMPORTANT
        }

        private void AddColumn(SAPbouiCOM.DataTable dt, string name)
        {
            for (int i = 0; i < dt.Columns.Count; i++)
                if (dt.Columns.Item(i).Name == name) return;

            dt.Columns.Add(name, SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
        }

        // =========================
        // MATRIX BIND
        // =========================
        private void BindMatrixColumns()
        {
            Matrix0.Columns.Item("ItemCode").DataBind.Bind("DT_1", "ItemCode");
            Matrix0.Columns.Item("ItemDesc").DataBind.Bind("DT_1", "ItemDesc");
            Matrix0.Columns.Item("Freeze").DataBind.Bind("DT_1", "Freeze");
            Matrix0.Columns.Item("WhsCode").DataBind.Bind("DT_1", "WhsCode");
            Matrix0.Columns.Item("Counted").DataBind.Bind("DT_1", "Counted");

            Matrix0.Columns.Item("InWhs").DataBind.Bind("DT_1", "InWhs"); // 🔥 FIX

            Matrix0.Columns.Item("ItemCode").Editable = true;
            Matrix0.Columns.Item("InWhs").Editable = false;
        }

        // =========================
        // NEW TRANSACTION
        // =========================
        private void NewTransaction()
        {
            var dt = this.UIAPIRawForm.DataSources.DataTables.Item("DT_1");

            dt.Rows.Clear();
            dt.Rows.Add();

            dt.SetValue("ItemCode", 0, "");
            dt.SetValue("ItemDesc", 0, "");
            dt.SetValue("Freeze", 0, "");
            dt.SetValue("WhsCode", 0, "");
            dt.SetValue("Counted", 0, "");
            dt.SetValue("InWhs", 0, "");

            Matrix0.LoadFromDataSource();
        }

        // =========================
        // CFL SETUP
        // =========================
        private void AddItemCFL()
        {
            var cfl = (SAPbouiCOM.ChooseFromListCreationParams)
                Application.SBO_Application.CreateObject(
                SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);

            cfl.ObjectType = "4";
            cfl.UniqueID = "CFL_Item";

            this.UIAPIRawForm.ChooseFromLists.Add(cfl);

            Matrix0.Columns.Item("ItemCode").ChooseFromListUID = "CFL_Item";
            Matrix0.Columns.Item("ItemCode").ChooseFromListAlias = "ItemCode";
        }

        private void AddWarehouseCFL()
        {
            var cfl = (SAPbouiCOM.ChooseFromListCreationParams)
                Application.SBO_Application.CreateObject(
                SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);

            cfl.ObjectType = "64";
            cfl.UniqueID = "CFL_Whs";

            this.UIAPIRawForm.ChooseFromLists.Add(cfl);

            Matrix0.Columns.Item("WhsCode").ChooseFromListUID = "CFL_Whs";
            Matrix0.Columns.Item("WhsCode").ChooseFromListAlias = "WhsCode";
        }

        // =========================
        // EVENTS
        // =========================
        private void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.FormTypeEx != "Inventory_Counting.Form1")
                return;

            if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST && !pVal.BeforeAction)
            {
                HandleCFL(pVal);
            }
        }

        // =========================
        // HANDLE CFL
        // =========================
        private void HandleCFL(SAPbouiCOM.ItemEvent pVal)
        {
            var cfl = (SAPbouiCOM.IChooseFromListEvent)pVal;

            if (cfl.SelectedObjects == null)
                return;

            Matrix0.FlushToDataSource();

            var dt = this.UIAPIRawForm.DataSources.DataTables.Item("DT_1");
            int row = pVal.Row - 1;

            // ITEM
            if (pVal.ColUID == "ItemCode")
            {
                string code = cfl.SelectedObjects.GetValue("ItemCode", 0).ToString();
                string name = cfl.SelectedObjects.GetValue("ItemName", 0).ToString();

                dt.SetValue("ItemCode", row, code);
                dt.SetValue("ItemDesc", row, name);

                if (row == dt.Rows.Count - 1)
                {
                    dt.Rows.Add();
                    int newRow = dt.Rows.Count - 1;

                    dt.SetValue("ItemCode", newRow, "");
                    dt.SetValue("ItemDesc", newRow, "");
                    dt.SetValue("Freeze", newRow, "");
                    dt.SetValue("WhsCode", newRow, "");
                    dt.SetValue("Counted", newRow, "");
                    dt.SetValue("InWhs", newRow, "");
                }
            }

            // WAREHOUSE
            if (pVal.ColUID == "WhsCode")
            {
                string whs = cfl.SelectedObjects.GetValue("WhsCode", 0).ToString();

                dt.SetValue("WhsCode", row, whs);

                string item = dt.GetValue("ItemCode", row).ToString();

                if (!string.IsNullOrEmpty(item))
                {
                    double onHand = GetOnHand(item, whs);
                    dt.SetValue("InWhs", row, onHand);
                }
            }

            Matrix0.LoadFromDataSource();
        }

        // =========================
        // GET STOCK
        // =========================
        private double GetOnHand(string itemCode, string whsCode)
        {
            var company = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();

            var rs = (SAPbobsCOM.Recordset)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            rs.DoQuery($@"
                SELECT OnHand 
                FROM OITW 
                WHERE ItemCode = '{itemCode}' 
                AND WhsCode = '{whsCode}'");

            return rs.EoF ? 0 : Convert.ToDouble(rs.Fields.Item("OnHand").Value);
        }

        // =========================
        // COMBO
        // =========================
        private void SetupCountingType()
        {
            var combo = (SAPbouiCOM.ComboBox)this.UIAPIRawForm.Items.Item("Item_3").Specific;

            combo.ValidValues.Add("S", "Single Counter");
            combo.ValidValues.Add("M", "Multiple Counters");

            combo.Select("S", SAPbouiCOM.BoSearchKey.psk_ByValue);
        }
    }
}