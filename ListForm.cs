using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Reflection;
using Autodesk.Revit.Creation;

namespace PilotRevitAddin01
{
    public partial class ListForm : System.Windows.Forms.Form
    {
        Autodesk.Revit.DB.Document Doc;

        private DataTable datatable1 = new DataTable();
        private DataTable datatable2 = new DataTable();
        private Category selectCategory;
        private Element selectElement;
        private Parameter selectParameter;
        private DataTable changeValueDataTable = new DataTable();

        public ListForm()
        {
            InitializeComponent();
        }

        public ListForm(Autodesk.Revit.DB.Document doc)
        {
            InitializeComponent();
            Doc = doc;
            initDataGrid();
            SetComboBox_Categorys();
        }


        private void initDataGrid()
        {
            datatable1.Columns.Add("ElementID", typeof(string));
            datatable1.Columns.Add("Category 명", typeof(string));
            datatable1.Columns.Add("Family 명", typeof(string));
            datatable1.Columns.Add("타입(이름)명", typeof(string));
            datatable1.Columns.Add("Element", typeof(Element));
            //datatable.Rows.Add("test ID", "Category 2번", "Family",  "타입 11");
            dataGridView1.DataSource = datatable1;
            dataGridView1.Columns["element"].Visible = false;

            datatable2.Columns.Add("파라미터명", typeof(string));
            datatable2.Columns.Add("값", typeof(string));
            datatable2.Columns.Add("Parameter", typeof(Parameter));
            datatable2.Columns.Add("IsReadOnly", typeof(bool));
            datatable2.Columns.Add("StorageType", typeof(StorageType));

            dataGridView2.DataSource = datatable2;
            dataGridView2.Columns["Parameter"].Visible = false;
            dataGridView2.Columns["IsReadOnly"].Visible = false;
            dataGridView2.Columns["StorageType"].Visible = false;

            dataGridView2.Columns[0].ReadOnly = true;


            DataColumn[] dtkey = new DataColumn[1];
            dtkey[0] = new DataColumn("ParameterID", typeof(int));
            changeValueDataTable.Columns.Add(dtkey[0]);
            changeValueDataTable.Columns.Add("Parameter", typeof(Parameter));
            changeValueDataTable.Columns.Add("newValue", typeof(string));
            changeValueDataTable.Columns.Add("RowIndex", typeof(int));
            changeValueDataTable.PrimaryKey = dtkey;


            this.dataGridView1.DoubleBuffered(true);
            this.dataGridView2.DoubleBuffered(true);

        }
        private void SetComboBox_Categorys()
        {
            Settings documentSettings = Doc.Settings;
            Categories groups = documentSettings.Categories;
            List<KeyValuePair<int, string>> list = new List<KeyValuePair<int, string>>();
            foreach (Category category in groups)
            {
                //comboBox_Categorys.Items.Add(new { Text = category.Name, Value = category.Id });
                list.Add(new KeyValuePair<int, string>(category.Id.IntegerValue, category.Name + " : " + category.BuiltInCategory.ToString()));
            }
            List<KeyValuePair<int, string>> sortList = list.OrderBy(x => x.Value).ToList();

            comboBox_Categorys.DataSource = sortList;
            comboBox_Categorys.ValueMember = "Key";
            comboBox_Categorys.DisplayMember = "Value";
            comboBox_Categorys.SelectedIndex = 0;
        }

        private void comboBox_Categorys_SelectedIndexChanged(object sender, EventArgs e)
        {
            //string selectedItem = "";
            if (comboBox_Categorys.SelectedIndex >= 0)
            {
                int selectCateID = ((KeyValuePair<int, string>)comboBox_Categorys.SelectedItem).Key;
                GetElements(selectCateID);
            }
        }

        private void GetElements(int categoryId)
        {
            Settings documentSettings = Doc.Settings;
            Categories groups = documentSettings.Categories;
            selectCategory = groups.Cast<Category>().Where(x => (x.Id.IntegerValue == categoryId)).FirstOrDefault();
            
            FilteredElementCollector collector = new FilteredElementCollector(Doc);
            // 선택 카테고리의 Element 조회   WhereElementIsNotElementType - 요소 유형이 아닌 요소만 찾기
            IList<Element> elements = collector.OfCategory(selectCategory.BuiltInCategory).WhereElementIsNotElementType().ToElements();

            datatable1.Rows.Clear();
            foreach (Element element in elements)
            {
                string eleTypeName = "";
                
                //eleTypeName = element.GetType().Name;
                
                ElementType elemType01 = Doc.GetElement(element.GetTypeId()) as ElementType;
                if (elemType01 != null)
                {
                    eleTypeName = elemType01.FamilyName;
                }
                
                datatable1.Rows.Add(element.Id.ToString(), element.Category.Name, element.Name, eleTypeName, element);
            }
            dataGridView1.DataSource = datatable1;
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewRow row = dataGridView1.SelectedRows[0];
            string data = row.Cells[0].Value.ToString();
            //TaskDialog.Show("확인", $"data : {data} ");

            selectElement = ((Element)row.Cells[4].Value);
            //TaskDialog.Show("확인", $"data : {selectElement.Name} ");
            GetParameters(selectElement);
        }


        private void dataGridView2_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            selectParameter = (Parameter)datatable2.Rows[e.RowIndex][2];
            string changeValue = datatable2.Rows[e.RowIndex][1].ToString();

            DataRow selectRow = changeValueDataTable.Rows.Find(selectParameter.Id.IntegerValue);

            if (selectRow != null)
            {
                TaskDialog.Show($" {selectRow[0]} ", $" {selectRow[2]} >> {changeValue} ");
                selectRow[2] = changeValue;
            }
            else
            {
                changeValueDataTable.Rows.Add(selectParameter.Id.IntegerValue, selectParameter, changeValue, e.RowIndex);
            }
            dataGridView2.Rows[e.RowIndex].Cells[1].Style.BackColor = System.Drawing.Color.PowderBlue;
            //bool isSave = SetParameter(selectParameter, datatable2.Rows[e.RowIndex][1].ToString());
        }


        private void dataGridView2_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            /*
            if (e.RowIndex < 0) return;
            DataGridViewRow row = dataGridView2.Rows[e.RowIndex];

            
            //Parameter p = (Parameter)row.Cells[2].Value;

            if (dataGridView2.Rows[e.RowIndex].Cells["IsReadOnly"].Value.Equals(true))
            {
                row.Cells[0].Style.ForeColor = System.Drawing.Color.Red;
                row.Cells[1].Style.ForeColor = System.Drawing.Color.Red;
                row.Cells[1].ReadOnly = true;
            }
            
            if (dataGridView2.Rows[e.RowIndex].Cells["StorageType"].Value.Equals(((int)StorageType.ElementId)))
            {
                row.Cells[0].Style.ForeColor = System.Drawing.Color.Blue;
                row.Cells[1].Style.ForeColor = System.Drawing.Color.Blue;
                row.Cells[1].ReadOnly = true;
            }
            */

        }



        private void GetParameters(Element seleElement)
        {
            datatable2.Rows.Clear();
            foreach (Parameter param in seleElement.Parameters)
            {
                string[] arrPara = GetParameterInformation(param, Doc);
                datatable2.Rows.Add(arrPara[0], arrPara[1], param, param.IsReadOnly, param.StorageType);
            }
            dataGridView2.DataSource = datatable2;

            changeValueDataTable.Rows.Clear();

            for(int i = 0; i < datatable2.Rows.Count; i++)
            {
                if (dataGridView2.Rows[i].Cells["IsReadOnly"].Value.Equals(true))
                {
                    dataGridView2.Rows[i].Cells[0].Style.ForeColor = System.Drawing.Color.Red;
                    dataGridView2.Rows[i].Cells[1].Style.ForeColor = System.Drawing.Color.Red;
                    dataGridView2.Rows[i].Cells[1].ReadOnly = true;
                }
                if (dataGridView2.Rows[i].Cells["StorageType"].Value.Equals(((int)StorageType.ElementId)))
                {
                    dataGridView2.Rows[i].Cells[0].Style.ForeColor = System.Drawing.Color.Blue;
                    dataGridView2.Rows[i].Cells[1].Style.ForeColor = System.Drawing.Color.Blue;
                    dataGridView2.Rows[i].Cells[1].ReadOnly = true;
                }
            }
        }

        public bool SetParameter(Parameter parameter, string value)
        {
            bool result = false;
            using (Transaction t = new Transaction(Doc, "param"))
            {
                t.Start();
                try
                {
                    switch (parameter.StorageType)
                    {
                        case StorageType.Double:
                            result = parameter.SetValueString(value);
                            
                            break;
                        case StorageType.ElementId:

                            break;
                        case StorageType.Integer:
                            int int_val = 0;
                            if (SpecTypeId.Boolean.YesNo == parameter.Definition.GetDataType())
                            {
                                if (Int32.TryParse(value, out int_val))
                                {
                                    if (int_val <= 1)
                                    {
                                        result = parameter.Set(int_val);
                                    }
                                }
                            }
                            else
                            {
                                if (Int32.TryParse(value, out int_val))
                                {
                                    result = parameter.Set(int_val);
                                }
                            }
                            break;
                        case StorageType.String:
                            result = parameter.Set(value);
                            break;
                        default:
                            break;
                    }
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("error", ex.Message.ToString());
                }
                t.Commit();
            }
            return result;
        }


        private string[] GetParameterInformation(Parameter para, Autodesk.Revit.DB.Document document)
        {
            string[] arrResult = new string[2];

            string defName = para.Definition.Name + "\t : ";
            string defValue = string.Empty;

            switch (para.StorageType)
            {
                case StorageType.Double:
                    //covert the number into Metric
                    defValue = para.AsValueString();
                    break;
                case StorageType.ElementId:
                    //find out the name of the element
                    Autodesk.Revit.DB.ElementId id = para.AsElementId();
                    if (id.IntegerValue >= 0)
                    {
                        defValue = document.GetElement(id).Name;
                    }
                    else
                    {
                        defValue = id.IntegerValue.ToString();
                        BuiltInCategory builtIn;
                        if (Enum.TryParse(defValue, out builtIn))
                        {
                            if (Enum.IsDefined(typeof(BuiltInCategory), builtIn) | builtIn.ToString().Contains(","))
                            {
                                defValue = builtIn.ToString() + " (" + defValue + ")";
                            }
                        }
                    }
                    break;
                case StorageType.Integer:
                    if (SpecTypeId.Boolean.YesNo == para.Definition.GetDataType())
                    {
                        if (para.AsInteger() == 0)
                        {
                            defValue = "False";
                        }
                        else
                        {
                            defValue = "True";
                        }
                    }
                    else
                    {
                        defValue = para.AsInteger().ToString();
                    }
                    break;
                case StorageType.String:
                    defValue = para.AsString();
                    break;
                default:
                    defValue = "Unexposed parameter.";
                    break;
            }
            //var result = Tuple.Create<string, string>(defName, defValue);
            arrResult[0] = defName;
            arrResult[1] = defValue;
            return arrResult;
        }



        private void applicationEventsInfoWindows_Shown(object sender, EventArgs e)
        {
            // set window's display location
            //int left = Screen.PrimaryScreen.WorkingArea.Left + this.Width + 25;
            //int top = Screen.PrimaryScreen.WorkingArea.Top + this.Height + 50;
            int left = Screen.PrimaryScreen.WorkingArea.Width / 2 - this.Width / 2;
            int top = Screen.PrimaryScreen.WorkingArea.Height/ 2 - this.Height / 2;

            System.Drawing.Point windowLocation = new System.Drawing.Point(left, top);
            this.Location = windowLocation;
        }

        private void button_save_Click(object sender, EventArgs e)
        {
            int count = changeValueDataTable.Rows.Count;

            TaskDialog dialog = new TaskDialog("파라미터 저장");
            dialog.MainContent = $" {count} 개 파라미터 저장 시작 ";
            dialog.AllowCancellation = true;
            dialog.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No;
            TaskDialogResult result = dialog.Show();

            //DialogResult result1 = MessageBox.Show("파라미터 저장", $" {count} 개 파라미터 저장 시작 ", MessageBoxButtons.YesNo);
            //if (result1 == DialogResult.Yes)    

            if (result == TaskDialogResult.Yes)
            {
                int saveCount = 0;
                int notsaveCount = 0;
                for (int i = 0; i < count; i++)
                {
                    var param = (Parameter)changeValueDataTable.Rows[i][1];
                    string val = changeValueDataTable.Rows[i][2].ToString();

                    bool isSave = SetParameter(param, val);
                    if (isSave)
                    {
                        saveCount++;
                        int rowIndex = (int)changeValueDataTable.Rows[i][3];
                        dataGridView2.Rows[rowIndex].Cells[1].Style.BackColor = System.Drawing.Color.White;
                    }
                    else
                    {
                        notsaveCount++;
                    }
                }
                TaskDialog.Show("파라미터 저장", $" {saveCount}개 저장 , {notsaveCount}개 오류!");
                //MessageBox.Show("파라미터 저장", $" {saveCount}개 저장 , {notsaveCount}개 오류!");

                changeValueDataTable.Rows.Clear();
            }
        }

        private void dataGridView2_Sorted(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                if (dataGridView2.Rows[i].Cells["IsReadOnly"].Value.Equals(true))
                {
                    dataGridView2.Rows[i].Cells[0].Style.ForeColor = System.Drawing.Color.Red;
                    dataGridView2.Rows[i].Cells[1].Style.ForeColor = System.Drawing.Color.Red;
                    dataGridView2.Rows[i].Cells[1].ReadOnly = true;
                }
                if (dataGridView2.Rows[i].Cells["StorageType"].Value.Equals(((int)StorageType.ElementId)))
                {
                    dataGridView2.Rows[i].Cells[0].Style.ForeColor = System.Drawing.Color.Blue;
                    dataGridView2.Rows[i].Cells[1].Style.ForeColor = System.Drawing.Color.Blue;
                    dataGridView2.Rows[i].Cells[1].ReadOnly = true;
                }
            }
        }
    }

    public static class ExtensionMethods
    {
        public static void DoubleBuffered(this DataGridView dgv, bool setting)
        {
            Type dgvType = dgv.GetType();
            PropertyInfo pi = dgvType.GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic);
            pi.SetValue(dgv, setting, null);
        }
    }

}
