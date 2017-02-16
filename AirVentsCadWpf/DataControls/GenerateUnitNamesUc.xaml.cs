using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using AirVentsCadWpf.AirVentsClasses;
//using EdmLib;
using SolidWorks.Interop.sldworks;

namespace AirVentsCadWpf.DataControls
{
    /// <summary>
    /// Interaction logic for GenerateUnitNamesUc.xaml
    /// </summary>
    public partial class GenerateUnitNamesUc : UserControl
    {
        public GenerateUnitNamesUc()
        {
            InitializeComponent();
        }

        #region VARIABLES
        public string cmb16 = "Блок нагревателя водяного";
        public string cmb17 = "Блок охладителя водяного";
        public string cmb18 = "Блок охладителя фреонового";

        private string typeblock;
        private string typeline;
        private const string profilmm = "30";
        private const string beskarkaska = "F";
        private string numunit;
        private string blocktypename;
        #endregion

        public string Connect = Properties.Settings.Default.ConnectionToSQL;

        private void Grid_Loaded_1(object sender, RoutedEventArgs e)
        {
            LoadCmbVaults();
            loadCmb();

            CmbTypeLane.LayoutUpdated += ZoneText_TextChanged;
            CmbBlockType.LayoutUpdated += ZoneText_TextChanged;
            DataGridSql.SelectionChanged += ZoneText_TextChanged;

            chb30.Checked += ZoneText_TextChanged;
            chb30.Unchecked += ZoneText_TextChanged;

            chb50.Checked += ZoneText_TextChanged;
            chb50.Unchecked += ZoneText_TextChanged;

            chb70.Checked += ZoneText_TextChanged;
            chb70.Unchecked += ZoneText_TextChanged;

            chbF.Checked += ZoneText_TextChanged;
            chbF.Unchecked += ZoneText_TextChanged;

            //RBLeft.Checked += ZoneText_TextChanged;
            //RBLeft.Unchecked += ZoneText_TextChanged;

            //RBRight.Checked += ZoneText_TextChanged;
            //RBRight.Unchecked += ZoneText_TextChanged;

        }

        private string Rowbegin;

        // Переменная DataGridView ROW
        public string RowLate
        {
            get { return Rowbegin; }
            set { Rowbegin = value; }
        }

        private string BlockTypeName;

        public string blocktypenameprop
        {
            get { return BlockTypeName; }
            set { BlockTypeName = value; }
        }

        

        private string blockTypeRow;

        public string BlockTypeRow
        {
            get { return blockTypeRow; }
            set { blockTypeRow = value; }
        }


        #region Load to Combobox TypeLine & TypeBlock
        public void loadCmb()
        {
            try
            {
                var sql = new GenerateUnitNames();

                string valueMember, displayMember, blocktypename;
                object tables;

                sql.LoadTypeLine(out valueMember, out displayMember, out tables);

                CmbTypeLane.SelectedValuePath = valueMember;
                CmbTypeLane.DisplayMemberPath = displayMember;
                CmbTypeLane.ItemsSource = (IEnumerable)tables;

                CmbTypeLane.SelectedIndex = 0;

                sql.LoadBlockType(out valueMember, out displayMember, out tables);

                CmbBlockType.SelectedValuePath = valueMember;
                CmbBlockType.DisplayMemberPath = displayMember;
                CmbBlockType.ItemsSource = (IEnumerable)tables;

                CmbBlockType.SelectedIndex = 0;

                

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        #endregion
        private void ZoneText_TextChanged(object sender, EventArgs e)
        {

            // = Convert.ToString(CmbBlockType.SelectedValue);

            typeblock = BlockTypeRow;
            typeline = CmbTypeLane.Text;
            //numunit = RowLate;

           
            if (CmbBlockType.Text == cmb16)
            {
                blocktypename = "H";
            }
            if (CmbBlockType.Text == cmb17)
            {
                blocktypename = "C";
            }
            if (CmbBlockType.Text == cmb18)
            {
                blocktypename = "DX";
            }


            stringunit(typeblock, "-", typeline, summ(Prof30, Prof50, Prof70, ProfF), RowLate, blocktypename, summ2(radiobuttonleft, radiobuttonright), LR, "", "", "");
        }

        #region GENERATE UNIT NAME
        private void stringunit(string typeblock, string dot1,
        string typeline, string dot2,
        string profilmm, string dot3,
        string beskarkaska, string dot4,
        string numunit, string dot5,
        string typebloccode)
        {
            ZoneText.Text = typeblock + dot1 +
                                typeline + dot2 +
                                profilmm + dot3 +
                                beskarkaska + dot4 +
                                numunit + dot5 +
                                typebloccode;
        }
        #endregion

        public string radiobuttonleft;
        public string radiobuttonright;

        private void RBLeft_Checked(object sender, RoutedEventArgs e)
        {
            radiobuttonleft = "-L";
        }

        private void RBLeft_Unchecked(object sender, RoutedEventArgs e)
        {
            radiobuttonleft = "";
        }

        private void RBRight_Checked(object sender, RoutedEventArgs e)
        {
            radiobuttonright = "-R";
        }

        private void RBRight_Unchecked(object sender, RoutedEventArgs e)
        {
            radiobuttonright = "";
        }

        public string summ2(string radiobuttonleft, string radiobuttonright)
        {
            return radiobuttonleft + radiobuttonright;
        }

        #region Summ profil checkBox
        public string Prof30;
        public string Prof50;
        public string Prof70;
        public string ProfF;

        private void chb30_Checked(object sender, RoutedEventArgs e)
        {
            Prof30 = "-30";
            summ(Prof30, Prof50, Prof70, ProfF);
        }

        private void chb30_Unchecked(object sender, RoutedEventArgs e)
        {
            Prof30 = "";
        }

        private void chb50_Checked(object sender, RoutedEventArgs e)
        {
            Prof50 = "-50";
            summ(Prof30, Prof50, Prof70, ProfF);
        }

        private void chb50_Unchecked(object sender, RoutedEventArgs e)
        {
            Prof50 = "";
        }

        private void chb70_Checked(object sender, RoutedEventArgs e)
        {
            Prof70 = "-70";
            summ(Prof30, Prof50, Prof70, ProfF);
        }

        private void chb70_Unchecked(object sender, RoutedEventArgs e)
        {
            Prof70 = "";
        }

        private void chbF_Checked(object sender, RoutedEventArgs e)
        {
            ProfF = "-F";
            summ(Prof30, Prof50, Prof70, ProfF);
        }

        private void chbF_Unchecked(object sender, RoutedEventArgs e)
        {
            ProfF = "";
        }

        public string summ(string prof30, string prof50, string prof70, string profF)
        {
            return Prof30 + Prof50 + Prof70 + ProfF;
        }

        #endregion

        #region Load Vaults to Combobox
        public
            void LoadCmbVaults()
        {
            try
            {
            //    IEdmVault5 vault1 = new EdmVault5();
            //    IEdmVault8 vault = (IEdmVault8)vault1;
            //    //EdmViewInfo[] Views = null;
            //    Array Views = null;

            //    vault.GetVaultViews(out Views, false);
            //    ComboBoxLoadVault.Items.Clear();

            //    foreach (EdmViewInfo View in Views)
            //    {
            //        ComboBoxLoadVault.Items.Add(View.mbsVaultName);
            //    }
            //    if (ComboBoxLoadVault.Items.Count > 0)
            //    {
            //        ComboBoxLoadVault.Text = (string)ComboBoxLoadVault.Items[0];
            //    }
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show("HRESULT = 0x" + ex.ErrorCode.ToString("X") + " " + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region LOAD DGRID FROM COMBO BOX
        private void CmbTypeLane_LayoutUpdated(object sender, EventArgs e)
        {

        }

        private void CmbBlockType_LayoutUpdated(object sender, EventArgs e)
        {

        }

        private void CmbTypeLane_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string IDSize = Convert.ToString(CmbTypeLane.SelectedValue);
            string IDBlockType = Convert.ToString(CmbBlockType.SelectedValue);

            

            datagridload(IDBlockType, IDSize);
            
            // ListBox load CheckBox ListCollection
            //CreateCheckBoxList();
        }

        private void CmbBlockType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string IDSize = Convert.ToString(CmbTypeLane.SelectedValue);
            string IDBlockType = Convert.ToString(CmbBlockType.SelectedValue);

            datagridload(IDBlockType, IDSize);
            BlockTypeCode();
        }


        public void BlockTypeCode()
        {

            var con = new SqlConnection(Connect);
            var cmd = new SqlCommand();
            cmd.Connection = con;

            con.Open();

            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "select BlockTypeCode from AirVents.BlockType WHERE BlockTypeID =" + CmbBlockType.SelectedValue;

            var objDs = new DataSet();
            var dAdapter = new SqlDataAdapter();
            dAdapter.SelectCommand = cmd;
            dAdapter.Fill(objDs);

            SqlDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                string row = dr.GetValue(0).ToString();
                BlockTypeRow = row;
            }

            dr.Close();
            con.Close();

        }

        public string LR { get; set; }

        //DataGrid Selection Change to sql and get row
        private void DataGridSql_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var con = new SqlConnection(Connect);
            var cmd = new SqlCommand();
            cmd.Connection = con;

            con.Open();

            object item = DataGridSql.SelectedItem;

            if (item == null)
            {
                return;
            }

            string ModelName = (DataGridSql.SelectedCells[0].Column.GetCellContent(item) as TextBlock).Text;

            // Left or Right
            string ModelNameLeftOrRight = (DataGridSql.SelectedCells[0].Column.GetCellContent(item) as TextBlock).Text;
            string LeftOrRight = ModelNameLeftOrRight.Substring((ModelNameLeftOrRight.Length - 2));
            string lr = LeftOrRight.Substring((LeftOrRight.Length - 1));
            string dash = "-";

            if (lr != "R" & lr != "L") { lr = ""; dash = ""; }

            LR = dash + lr;
            
            //MessageBox.Show(LR);

            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT Row FROM AirVents.HeaterExchander WHERE Model = '" + ModelName + "'";

            var objDs = new DataSet();
            var dAdapter = new SqlDataAdapter();
            dAdapter.SelectCommand = cmd;
            dAdapter.Fill(objDs, "HeaterExchander");

            SqlDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                string row = dr.GetValue(0).ToString();
                RowLate = "-" + row;
               

            }

            dr.Close();
            con.Close();
        }
        #endregion

        #region Load DGV & ListBOX
        // Load To DGV Model
        public void datagridload(string IDBlockType, string IDSize)
        {
            var con = new SqlConnection(Connect);
            var cmd = new SqlCommand();
            cmd.Connection = con;

            con.Open();

            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT Model as 'Модель', BlockName as 'Блок', Notes as 'Заметка', Row as 'Ряд' FROM AirVents.HeaterExchander WHERE BlockTypeID = '" + IDBlockType + "' AND SizeID = '" + IDSize + "' order by Model ";

            var objDs = new DataSet();
            var dAdapter = new SqlDataAdapter();

            dAdapter.SelectCommand = cmd;
            dAdapter.Fill(objDs);
            DataGridSql.CanUserAddRows = false;

            if (objDs.Tables[0].Rows.Count > 0)
            {
                DataGridSql.ItemsSource = objDs.Tables[0].DefaultView;
            }
            con.Close();

            //int n = DataGridSql.Columns.Count;
            //DataGridSql.Columns[n].Width = DataGridLength.Auto;
        }

        // Load to ListBOX TypeLine
        public void CreateCheckBoxList()
        {
            var con = new SqlConnection(Connect);
            var cmd = new SqlCommand();
            cmd.Connection = con;

            con.Open();

            cmd.CommandType = CommandType.Text;

            cmd.CommandText = @"SELECT
                                Profil.Description
                                FROM AirVents.DimensionType
                                INNER JOIN AirVents.Dimension
                                ON AirVents.DimensionType.DimensionID = AirVents.Dimension.DimensionID
                                INNER JOIN AirVents.StandardSize
                                ON AirVents.DimensionType.SizeID = AirVents.StandardSize.SizeID
                                INNER JOIN AirVents.Profil
                                ON AirVents.DimensionType.ProfilID = AirVents.Profil.ProfilID
                                WHERE AirVents.StandardSize.SizeID = " + CmbTypeLane.SelectedValue;

            var dAdapter = new SqlDataAdapter();
                dAdapter.SelectCommand = cmd;

            var dt = new DataTable();

            dAdapter.Fill(dt);

            Customers = new ObservableCollection<CheckedListItem<Customer>>();

            ListBoxTypeLine.ItemsSource = Customers;
            Customers.Clear();

            foreach (DataRow datarow in dt.Rows)
            {
                Customers.Add(new CheckedListItem<Customer>(new Customer() { Name = datarow["Description"].ToString() }));
            }

            DataContext = this;
            con.Close();
        }
        #endregion
       
        #region CheckBoxListItems ListCollection
        public class Customer
        {
            public string Name { get; set; }
        }

        public class CheckedListItem<T> : INotifyPropertyChanged
        {
            public event PropertyChangedEventHandler PropertyChanged;
      
            private bool isChecked;
            private T item;

            public CheckedListItem()
            { }

            public CheckedListItem(T item, bool isChecked = false)
            {
                this.item = item;
                this.isChecked = isChecked;
            }

            public T Item
            {
                get { return item; }
                set
                {
                    item = value;
                    if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs("Item"));
                }
            }

            public bool IsChecked
            {
                get { return isChecked; }
                set
                {
                    isChecked = value;
                    if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs("IsChecked"));

                }
            }
        }

        public ObservableCollection<CheckedListItem<Customer>> Customers { get; set; }

        private string NameValueCheckBox;

        public string nameValueChkBx
        {
            get { return NameValueCheckBox; }
            set { NameValueCheckBox = value; }
        }

        #endregion

       
   

        #region SAVE BUTTON


        private string foldernameSQL;

        public string folderNameSql
        {
            get { return foldernameSQL; }
            set { foldernameSQL = value; }
        }
       
        private void ButtonSave_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SldWorks swapp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
                ModelDoc2 swmodel = swapp.ActiveDoc;

                //IEdmVault5 vault1 = new EdmVault5();

                //IEdmFolder5 folder;

                //vault1.LoginAuto(Properties.Settings.Default.PdmBaseName, 0);
                ////vault1.LoginAuto("Tets_debag", 0);

                //string vaultname = vault1.RootFolderPath;

                ////MessageBox.Show(vaultName);

                //vault1.GetFileFromPath(vaultname, out folder);

                string nameunit = ZoneText.Text + ".sldasm";

                string foldername16 = @"\Проекты\Blauberg\16 - Узел нагревателя водяного\";
                string foldername17 = @"\Проекты\Blauberg\17 - Узел охладителя водяного\";
                string foldername18 = @"\Проекты\Blauberg\18 - Узел фреонового охладителя\";

                folderNameSql = foldername16 + nameunit;

                //string filename16 = vaultName + @"\" + nameunit;

                string filename16 = null;// = vaultname + foldername16 + nameunit;
                string filename17 = null;// = vaultname + foldername17 + nameunit;
                string filename18 = null;// = vaultname + foldername18 + nameunit;

                //MessageBox.Show(filename16);

                if (CmbBlockType.Text == cmb16)
                {
                    swmodel.SaveAs(filename16);
                    MessageBox.Show(@"Документы сохранен в " + filename16);
                }

                if (CmbBlockType.Text == cmb17)
                {
                    swmodel.SaveAs(filename17);
                    MessageBox.Show(@"Документы сохранен в " + filename17);
                }

                if (CmbBlockType.Text == cmb18)
                {
                    swmodel.SaveAs(filename18);
                    MessageBox.Show(@"Документы сохранен в " + filename18);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                //MessageBox.Show("Запустите SolidWorks и откройте сборку!");
            }

            //var permissions = Customers;
            foreach (CheckBox item in ListBoxTypeLine.Items)
            {
                if (item.IsChecked == true)
                {
                    try
                    {
                        int size = Convert.ToInt32(CmbTypeLane.SelectedValue);
                        int blocktype = Convert.ToInt32(CmbBlockType.SelectedValue);
                        string profil = Convert.ToString(item.Content);
                        object itemm = DataGridSql.SelectedItem;
                        string Coil = (DataGridSql.SelectedCells[0].Column.GetCellContent(itemm) as TextBlock).Text;

                        SqlConnection con = new SqlConnection(Connect);
                        con.Open();

                        SqlCommand spcmd = new SqlCommand("AirVents_SetHeaterExchanderPath", con);
                        spcmd.CommandType = CommandType.StoredProcedure;

                        spcmd.Parameters.Add("@Size", SqlDbType.Int).Value = size;
                        spcmd.Parameters.Add("@BlockType", SqlDbType.Int).Value = blocktype;
                        spcmd.Parameters.Add("@Profil", SqlDbType.NVarChar).Value = profil;
                        spcmd.Parameters.Add("@Coil", SqlDbType.NVarChar).Value = Coil;
                        spcmd.Parameters.Add("@Path", SqlDbType.NVarChar).Value = folderNameSql;
           
                        spcmd.ExecuteReader();

                        con.Close();

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                
            }
    

        }
        #endregion
    }
}
