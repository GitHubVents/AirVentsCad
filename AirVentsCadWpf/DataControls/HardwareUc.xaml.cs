using System;
using System.Collections;
using System.Data;
using System.Data.SqlClient;
using System.Windows;
using System.Windows.Controls;
using AirVentsCadWpf.AirVentsClasses;

namespace AirVentsCadWpf.DataControls
{
    /// <summary>
    /// Interaction logic for HardwareUc.xaml
    /// </summary>
    public partial class HardwareUc : UserControl
    {
        public HardwareUc()
        {
            InitializeComponent();
        }

        public string Connect = Properties.Settings.Default.ConnectionToSQL;

        private void Grid_Loaded_1(object sender, RoutedEventArgs e)
        {
            loadCmb();
        }


        #region Load to Combobox TypeLine & TypeBlock

        public void loadCmb()
        {
            try
            {
                var sql = new GenerateUnitNames();

                string valueMember, displayMember;
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

        private void CmbTypeLane_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string IDSize = Convert.ToString(CmbTypeLane.SelectedValue);
            string IDBlockType = Convert.ToString(CmbBlockType.SelectedValue);

            datagridload(IDBlockType, IDSize);
        }

        private void CmbBlockType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            string IDSize = Convert.ToString(CmbTypeLane.SelectedValue);
            string IDBlockType = Convert.ToString(CmbBlockType.SelectedValue);

            datagridload(IDBlockType, IDSize);
        }

        public void datagridload(string IDBlockType, string IDSize)
        {
            var con = new SqlConnection(Connect);
            var cmd = new SqlCommand();
            cmd.Connection = con;

            con.Open();
            
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT CoilID, Model as 'Модель', BlockName as 'Блок', Notes as 'Заметка', Row as 'Ряд' FROM AirVents.HeaterExchander WHERE BlockTypeID = '" + IDBlockType + "' AND SizeID = '" + IDSize + "' order by Model ";

            var objDs = new DataSet();
            var dAdapter = new SqlDataAdapter {SelectCommand = cmd};

            dAdapter.Fill(objDs);
            //DataGridSql.CanUserAddRows = false;

            if (objDs.Tables[0].Rows.Count > 0)
            {
                DataGridSql2.ItemsSource = objDs.Tables[0].DefaultView;
        
            }
            con.Close();
        }

        #region ADD, EDIT, DELETE
        // SAVE EDIT ROW
        private void BtnSaveEdit_Click(object sender, RoutedEventArgs e)
        {
            try
            {
            
            var con = new SqlConnection(Connect);
            var cmd = new SqlCommand();

            cmd.Connection = con;

            con.Open();

            cmd.CommandType = CommandType.Text;


            object item = DataGridSql2.SelectedItem;

            if (item == null)
            {
                return;
            }

            string CoilID = (DataGridSql2.SelectedCells[0].Column.GetCellContent(item) as TextBlock).Text;
            string Model = (DataGridSql2.SelectedCells[1].Column.GetCellContent(item) as TextBlock).Text;
            string BlockName = (DataGridSql2.SelectedCells[2].Column.GetCellContent(item) as TextBlock).Text;
            string Notes = (DataGridSql2.SelectedCells[3].Column.GetCellContent(item) as TextBlock).Text;
            string RowColumn = (DataGridSql2.SelectedCells[4].Column.GetCellContent(item) as TextBlock).Text;

            
            cmd.CommandText = "UPDATE AirVents.HeaterExchander SET Model='" + Model +
                                                
                                                    "', BlockName='" + BlockName +
                                                    "', Notes='" + Notes +
                                                    "', Row='" + RowColumn +
                                                            "' WHERE CoilID='" + CoilID + "'";
            cmd.ExecuteNonQuery();

                updatedgv(); // Обновление DGV

            con.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
         

        }

        // ADD ROW
        private void BtnAdd_Click(object sender, RoutedEventArgs e)
        {
         try
            {

                object item = DataGridSql2.SelectedItem;

            if (item == null)
            {
                return;
            }

            //string CoilID = (DataGridSql2.SelectedCells[0].Column.GetCellContent(item) as TextBlock).Text;
            string Model = (DataGridSql2.SelectedCells[1].Column.GetCellContent(item) as TextBlock).Text;
            string BlockName = (DataGridSql2.SelectedCells[2].Column.GetCellContent(item) as TextBlock).Text;
            string Notes = (DataGridSql2.SelectedCells[3].Column.GetCellContent(item) as TextBlock).Text;
            string RowColumn = (DataGridSql2.SelectedCells[4].Column.GetCellContent(item) as TextBlock).Text;

            string IDSize = Convert.ToString(CmbTypeLane.SelectedValue);
            string IDBlockType = Convert.ToString(CmbBlockType.SelectedValue);

            var con = new SqlConnection(Connect);
            var cmd = new SqlCommand();

             con.Open();

            cmd.Connection = con;

            cmd.CommandType = CommandType.Text;

            cmd.CommandText = "INSERT INTO AirVents.HeaterExchander (SizeID, BlockTypeID, Model, BlockName, Notes, Row) " +
                              "VALUES (" + IDSize + "," + IDBlockType + ",'" + Model + "', '" + BlockName + "', '" + Notes + "', '" + RowColumn + "')";
                                                    
            cmd.ExecuteNonQuery();

                updatedgv(); // Обновление DGV


            con.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void BtnDelete_Click(object sender, RoutedEventArgs e)
        {
            try
            {

            var con = new SqlConnection(Connect);
            var cmd = new SqlCommand();

            con.Open();

            var item = DataGridSql2.SelectedItem;
            if (item == null)
            {
                return;
            }
            string CoilID = (DataGridSql2.SelectedCells[0].Column.GetCellContent(item) as TextBlock).Text;

            cmd.Connection = con;

            cmd.CommandType = CommandType.Text;

            cmd.CommandText = "DELETE FROM AirVents.HeaterExchander " +
                              "WHERE CoilID=" + CoilID;

            cmd.ExecuteNonQuery();

            updatedgv(); //Обновление DGV


            con.Close();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }


        }
        #endregion

        // UPDATE DGV

        public void updatedgv()
        {

            string TypeLane = Convert.ToString(CmbTypeLane.SelectedValue);

            CmbTypeLane.SelectedIndex = 0;
            CmbTypeLane.SelectedValue = TypeLane;

            if (CmbTypeLane.SelectedIndex == 0)
            {
                CmbTypeLane.SelectedIndex = 1;
                CmbTypeLane.SelectedValue = TypeLane;
            }

            string BlockType = Convert.ToString(CmbBlockType.SelectedValue);

            CmbBlockType.SelectedIndex = 0;
            CmbBlockType.SelectedValue = BlockType;

            if (CmbBlockType.SelectedIndex == 0)
            {
                CmbBlockType.SelectedIndex = 1;
                CmbBlockType.SelectedValue = BlockType;
            }

            
        }

        private void DataGridSql2_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        
    }
}
