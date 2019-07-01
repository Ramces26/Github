using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data.OleDb;
using Microsoft.Office.Interop;
using System.IO;


namespace importexport
{
    public partial class Form1 : Form
    {
        OleDbConnection conn;
        OleDbDataAdapter MyDataAdapter;
        //DataTable dt;
        // SqlConnection con = new SqlConnection(@"Data Source=RAMCES\RAMCES;Initial Catalog=Indecargo;Integrated Security=True");
        SqlConnection con = new SqlConnection(@"Data Source=RAMCES\MARADIAGA;Initial Catalog=apk;Integrated Security=true");
        private string proyect="";
        private string id="";
        private string profo="";
        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {

                System.Data.DataTable dt = new System.Data.DataTable();
            //string query = "SELECT Material.IdMaterial, Material.Producto, Kardex.CantSal FROM Kardex INNER JOIN Material ON Kardex.IdMaterial = Material.IdMaterial";
            string query = "select * from Persona";
            SqlCommand cmd = new SqlCommand(query, con);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt);


                dataGridView1.DataSource = dt;


            //DataGridViewTextBoxColumn txt2 = new DataGridViewTextBoxColumn();
            //txt2.Name = "IdProforma";
            //txt2.HeaderText = "IdProforma";
            //DataGridViewTextBoxCell dtc = new DataGridViewTextBoxCell();
            //txt2.CellTemplate = dtc;
            //dataGridView1.Columns.Insert(2, txt2);


            //    DataGridViewTextBoxColumn tx = new DataGridViewTextBoxColumn();
            //    tx.Name = "O/Compra";
            //    tx.HeaderText = "O/Compra";
            //    DataGridViewTextBoxCell txt = new DataGridViewTextBoxCell();
            //    tx.CellTemplate = txt;
            //    dataGridView1.Columns.Insert(4, tx);



            //if (dataGridView1.Columns.Count > 4)
            //{
            //    DataGridViewTextBoxColumn txt8 = new DataGridViewTextBoxColumn();
            //    txt8.Name = "IdProduccion";
            //    txt8.HeaderText = "IdProduccion";
            //    DataGridViewTextBoxCell dtc8 = new DataGridViewTextBoxCell();
            //    txt8.CellTemplate = dtc;
            //    dataGridView1.Columns.Insert(5, txt8);
            //}


            //for (int q = 0; q < dataGridView1.Rows.Count ; q++)
            //{
            //    for (int w = 4; w < dataGridView1.Columns.Count-1; w++)
            //    {
            //        dataGridView1.Rows[q].Cells[w-2].Value = "//";
            //    }
            //}


            //for (int i = 0; i < dataGridView1.Rows.Count ; i++)
            //    {
            //        for (int j = 3; j < dataGridView1.Columns.Count-1; j++)
            //        {
            //            dataGridView1.Rows[i].Cells[j+1].Value = "///";
            //        }
            //    }

            

               
            

        }


        private void button3_Click(object sender, EventArgs e)
        {
            try
            {

                SaveFileDialog fichero = new SaveFileDialog();
                fichero.Filter = "Excel (*.xls)|*.xls";
                fichero.FileName = "ArchivoExportado";
                if (fichero.ShowDialog() == DialogResult.OK)
                {
                    Microsoft.Office.Interop.Excel.Application aplicacion;
                    Microsoft.Office.Interop.Excel.Workbook libros_trabajo;
                    Microsoft.Office.Interop.Excel.Worksheet hoja_trabajo;

                    aplicacion = new Microsoft.Office.Interop.Excel.Application();
                    libros_trabajo = aplicacion.Workbooks.Add();
                    hoja_trabajo =
                        (Microsoft.Office.Interop.Excel.Worksheet)libros_trabajo.Worksheets.get_Item(1);

                    //Recorremos el DataGridView rellenando la hoja de trabajo
                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataGridView1.Columns.Count; j++)
                        {
                            if ((dataGridView1.Rows[i].Cells[j].Value == null) == false)
                            {
                                hoja_trabajo.Cells[1, j + 1] = dataGridView1.Columns[j].HeaderText;
                                hoja_trabajo.Cells[i+2, j+1 ] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                            }
                        }
                    }
                    libros_trabajo.SaveAs(fichero.FileName,
                        Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal);
                    libros_trabajo.Close(true);
                    aplicacion.Quit();
                    MessageBox.Show("Exportado");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al exportar la informacion debido a: " + ex.ToString());
            }
        }

        private void btninsert_Click(object sender, EventArgs e)
        {
            //proyect = txtproyecto.Text;
            
            //System.Data.DataTable dt2 = new System.Data.DataTable();
            ////string query = "insert into Produccion (Proyecto,Fecha,Arquitecto,Carpintero,Descripcion) values ('" + txtproyecto.Text + "','" + Convert.ToDateTime(dateTimePicker1.Value).ToString("dd/MM/yyyy") +"','"+txtxarpintero.Text+"','"+txtxarpintero.Text+"','"+txtdescripcion.Text+"')";

           
            //SqlCommand cmd2 = new SqlCommand("insetarproduccion", con);
            //SqlDataAdapter da2 = new SqlDataAdapter(cmd2);
            //cmd2.CommandType = CommandType.StoredProcedure;
            //cmd2.Parameters.AddWithValue("@Pro", txtproyecto.Text);
            //cmd2.Parameters.AddWithValue("@Fecha", dateTimePicker1.Value.Date);
            //cmd2.Parameters.AddWithValue("@Arq", txtarquitecto.Text);
            //cmd2.Parameters.AddWithValue("@Car", txtxarpintero.Text);
            //cmd2.Parameters.AddWithValue("@Desc", txtdescripcion.Text);
            //da2.Fill(dt2);

            //System.Data.DataSet dt6 = new System.Data.DataSet();
            //string query = "select IdProccion from Produccion where Proyecto='" + proyect + "'";
            //SqlCommand cmd6 = new SqlCommand(query, con);
            //SqlDataAdapter da6 = new SqlDataAdapter(cmd6);
            //da6.Fill(dt6);

            
            //id = dt6.Tables[0].Rows[0][0].ToString();

            //dataGridView2.Columns.Remove(dataGridView2.Columns[1]);

            //for (int t = 0; t < dataGridView2.Rows.Count; t++)
            //{
            //    for (int f = 3; f < dataGridView2.Columns.Count - 1; f++)
            //    {
            //        dataGridView2.Rows[t].Cells[f - 2].Value = profo;
            //    }
            //}
            //for (int a = 0; a < dataGridView2.Rows.Count; a++)
            //{
            //    for (int b = 3; b < dataGridView2.Columns.Count - 1; b++)
            //    {
            //        dataGridView2.Rows[a].Cells[b + 1].Value = id;
            //    }
            //}


            con.Open();
            //string query2 = "INSERT INTO MaterialProforma (IdMaterial,IdProforma,Cantidad,[O/Compra],IdProduccion) VALUES (@param1, @param2,@param3,@param4,@param5)";
            string query2 = "INSERT INTO Persona (Carnet,Cedula,Nombre1,Nombre2,Apellido1,Apellido2,Sexo,Perfil,Carrera,Facultad,Turno) VALUES (@param1, @param2,@param3,@param4,@param5,@param6,@param7,@param8,@param9,@param10,@param11)";
            SqlCommand cmd9 = new SqlCommand(query2, con);


            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                cmd9.Parameters.Clear();

                //cmd9.Parameters.AddWithValue("@param1", row.Cells["IdMaterial"].Value);
                //cmd9.Parameters.AddWithValue("@param2", row.Cells["IdProforma"].Value);
                //cmd9.Parameters.AddWithValue("@param3", row.Cells["CantSal"].Value);
                //cmd9.Parameters.AddWithValue("@param4", row.Cells["O/Compra"].Value);
                //cmd9.Parameters.AddWithValue("@param5", row.Cells["IdProduccion"].Value);
                cmd9.Parameters.AddWithValue("@param1", row.Cells["Carnet"].Value);
                cmd9.Parameters.AddWithValue("@param2", row.Cells["Cedula"].Value);
                cmd9.Parameters.AddWithValue("@param3", row.Cells["Nombre1"].Value);
                cmd9.Parameters.AddWithValue("@param4", row.Cells["Nombre2"].Value);
                cmd9.Parameters.AddWithValue("@param5", row.Cells["Apellido1"].Value);
                cmd9.Parameters.AddWithValue("@param6", row.Cells["Apellido2"].Value);
                cmd9.Parameters.AddWithValue("@param7", row.Cells["Sexo"].Value);
                cmd9.Parameters.AddWithValue("@param8", row.Cells["Perfil"].Value);
                cmd9.Parameters.AddWithValue("@param9", row.Cells["Carrera"].Value);
                cmd9.Parameters.AddWithValue("@param10", row.Cells["Facultad"].Value);
                cmd9.Parameters.AddWithValue("@param11",row.Cells["Turno"].Value);

                cmd9.ExecuteNonQuery();
            }


        }

        private void btnverificar_Click(object sender, EventArgs e)
        {
            profo = txtverifi.Text;
            con.Open();
            System.Data.DataSet dt2 = new System.Data.DataSet();
            string query = "select IdProforma from Proforma where IdProforma='"+txtverifi.Text+"'";

            SqlCommand cmd2 = new SqlCommand(query, con);
            SqlDataAdapter da2 = new SqlDataAdapter(cmd2);
           
            da2.Fill(dt2);
            con.Close();
            if (dt2.Tables[0].Rows.Count != 0)
            {
                MessageBox.Show("Proforma Existente");
                dateTimePicker1.Enabled = true;
                txtproyecto.Enabled = true;
                txtarquitecto.Enabled = true;
                txtxarpintero.Enabled = true;
                txtdescripcion.Enabled = true;

                btnabrir.Visible = true;
                btninsert.Visible = true;
                dataGridView2.Visible = true;
            }
            else
            {
                MessageBox.Show("Proforma No Existente");
                dateTimePicker1.Enabled = false;
                txtproyecto.Enabled = false;
                txtarquitecto.Enabled = false;
                txtxarpintero.Enabled = false;
                txtdescripcion.Enabled =false;

                btnabrir.Visible = false;
                btninsert.Visible = false;
                dataGridView2.Visible = false;
            }
        }

        private void btnabrir_Click(object sender, EventArgs e)
        {
           string ruta = "";
            OpenFileDialog dial = new OpenFileDialog();
            dial.Filter ="(*.xls;*xlsx)|*.xls;*xlsx";
            dial.Title = "Importar datos";
            dial.FileName =string.Empty;
            if (dial.ShowDialog()==DialogResult.OK)
            {
                if (dial.FileName.Equals("") == false)
                {
                    ruta = dial.FileName;
                }
            }

            string nombreHoja = "";
            nombreHoja = "hoja1";
            conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;data source=" + ruta + ";Extended Properties='Excel 12.0 Xml;HDR=Yes'");
            MyDataAdapter = new OleDbDataAdapter("Select * from [" + nombreHoja + "$]", conn);
            //dt = new DataTable();
            System.Data.DataTable data = new System.Data.DataTable();
            System.Data.DataSet datase = new System.Data.DataSet();
            conn.Open();
            MyDataAdapter.Fill(datase,"Mydata");
            data = datase.Tables["Mydata"];
            dataGridView2.DataSource = datase;
            dataGridView2.DataMember = "Mydata";
            conn.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //dataGridView2.Columns.Remove(dataGridView2.Columns[1]);

            //for (int n = 0;n<dataGridView2.Rows.Count;n++)
            //{
            //    for (int m = 0; m < dataGridView2.Columns.Count-1; m++)
            //    {

            //        //dataGridView2.Rows[n].Cells[m-2].Style.BackColor = System.Drawing.Color.Red;

            //           dataGridView2.Rows[n].Cells[m].Style.BackColor = System.Drawing.Color.Red;
            //            //dataGridView2.CurrentCell.Style.BackColor = System.Drawing.Color.Red;
            //            if(dataGridView2.Rows[n].Cells[m])
            //        { }

            //    }
            //}
            dataGridView2.ClearSelection();
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.Value.ToString()=="0")
                    {
                        
                        //cell.Selected = true;
                        row.Selected = true;

                        //dataGridView2.Rows.RemoveAt(dataGridView2.CurrentRow.Index);
                        //dataGridView2.Rows.Remove();
                        //dataGridView2.Rows[Convert.ToInt32(row)].Cells[Convert.ToInt32(cell)].Style.BackColor = Color.Red;
                    }
                }
            }

            foreach (DataGridViewRow row2 in dataGridView2.SelectedRows)
            {
                dataGridView2.Rows.Remove(row2);
            }


            //for (int t = 0; t < dataGridView2.Rows.Count; t++)
            //{
            //    for (int f =3 ; f < dataGridView2.Columns.Count - 1; f++)
            //    {
            //        dataGridView2.Rows[t].Cells[f - 2].Value = profo;
            //    }
            //}
            //for (int a = 0; a < dataGridView2.Rows.Count; a++)
            //{
            //    for (int b = 3; b < dataGridView2.Columns.Count - 1; b++)
            //    {
            //        dataGridView2.Rows[a].Cells[b+1].Value = id;
            //    }
            //}


            //con.Open();
            //string query2 = "INSERT INTO MaterialProforma (IdMaterial,IdProforma,Cantidad,[O/Compra],IdProduccion) VALUES (@param1, @param2,@param3,@param4,@param5)";
            //SqlCommand cmd9 = new SqlCommand(query2, con);


            //foreach (DataGridViewRow row in dataGridView2.Rows)
            //{
            //    cmd9.Parameters.Clear();

            //    cmd9.Parameters.AddWithValue("@param1", row.Cells["IdMaterial"].Value);
            //    cmd9.Parameters.AddWithValue("@param2", row.Cells["IdProforma"].Value);
            //    cmd9.Parameters.AddWithValue("@param3", row.Cells["CantSal"].Value);
            //    cmd9.Parameters.AddWithValue("@param4", row.Cells["O/Compra"].Value);
            //    cmd9.Parameters.AddWithValue("@param5", row.Cells["IdProduccion"].Value);

            //    cmd9.ExecuteNonQuery();
            //}

        }
    }
}
