using SpreadsheetLight;
using System.Collections;

namespace BOT_2._0
{
    public partial class Form1 : Form
    {
        private List<string> LstIdInstances;
        static class Global
        {
            public static string _globalInstance = "";

        }
        public Form1()
        {
            InitializeComponent();
            LstIdInstances = new List<string>();
            string instanciaActual;
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnGuardarInstancia_Click(object sender, EventArgs e)
        {

            // create folder with idInstancia as a name
            string dir = @"C:\bots\" + txtInstancia.Text;
            Global._globalInstance = txtInstancia.Text;
            //instanciaActual = txtInstancia.Text;
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            // show number of instances
            txtNoIntancias.Text += System.Environment.NewLine + txtInstancia.Text;
            // txtNoIntancias.Text += System.Environment.NewLine + Global._globalInstance;

            //save instances in an array
            LstIdInstances.Add(txtInstancia.Text);

            txtInstancia.Clear();

        }

        private void txtInstancia_TextChanged(object sender, EventArgs e)
        {


        }

        private void btnGuardarProveedor_Click(object sender, EventArgs e)
        {
            //create path
            string instancia = txtInstancia.Text;
            string idProducto = lblIdProducto.Text;
            string path = @"C:\bots\" + Global._globalInstance + @"\" + Global._globalInstance + idProducto + ".txt";
            //create txt file & save data
            //txtInstancia.Text
            savedDataToTxtFile(Global._globalInstance, lblIdProducto.Text, path, txtIdProveedor.Text);
            txtIdProveedor.Clear();
            /*
            if (!File.Exists(path))
            {
                //FileStream file = File.Create(path);
                string appendText = txtIdProveedor.Text + "\n";
                //File.AppendAllText(path, appendText);

            }
            else if (File.Exists(path)) {
                string appendText = txtIdProveedor.Text + "\n";
                File.AppendAllText(path, appendText);
            }
            */
        }

        private void btnSunProveedor_Click(object sender, EventArgs e)
        {
            string instancia = txtInstancia.Text;
            string sunProducto = lblSunProveedor.Text;
            string path = @"C:\bots\" + Global._globalInstance + @"\" + Global._globalInstance + sunProducto + ".txt";

            savedDataToTxtFile(txtInstancia.Text, lblSunProveedor.Text, path, txtSunProveedor.Text);
            txtSunProveedor.Clear();

        }

        private void btnProveedor_Click(object sender, EventArgs e)
        {
            string instancia = txtInstancia.Text;
            string proveedor = lblProveedor.Text;
            string path = @"C:\bots\" + Global._globalInstance + @"\" + Global._globalInstance + proveedor + ".txt";

            savedDataToTxtFile(txtInstancia.Text, lblProveedor.Text, path, txtProveedor.Text);
            txtProveedor.Clear();
        }

        private void btnTipoProveedor_Click(object sender, EventArgs e)
        {
            string instancia = txtInstancia.Text;
            string tipoProveedor = lblTipoProveedor.Text;
            string path = @"C:\bots\" + Global._globalInstance + @"\" + Global._globalInstance + tipoProveedor + ".txt";

            savedDataToTxtFile(txtInstancia.Text, lblTipoProveedor.Text, path, txtTipoProveedor.Text);
            txtTipoProveedor.Clear();
        }

        private void btnPais_Click(object sender, EventArgs e)
        {
            string instancia = txtInstancia.Text;
            string pais = lblPais.Text;
            string path = @"C:\bots\" + Global._globalInstance + @"\" + Global._globalInstance + pais + ".txt";

            savedDataToTxtFile(txtInstancia.Text, lblPais.Text, path, txtPais.Text);
            txtPais.Clear();

        }

        private void btnEstatus_Click(object sender, EventArgs e)
        {
            string instancia = txtInstancia.Text;
            string estatus = lblStatus.Text;
            string path = @"C:\bots\" + Global._globalInstance + @"\" + Global._globalInstance + estatus + ".txt";

            savedDataToTxtFile(txtInstancia.Text, lblStatus.Text, path, txtEstatus.Text);
            txtEstatus.Clear();
        }

        private void btnSolicitante_Click(object sender, EventArgs e)
        {
            string instancia = txtInstancia.Text;
            string solicitante = lblSolicitante.Text;
            string path = @"C:\bots\" + Global._globalInstance + @"\" + Global._globalInstance + solicitante + ".txt";

            savedDataToTxtFile(txtInstancia.Text, lblSolicitante.Text, path, txtSolicitante.Text);
            txtSolicitante.Clear();
        }

        public void savedDataToTxtFile(string instancia, string labelName, string path, string txtBox)
        {
            try
            {

                if (!File.Exists(path))
                {
                    //FileStream file = File.Create(path);
                    string appendText = txtBox + "\n";
                    File.AppendAllText(path, appendText);

                }
                else if (File.Exists(path))
                {
                    string appendText = txtBox + "\n";
                    File.AppendAllText(path, appendText);
                }

            }
            catch (DirectoryNotFoundException ex)
            {

                Console.WriteLine(ex.ToString());
                MessageBox.Show("The directory does not exist, DirectoryNotFoundException, below the error message: \n\n" + ex.ToString(), "Error Message !",
                MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                MessageBox.Show("Error when trying to create txt file, below the error message: \n\n" + ex.ToString(), "Error Message !",
                MessageBoxButtons.OK, MessageBoxIcon.Error);

            }

        }

        private void btnGenerar_Click(object sender, EventArgs e)
        {
            //path
            string path = @"C:\bots\import\";
            string pathTxt = @"C:\bots\";

            //loop all IdInstances in the list 
            foreach (string IdInstance in LstIdInstances)
            {
                // Validate if directory exists
                string folder = path + @"xls";
                string excelPath = folder + @"\" + IdInstance + ".xls";
                if (Directory.Exists(folder))
                {
                    //if the excel exists, delete old version
                    if (File.Exists(excelPath))
                        File.Delete(excelPath);
                }
                else
                    Directory.CreateDirectory(folder);

                // set all values from txt files
                // generar sentencia que todos los string sean del mismo tamano, no genera excel
                // si no cumple las condiciones, dejar en el log que no se completo por la condicion
                // funcion de LOG: (llamar log - reporte), recibe mensaje, id mensaje, linea/unbicacion de mensaje
                // 
                try
                {

                    string[] IdProveedor = File.ReadAllLines(pathTxt + @"\" + IdInstance + @"\" + IdInstance + lblIdProducto.Text + ".txt");
                    string[] SunProveedor = File.ReadAllLines(pathTxt + @"\" + IdInstance + @"\" + IdInstance + lblSunProveedor.Text + ".txt");
                    string[] Proveedor = File.ReadAllLines(pathTxt + @"\" + IdInstance + @"\" + IdInstance + lblProveedor.Text + ".txt");
                    string[] TipoProveedor = File.ReadAllLines(pathTxt + @"\" + IdInstance + @"\" + IdInstance + lblTipoProveedor.Text + ".txt");
                    string[] Pais = File.ReadAllLines(pathTxt + @"\" + IdInstance + @"\" + IdInstance + lblPais.Text + ".txt");
                    string[] Estatus = File.ReadAllLines(pathTxt + @"\" + IdInstance + @"\" + IdInstance + lblStatus.Text + ".txt");
                    string[] Solicitante = File.ReadAllLines(pathTxt + @"\" + IdInstance + @"\" + IdInstance + lblSolicitante.Text + ".txt");

                    int lengthIdProveedor = IdProveedor.Length;
                    int lengthSunProveedor = SunProveedor.Length;

                    MessageBox.Show("Length idProveedores:" + lengthIdProveedor + "\n\n" +
                        "Length sunProveedor: " + lengthSunProveedor , "length of arrays !",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);

                    if (IdProveedor.Contains("")) {
                        MessageBox.Show("It contains empty cells", "length of arrays !",
                         MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    // Create a new excel from txt files
                    using (SLDocument sl = new SLDocument())
                    {
                        sl.SetCellValue("A1", "IdProveedor");
                        sl.SetCellValue("B1", "SunProveedor");
                        sl.SetCellValue("C1", "Proveedor");
                        sl.SetCellValue("D1", "TipoProveedor");
                        sl.SetCellValue("E1", "Pais");
                        sl.SetCellValue("F1", "Estatus");
                        sl.SetCellValue("G1", "Solicitante");
                        for (int i = 1; i <= IdProveedor.Length; i++)
                        {
                            // check if an array has en emtpy element or if one of them has a different length

                            

                            sl.SetCellValue(i + 1, 1, IdProveedor[i - 1]);
                            sl.SetCellValue(i + 1, 2, SunProveedor[i - 1]);
                            sl.SetCellValue(i + 1, 3, Proveedor[i - 1]);
                            sl.SetCellValue(i + 1, 4, TipoProveedor[i - 1]);
                            sl.SetCellValue(i + 1, 5, Pais[i - 1]);
                            sl.SetCellValue(i + 1, 6, Estatus[i - 1]);
                            sl.SetCellValue(i + 1, 7, Solicitante[i - 1]);
                        }
                        sl.SaveAs(excelPath);
                    }
                }
                catch (FileNotFoundException ex)
                {
                    Console.WriteLine(ex.ToString());
                    MessageBox.Show("Error when trying to read txt files, FILE NOT FOUND, below the error message: \n\n" + ex.ToString(), "Error Message !",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);

                }

                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                    MessageBox.Show("Error in GENERATE BUTTON, below the error message: \n\n" + ex.ToString(), "Error Message !",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);

                }



            }

            LstIdInstances = new List<string>();

            this.Close();
        }

        private void txtNoIntancias_TextChanged(object sender, EventArgs e)
        {

        }

        private Boolean checkArrayslength(List<string[]> list) {
            Boolean sameLength = true;
            int Aux = 0;
            int arrayLength = 0;

            for (int i = 0; i < list.Count();i++) {
                for(int j = 0; j < list.Count; i++) { 
                    //Aux = list[i].Length;
                    if (list[i].Length != list[i + 1].Length) { 
                    
                    }

                }


            }
            

            return false;
        }


    }
}