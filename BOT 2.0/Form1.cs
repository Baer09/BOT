using SpreadsheetLight;
using System.Collections;
using System.Web;

namespace BOT_2._0
{
    public partial class Form1 : Form
    {
        private string fileLog;
        private string logDir;
        private List<string> LstIdInstances;
        static class Global
        {
            public static string _globalInstance = "";

        }
        public Form1()
        {
            InitializeComponent();
            fileLog = @"c:\bots\log\" + DateTime.Now.ToString("dd MM yyyy") + ".txt";
            logDir = @"c:\bots\log\";
            LstIdInstances = new List<string>();
            string instanciaActual;
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            log(fileLog,e.ToString() + " " + btnCerrar.Text);
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

            if (!Directory.Exists(logDir))
            {
                Directory.CreateDirectory(logDir);
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
            savedDataToTxtFile(txtInstancia.Text, lblIdProducto.Text, path, txtIdProveedor.Text, e.ToString());
            txtIdProveedor.Clear();            
        }

        private void btnSunProveedor_Click(object sender, EventArgs e)
        {
            string instancia = txtInstancia.Text;
            string sunProducto = lblSunProveedor.Text;
            string path = @"C:\bots\" + Global._globalInstance + @"\" + Global._globalInstance + sunProducto + ".txt";

            savedDataToTxtFile(txtInstancia.Text, lblSunProveedor.Text, path, txtSunProveedor.Text, e.ToString());
            txtSunProveedor.Clear();

        }

        private void btnProveedor_Click(object sender, EventArgs e)
        {
            string instancia = txtInstancia.Text;
            string proveedor = lblProveedor.Text;
            string path = @"C:\bots\" + Global._globalInstance + @"\" + Global._globalInstance + proveedor + ".txt";

            savedDataToTxtFile(txtInstancia.Text, lblProveedor.Text, path, txtProveedor.Text, e.ToString());
            txtProveedor.Clear();
        }

        private void btnTipoProveedor_Click(object sender, EventArgs e)
        {
            string instancia = txtInstancia.Text;
            string tipoProveedor = lblTipoProveedor.Text;
            string path = @"C:\bots\" + Global._globalInstance + @"\" + Global._globalInstance + tipoProveedor + ".txt";

            savedDataToTxtFile(txtInstancia.Text, lblTipoProveedor.Text, path, txtTipoProveedor.Text,e.ToString());
            txtTipoProveedor.Clear();
        }

        private void btnPais_Click(object sender, EventArgs e)
        {
            string instancia = txtInstancia.Text;
            string pais = lblPais.Text;
            string path = @"C:\bots\" + Global._globalInstance + @"\" + Global._globalInstance + pais + ".txt";

            savedDataToTxtFile(txtInstancia.Text, lblPais.Text, path, txtPais.Text, e.ToString());
            txtPais.Clear();

        }

        private void btnEstatus_Click(object sender, EventArgs e)
        {
            string instancia = txtInstancia.Text;
            string estatus = lblStatus.Text;
            string path = @"C:\bots\" + Global._globalInstance + @"\" + Global._globalInstance + estatus + ".txt";

            savedDataToTxtFile(txtInstancia.Text, lblStatus.Text, path, txtEstatus.Text, e.ToString());
            txtEstatus.Clear();
        }

        private void btnSolicitante_Click(object sender, EventArgs e)
        {
            string instancia = txtInstancia.Text;
            string solicitante = lblSolicitante.Text;
            string path = @"C:\bots\" + Global._globalInstance + @"\" + Global._globalInstance + solicitante + ".txt";

            

            savedDataToTxtFile(txtInstancia.Text, lblSolicitante.Text, path, txtSolicitante.Text, e.ToString());
            txtSolicitante.Clear();
        }

        public void savedDataToTxtFile(string instancia, string labelName, string path, string txtBox,string message)
        {
            path = path + instancia;
            log(path,txtBox);
            log(fileLog,message + " " + labelName+ " " + txtBox);
        }

        private void btnGenerar_Click(object sender, EventArgs e)
        {
            log(fileLog,e.ToString() + " " + btnGenerar.Text);
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

                    // adding array into a list
                    List<string[]> listOfArrays = new List<string[]>();
                    listOfArrays.Add(IdProveedor);
                    listOfArrays.Add(SunProveedor);
                    listOfArrays.Add(Proveedor);
                    listOfArrays.Add(TipoProveedor);
                    listOfArrays.Add(Pais);
                    listOfArrays.Add(Estatus);
                    listOfArrays.Add(Solicitante);

                    if ((checkLength(listOfArrays) && !checkBlanks(listOfArrays)))
                    {
                        //MessageBox.Show("el tamano de los arreglos son iguales", "length of arrays !",
                        //sMessageBoxButtons.OK, MessageBoxIcon.Error);

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
                    else if (!checkLength(listOfArrays) ){
                            //MessageBox.Show("el tamano de los arreglos no son iguales o txt tiene espacios vacios", "length of arrays !",
                            //MessageBoxButtons.OK, MessageBoxIcon.Error);

                            log(fileLog,"El Tamaño de los arreglos no son iguales");  

                    }
                    else if (checkBlanks(listOfArrays))
                    {
                        log(fileLog, "txt tiene espacios vacios");
                    }

                }
                catch (FileNotFoundException ex)
                {
                    log(fileLog,ex.ToString());
                }

                catch (Exception ex)
                {
                    log(fileLog, ex.ToString());
                }
            }

            LstIdInstances = new List<string>();

            this.Close();
        }

        private void txtNoIntancias_TextChanged(object sender, EventArgs e)
        {

        }

        private Boolean checkLength(List <string[]> listOfArrays) {
            Boolean sameSize = true;
            int sizeOfArray = listOfArrays[0].Length;
            for (int i = 1; i < listOfArrays.Count();i++) {
                if (sizeOfArray != listOfArrays[i].Length) {

                    sameSize = false;
                    break;                   
                }
            }           
            return sameSize;
        }

        private Boolean checkBlanks(List<string[]> listOfArrays) {
            Boolean blanks = false;
            for (int i = 0; i < listOfArrays.Count;i++) {
                for (int j = 0; j < listOfArrays[i].Length; j++) {
                    if (string.IsNullOrWhiteSpace(listOfArrays[i][j])) {
                        blanks = true;
                        break;                        
                    }
                }               
            }
            return blanks;
        }

        private void log(string path,string txtMessage) {
            try
            {
                if (!File.Exists(path))
                    File.WriteAllText(path, txtMessage + "\n");
                else
                    File.AppendAllText(path, txtMessage + "\n");
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
    }
}