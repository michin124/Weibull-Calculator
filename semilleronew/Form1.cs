using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SpreadsheetLight;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using DocumentFormat.OpenXml.Office.CustomUI;
using DocumentFormat.OpenXml.Office2010.CustomUI;
using System.Reflection.Emit;

namespace semilleronew
{
    
    public partial class Form1 : Form
    {
        public string filePath = string.Empty;
        public string filePath2 = string.Empty;
        public string filePath3 = string.Empty;
        string[] x = new string[1000];
        string[] y = new string[1000];
        string[] x1 = new string[1000];
        string[] x2 = new string[1000];
        string[] y1 = new string[1000];
        double[] yo = new double[1000];
        public double[] tabla2 = new double[1000];
        public double[,] matriz3 = new double[3, 1500];
        public int lin = 0;
        public double[] vimp = new double[315];
        int barrer = 6;
        int valor=6;
        int lo;
        public string fileContent = string.Empty;
        public SLDocument s;
        public SLDocument sus;
        
        //public SLDocument si2 = new SLDocument("C:/Users/Legion-Pc/Desktop/Datos_Enero.xlsx");
        

        public Form1()
        {
            InitializeComponent();
            
            
        }

        private void chart1_Click(object sender, EventArgs e)
        {
           


        }

        private void clock_Tick(object sender, EventArgs e)
        {
            try {
                if (filePath3 =="")
                {
                    label2.Visible = false;
                    comboBox2.Visible = false;
                    label3.Visible = false;
                    label1.Visible = false;
                    label17.Visible = false;
                    label16.Visible = false;
                    label15.Visible = false;
                    chart.Visible = false;
                    comboBox1.Visible = false;
                    button2.Visible = false;
                    label4.Visible = false;
                    label5.Visible = false;
                    label6.Visible = false;
                    label7.Visible = false;
                    label8.Visible = false;
                    label9.Visible = false;
                    label10.Visible = false;
                    label11.Visible = false;
                    label12.Visible = false;
                    label13.Visible = false;
                    label14.Visible = false;
                    button1.Visible = false;
                    button2.Visible = false;
                    button4.Visible = false;
                    button5.Visible = false;
                    button3.Visible = false;
                    label15.Visible = false;
                    textBox1.Visible = false;
                }



            if (comboBox2.Text==""&& filePath3 != "")
            {
                button6.Visible = false;
                label18.Visible=false;
                label2.Visible = true;
                comboBox2.Visible = true;
                label3.Visible = true;
                label1.Visible = false;
                label17.Visible = false;
                label16.Visible = false;
                label15.Visible = false;
                chart.Visible = false;
                comboBox1.Visible = false;
                button2.Visible = false;
                label4.Visible = false;
                label5.Visible = false;
                label6.Visible = false;
                label7.Visible = false;
                label8.Visible = false;
                label9.Visible = false;
                label10.Visible = false;
                label11.Visible = false;
                label12.Visible = false;
                label13.Visible = false;
                label14.Visible = false;
                button1.Visible = false;
                button2.Visible = false;
                button4.Visible = false;
                button5.Visible = false;
                button3.Visible = false;
                label15.Visible = false;
                textBox1.Visible = false;


            }

     
            if (comboBox2.Text == "OBSERVAR GRAFICA")
            {
                label16.Visible = false;
                label17.Visible = false;
                textBox1.Visible = false;
                button4.Visible = false;
                button2.Visible = false;
                if (filePath2!="")
                {
                    SLDocument sl = new SLDocument(filePath2);

                    chart.Visible = true;
                    label17.Visible = false;
                    comboBox1.Visible = true;
                    button5.Visible = false;
                    label15.Visible = false;
                    button2.Visible = false;
                    label4.Visible = true;
                    label5.Visible = true;
                    label6.Visible = true;
                    label7.Visible = true;
                    label8.Visible = true;
                    label9.Visible = true;
                    label10.Visible = true;
                    label11.Visible = true;
                    label12.Visible = true;
                    label13.Visible = true;
                    label14.Visible = true;
                    label1.Visible = false;
                    textBox1.Visible = false;
                    button4.Visible = false;
                    button5.Visible = false;
                    chart.Series["f(v)"].Points.Clear();
                    chart.Series["Frecuencia"].Points.Clear();
                    label6.Text = "BETA " + sl.GetCellValueAsString(120, 11);
                    label7.Text = "Intercepto: " + sl.GetCellValueAsString(121, 11);
                    label8.Text = "Eta: " + sl.GetCellValueAsString(122, 11);
                    label9.Text = "R2: " + sl.GetCellValueAsString(123, 11);
                    label10.Text = "Localizacion: " + sl.GetCellValueAsString(124, 11);
                    label11.Text = "MTBF: " + sl.GetCellValueAsString(125, 11);
                    label12.Text = "F(MTBF): " + sl.GetCellValueAsString(126, 11);
                    label13.Text = "R(MTBF): " + sl.GetCellValueAsString(127, 11);
                    label14.Text = "DESVIACION: " + sl.GetCellValueAsString(128, 11);


                    if (comboBox1.Text == "f(v)")
                    {

                        valor = 135;
                        while (!string.IsNullOrEmpty(sl.GetCellValueAsString(valor, 17)))
                        {

                            x[valor] = sl.GetCellValueAsString(valor, 3);
                            y[valor] = sl.GetCellValueAsString(valor, 17);
                            chart.Series["f(v)"].Points.AddXY(x[valor], y[valor]);
                            valor++;
                        }
                    }

                    if (comboBox1.Text == "f*(v)")
                    {

                        valor = 135;
                        while (!string.IsNullOrEmpty(sl.GetCellValueAsString(valor, 18)))
                        {

                            x[valor] = sl.GetCellValueAsString(valor, 3);
                            y[valor] = sl.GetCellValueAsString(valor, 18);
                            chart.Series["f(v)"].Points.AddXY(x[valor], y[valor]);
                            valor++;
                        }
                    }

                    if (comboBox1.Text == "Densidad Probabilidad Weibull y Frecuencia")
                    {

                        valor = 135;
                        while (!string.IsNullOrEmpty(sl.GetCellValueAsString(valor, 3)))
                        {


                            x[valor] = sl.GetCellValueAsString(valor, 3);
                            y[valor] = sl.GetCellValueAsString(valor, 18);
                            chart.Series["f(v)"].Points.AddXY(x[valor], y[valor]);


                            x1[valor] = sl.GetCellValueAsString(valor, 3);
                            y1[valor] = (sl.GetCellValueAsString(valor, 14)) + "0";
                            yo[valor] = double.Parse(y1[valor]) * 100;
                            y1[valor] = yo[valor].ToString();
                            chart.Series["Frecuencia"].Points.AddXY(x1[valor], y1[valor]);

                            valor++;
                        }
                        valor = 135;

                    }


                    if (comboBox1.Text == "Aproximacion Ecuacion de Recta")
                    {

                        valor = 6;
                        while (!string.IsNullOrEmpty(sl.GetCellValueAsString(valor, 9)))
                        {

                            x[valor] = sl.GetCellValueAsString(valor, 9);
                            y[valor] = sl.GetCellValueAsString(valor, 10);
                            chart.Series["f(v)"].Points.AddXY(x[valor], y[valor]);
                            valor++;
                        }
                    }

                    if (comboBox1.Text == "λ(v)")
                    {

                        valor = 135;
                        while (!string.IsNullOrEmpty(sl.GetCellValueAsString(valor, 19)))
                        {

                            x[valor] = sl.GetCellValueAsString(valor, 3);
                            y[valor] = sl.GetCellValueAsString(valor, 19);
                            chart.Series["f(v)"].Points.AddXY(x[valor], y[valor]);
                            valor++;
                        }
                    }



                    if (comboBox1.Text == "limpi")
                    {

                        valor = 6;
                        while (!string.IsNullOrEmpty(sl.GetCellValueAsString(valor, 3)))
                        {

                            x[valor] = "0";
                            y[valor] = "0";
                            chart.Series["f(v)"].Points.AddXY(x[valor], y[valor]);
                            valor++;

                        }
                        valor = 6;
                    }
                    }
                if (filePath2 == "")
                {
                    label1.Visible = false;
                    label1.Visible = false;
                    button5.Visible = true;
                    label15.Visible = true;
                }







                }
            if (comboBox2.Text == "GUARDAR DATO COMO")
            {
                button5.Visible = false;
                if (filePath == "")
                {
                    label17.Visible = false;
                    button2.Visible = true;
                    chart.Visible = false;
                    comboBox1.Visible = false;
                    label4.Visible = false;
                    label5.Visible = false;
                    label6.Visible = false;
                    label7.Visible = false;
                    label8.Visible = false;
                    label9.Visible = false;
                    label10.Visible = false;
                    label11.Visible = false;
                    label12.Visible = false;
                    label13.Visible = false;
                    label14.Visible = false;
                    button1.Visible = false;
                    label15.Visible = false;
                    label16.Visible = true;
                    label1.Visible = false;


                }
               if(filePath!="")
                {

                    label16.Visible = false;
                    label15.Visible = false;
                    chart.Visible = false;
                    button2.Visible = false;
                    label4.Visible = false;
                    label5.Visible = false;
                    label6.Visible = false;
                    label7.Visible = false;
                    label8.Visible = false;
                    label9.Visible = false;
                    label10.Visible = false;
                    label11.Visible = false;
                    label12.Visible = false;
                    label13.Visible = false;
                    label14.Visible = false;
                    button1.Visible = false;
                    button2.Visible = false;
                    button5.Visible = false;
                    button3.Visible = false;
                    label15.Visible = false;
                    label17.Visible = true;
                    label1.Visible = true;

                    SLDocument sl3 = new SLDocument(filePath);
                    button4.Visible = true;
                    comboBox1.Visible = false;
                    textBox1.Visible = true;                    
                    
                    chart.Series["f(v)"].Points.Clear();

                    valor = 0;
                    int num_fil, vx, vy;
                    double[] xo = new double[1000];
                    double[,] matriz2 = new double[3000, 1500];

                    vx = 0;
                    vy = 0;
                    num_fil = 1;
                    while (!string.IsNullOrEmpty(sl3.GetCellValueAsString(num_fil, 5)))
                    {

                        xo[valor] = sl3.GetCellValueAsDouble(num_fil, 5);
                        matriz2[vx, vy] = sl3.GetCellValueAsDouble(num_fil, 5);
                        //chart.Series["f(v)"].Points.AddXY(x[valor], y[valor]);
                        valor++;
                        num_fil++;
                        vy++;


                    }
                    int numv = valor;
                    int selec = 0;
                    int selec1 = 0;
                    int selec2 = 0;
                    //va coger todos
                    while (selec < numv)
                    {
                        //si el numero es igual cuente 
                        if (xo[selec1] == matriz2[0, selec])
                        {
                            matriz2[1, selec] = matriz2[1, selec] + 1;
                            selec1++;
                        }
                        else
                        {
                            //significa que el anterior numero calculado ya lo va a colocar en matriz3
                            if (matriz2[1, selec] > 0)
                            {
                                matriz3[1, selec2] = matriz2[1, selec];
                                matriz3[0, selec2] = matriz2[0, selec];
                                selec2++;
                            }
                            selec++;
                        }
                    }
                    //matriz3 en posicion 1, trae el numero de vecez repetidas
                    int z = 0;
                    int ze = 0;
                    int contador = 0;
                    double mayor = 0;
                    double penultimo = 0;
                    double antepenultimo = 0;
                    double bige = 0;
                    double bige2 = 0;
                    double bige3 = 0;
                    lin = 0;
                    int lak = 0;
                    while (matriz3[1, lak] != 0)
                    {
                        lak++;
                    }
                    while (z < lak)
                    {
                        if (matriz3[1, ze] >= matriz3[1, z])
                        {

                            mayor = matriz3[0, ze];
                            z++;

                        }
                        else
                        {
                            ze++;
                        }
                    }
                    vimp[lin] = mayor;
                    lin++;
                    z = 0;
                    ze = 1;
                    while (z < lak)
                    {
                        if (matriz3[1, ze] >= matriz3[1, z])
                        {
                            if (mayor != matriz3[0, ze])
                            {

                                penultimo = matriz3[0, ze];
                                z++;

                            }
                            else
                            {
                                z++;
                            }
                        }
                        else
                        {
                            ze++;
                        }
                    }
                    vimp[lin] = penultimo;
                    lin++;
                    z = 0;
                    ze = 2;
                    while (z < lak)
                    {
                        if (matriz3[0, ze] >= matriz3[0, z])
                        {
                            if (mayor != matriz3[0, ze])
                            {
                                if (penultimo != matriz3[0, ze])
                                {

                                    antepenultimo = matriz3[0, ze];
                                    z++;

                                }
                                else
                                {
                                    z++;
                                }
                            }
                            else
                            {
                                z++;
                            }
                        }
                        else
                        {
                            ze++;
                        }
                    }
                    vimp[lin] = antepenultimo;
                    lin++;
                    z = 0;
                    ze = 3;
                    while (z < lak)
                    {
                        if (matriz3[0, ze] >= matriz3[0, z])
                        {
                            if (mayor != matriz3[0, ze])
                            {
                                if (penultimo != matriz3[0, ze])
                                {
                                    if (antepenultimo != matriz3[0, ze])
                                    {
                                        bige = matriz3[0, ze];
                                        z++;

                                    }
                                    else
                                    {
                                        z++;
                                    }
                                }
                                else
                                {
                                    z++;
                                }
                            }
                            else
                            {
                                z++;
                            }
                        }
                        else
                        {
                            ze++;
                        }
                    }
                    vimp[lin] = bige;
                    lin++;
                    z = 0;
                    ze = 2;
                    while (z < lak)
                    {
                        if (matriz3[0, ze] >= matriz3[0, z])
                        {
                            if (mayor != matriz3[0, ze])
                            {
                                if (penultimo != matriz3[0, ze])
                                {
                                    if (antepenultimo != matriz3[0, ze])
                                    {
                                        if (bige != matriz3[0, ze])
                                        {

                                            bige2 = matriz3[0, ze];
                                            z++;

                                        }
                                        else
                                        {
                                            z++;
                                        }
                                    }
                                    else
                                    {
                                        z++;
                                    }
                                }
                                else
                                {
                                    z++;
                                }
                            }
                            else
                            {
                                z++;
                            }
                        }
                        else
                        {
                            ze++;
                        }
                    }
                    vimp[lin] = bige2;
                    lin++;
                    z = 0;
                    ze = 3;
                    while (z < lak)
                    {
                        if (matriz3[0, ze] >= matriz3[0, z])
                        {
                            if (mayor != matriz3[0, ze])
                            {
                                if (penultimo != matriz3[0, ze])
                                {
                                    if (antepenultimo != matriz3[0, ze])
                                    {
                                        if (bige != matriz3[0, ze])
                                        {
                                            if (bige2 != matriz3[0, ze])
                                            {

                                                bige3 = matriz3[0, ze];
                                                z++;

                                            }
                                            else
                                            {
                                                z++;
                                            }

                                        }
                                        else
                                        {
                                            z++;
                                        }
                                    }
                                    else
                                    {
                                        z++;
                                    }
                                }
                                else
                                {
                                    z++;
                                }
                            }
                            else
                            {
                                z++;
                            }
                        }
                        else
                        {
                            ze++;
                        }
                    }
                    vimp[lin] = bige3;
                    lin++;
                    double[] tabla = new double[1000];

                    double[] velez = new double[100];
                    double[] velez2 = new double[100];
                    int alu = 0;
                    int variable = 0;
                    int x10 = 0;
                    for (int willy = 99; willy > 0; willy--)
                    {

                        velez[willy] = matriz3[1, alu];
                        alu++;
                    }
                    Array.Sort(velez);
                    for (int willy = 0; willy < 100; willy++)
                    {
                        velez2[willy] = velez[99 - willy];
                    }



                    x10 = 0;
                    while (variable < 100)
                    {
                        for (int xi = 0; xi < velez2[variable]; xi++)
                        {
                            tabla[x10] = vimp[variable];
                            x10++;
                        }
                        variable++;
                    }

                    int por = 0;
                    while (tabla[por] != 0)
                    {
                        por++;
                    }
                    por = por - 1;
                    por = 100 / por;
                    variable = 0;
                    x10 = 0;
                    while (variable < 100)
                    {
                        for (int xi = 0; xi < velez2[variable] * por; xi++)
                        {
                            tabla[x10] = vimp[variable];
                            x10++;
                        }
                        variable++;
                    }

                    Array.Sort(tabla);
                    int num2 = 0;
                    for (int num = 0; num < 1000; num++)
                    {
                        if (tabla[num] != 0)
                        {
                            tabla2[num2] = tabla[num];
                            num2++;

                        }

                    }


                }





            }
            }

            catch { InitializeComponent(); }

        }
            
            

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        void button1_Click(object sender, EventArgs e)
        {
            try
            {


                chart.Series["...."].Points.Clear();

                if (comboBox1.Text == "F(t-δ)")
                {

                    valor = 6;
                    while (!string.IsNullOrEmpty(sus.GetCellValueAsString(valor, 3)))
                    {

                        x[valor] = sus.GetCellValueAsString(valor, 3);
                        y[valor] = sus.GetCellValueAsString(valor, 15);
                        yo[valor] = double.Parse(y[valor]) * 100;
                        y[valor] = yo[valor].ToString();
                        chart.Series["...."].Points.AddXY(x[valor], y[valor]);
                        valor++;
                    }
                    valor = 6;
                }




                if (comboBox1.Text == "R(t-δ)")
                {

                    valor = 6;
                    while (!string.IsNullOrEmpty(sus.GetCellValueAsString(valor, 3)))
                    {

                        x[valor] = sus.GetCellValueAsString(valor, 3);
                        valor++;
                    }
                    valor = 6;
                    while (!string.IsNullOrEmpty(sus.GetCellValueAsString(valor, 16)))
                    {

                        y[valor] = sus.GetCellValueAsString(valor, 16);

                        yo[valor] = double.Parse(y[valor]) * 100;
                        y[valor] = yo[valor].ToString();
                        valor++;
                    }
                    while (barrer < 100)
                    {
                        chart.Series["...."].Points.AddXY(x[barrer], y[barrer]);
                        barrer++;

                    }

                }

                if (comboBox1.Text == "f(t-δ)")
                {
                    valor = 6;
                    while (!string.IsNullOrEmpty(sus.GetCellValueAsString(valor, 3)))
                    {

                        x[valor] = sus.GetCellValueAsString(valor, 3);
                        valor++;
                    }
                    valor = 6;
                    while (!string.IsNullOrEmpty(sus.GetCellValueAsString(valor, 17)))
                    {

                        y[valor] = sus.GetCellValueAsString(valor, 17);


                        valor++;
                    }
                    while (barrer < 100)
                    {
                        chart.Series["...."].Points.AddXY(x[barrer], y[barrer]);
                        barrer++;

                    }

                }
                if (comboBox1.Text == "λ(t-δ)")
                {
                    valor = 6;
                    while (!string.IsNullOrEmpty(sus.GetCellValueAsString(valor, 3)))
                    {

                        x[valor] = sus.GetCellValueAsString(valor, 3);
                        valor++;
                    }
                    valor = 6;
                    while (!string.IsNullOrEmpty(sus.GetCellValueAsString(valor, 19)))
                    {

                        y[valor] = sus.GetCellValueAsString(valor, 19);


                        valor++;
                    }
                    while (barrer < 100)
                    {
                        chart.Series["...."].Points.AddXY(x[barrer], y[barrer]);
                        barrer++;

                    }

                }




                while (!string.IsNullOrEmpty(sus.GetCellValueAsString(valor, 2)))
                {
                    valor++;
                }
                lo = valor;
            }
            catch
            {
                InitializeComponent();
            }

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }
        
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {

            
            var fileContent = string.Empty;

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    filePath = openFileDialog.FileName;

                    //Read the contents of the file into a stream
                    var fileStream = openFileDialog.OpenFile();

                    using (StreamReader reader = new StreamReader(fileStream))
                    {
                        fileContent = reader.ReadToEnd();
                    }
                }
            }
            MessageBox.Show(fileContent, "File Content at path: " + filePath, MessageBoxButtons.OK);

            SLDocument sl3 = new SLDocument(filePath);
            }
            catch
            {
                InitializeComponent();
            }


        }

        int veces = 0;
        private void button3_Click(object sender, EventArgs e)
        {
            

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    filePath = openFileDialog.FileName;

                    //Read the contents of the file into a stream
                    var fileStream = openFileDialog.OpenFile();

                    using (StreamReader reader = new StreamReader(fileStream))
                    {
                        fileContent = reader.ReadToEnd();
                    }
                }
            }
            MessageBox.Show(fileContent, "File Content at path: " + filePath, MessageBoxButtons.OK);
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook sheet = excel.Workbooks.Open("C:/Users/Legion-Pc/Desktop/prueba.xlsx");
            Microsoft.Office.Interop.Excel.Worksheet l = excel.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;


            SLDocument sl3 = new SLDocument(filePath);
            int z = 0;
            int ze = 0;

            valor = 0;
            while (veces < 4)
            {
                int inici = 0;
                for (inici = 0; inici < matriz3[1, z]; inici++)
                {
                    // dt.Columns.Add();
                    l.Cells[valor, 3] = matriz3[0, z];
                    l.Cells[valor, 2] = valor;
                    //sl3.SetCellValue(1, 1, "true");
                    valor++;


                }
                veces++;
            }
            sheet.Save();
            sheet.Close(true, Type.Missing, Type.Missing);
            excel.Quit();
        }

        private void button4_Click(object sender, EventArgs e)
        {


            string archivo;
            archivo = textBox1.Text;
            label1.Text = "C:/Users/Legion-Pc/Desktop/"+archivo;
            archivo = "C:/Users/Legion-Pc/Desktop/" + archivo+ ".xlsx";
            int numerotab = 0;
            string fileTest = "C:/Users/Legion-Pc/Desktop/datosweill.xlsx";

            //Create COM Objects. Create a COM object for everything that is referenced
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fileTest);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;
            xlWorksheet.Cells[1, 1] = "hola";
            for (int numero=6;numero<100;numero++)
            {
                xlWorksheet.Cells[numero, 3] = tabla2[numerotab];
                numerotab++;
            }
            xlWorkbook.SaveCopyAs(archivo);
            xlWorkbook.Close();
            xlApp.Quit();
            


        }
        //elegir archivo a revisar
        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                var fileContent = string.Empty;

                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.InitialDirectory = "c:\\";
                    openFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
                    openFileDialog.FilterIndex = 2;
                    openFileDialog.RestoreDirectory = true;

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        //Get the path of specified file
                        filePath2 = openFileDialog.FileName;

                        //Read the contents of the file into a stream
                        var fileStream = openFileDialog.OpenFile();

                        using (StreamReader reader = new StreamReader(fileStream))
                        {
                            fileContent = reader.ReadToEnd();
                        }
                    }
                }
                s = new SLDocument(filePath2);
                MessageBox.Show(fileContent, "File Content at path: " + filePath2, MessageBoxButtons.OK);
            }
            catch
            {
                InitializeComponent();
            }
        }

        private void label18_Click(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            try
            {


                var fileContent = string.Empty;

                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.InitialDirectory = "c:\\";
                    openFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
                    openFileDialog.FilterIndex = 2;
                    openFileDialog.RestoreDirectory = true;

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        //Get the path of specified file
                        filePath3 = openFileDialog.FileName;

                        //Read the contents of the file into a stream
                        var fileStreams = openFileDialog.OpenFile();

                        using (StreamReader reader = new StreamReader(fileStreams))
                        {
                            fileContent = reader.ReadToEnd();
                        }
                    }
                }
                sus = new SLDocument(filePath3);
                MessageBox.Show(fileContent, "File Content at path: " + filePath3, MessageBoxButtons.OK);
            }
            catch
            {
                InitializeComponent();
            }

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }
    }
    }

