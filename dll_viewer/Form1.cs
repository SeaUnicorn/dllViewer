using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop;

namespace dll_viewer
{
    public partial class Form1 : Form
    {
        OpenFileDialog openFileDialog1 = new OpenFileDialog();
        public Form1()
        {
            InitializeComponent();
          
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string path = openFileDialog1.FileName;
                View(path);
               
               // string[] names = dll.GetManifestResourceNames();//получаем все ресурсы
            }
        }
        private void View(string path)
        {
            Microsoft.Office.Interop.Excel.Application objEx = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook objWb;
            Microsoft.Office.Interop.Excel.Worksheet objWsh;
            objWb = objEx.Workbooks.Add(System.Reflection.Missing.Value);
            objWsh = (Microsoft.Office.Interop.Excel.Worksheet)objWb.Sheets[1];

            Assembly dll = Assembly.LoadFile(path);//загружаем DLL
            Type[] types = dll.GetTypes();
            int row = 5;
            foreach (Type type in types)
            {
                objWsh.Cells[row, 2] = type.Name;
                row++;
                FieldInfo[] fields = type.GetFields();
                foreach (FieldInfo ff in fields)
                {
                    objWsh.Cells[row, 3] = ff.FieldType.Name;

                    objWsh.Cells[row, 4] = ff.Name;
                    row++;
                }
                row++;

                MethodInfo[] methods = type.GetMethods(BindingFlags.DeclaredOnly | BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public) ;
                foreach(MethodInfo mf  in methods)
                {

                    objWsh.Cells[row, 3] = mf.ReturnType.Name;
                    objWsh.Cells[row, 4] = mf.Name;
                    row++;
                }

            }

            objEx.Visible = true;
            objEx.UserControl = true;

        }
    }
}
