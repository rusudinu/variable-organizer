using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;

namespace VariableOrganizer
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }


       

        private void button2_Click(object sender, EventArgs e)
        {
            //register data
            string name;
            string type;
            string startingValue;
            string use;
            string links;
            string comments;
            string projectName;
            string username = Environment.UserName;

           


            if (textBox1.Text == "")
            {
                name = "-";
            }
            else   name = textBox1.Text;

            if (comboBox1.GetItemText(comboBox1.SelectedItem) == "")
            {
                type = "-";
            }
            else   type = comboBox1.GetItemText(comboBox1.SelectedItem);

            if(textBox2.Text == "")
            {
                startingValue = "-";
            }
            else startingValue = textBox2.Text;

            if (textBox3.Text == "")
            {
                use = "-";
            }
            else  use = textBox3.Text;

            if (textBox4.Text == "")
            {
                links = "-";
            }
            else  links = textBox4.Text;

            if (textBox5.Text == "")
            {
                comments = "-";
            }
            else comments = textBox5.Text;

            if (textBox6.Text == "")
            {
                projectName = "-";
            }
            else projectName = textBox6.Text;


            System.IO.Directory.CreateDirectory("C:\\Users\\" + username + "\\Desktop\\VariableOrganizer\\Variables " + projectName);
            System.IO.Directory.CreateDirectory("C:\\Users\\" + username + "\\Desktop\\VariableOrganizer\\ProgramData " + projectName);

            string variableCount = "1";
            int variableCountI = 1;
                      

            if (File.Exists("C:\\Users\\" + username + "\\Desktop\\VariableOrganizer\\ProgramData " + projectName + "\\" + "varCount.txt") == true)
            {
                StreamReader varCount = File.OpenText("C:\\Users\\" + username + "\\Desktop\\VariableOrganizer\\ProgramData " + projectName + "\\" + "varCount.txt");
                variableCount = varCount.ReadLine();
                varCount.Close();
            }

            else
            {
                StreamWriter varCountCreate = File.AppendText("C:\\Users\\" + username + "\\Desktop\\VariableOrganizer\\ProgramData " + projectName + "\\" + "varCount.txt");
                varCountCreate.WriteLine("1");
                varCountCreate.Close();

                StreamReader varCount = File.OpenText("C:\\Users\\" + username + "\\Desktop\\VariableOrganizer\\ProgramData " + projectName + "\\" + "varCount.txt");
                variableCount = varCount.ReadLine();
                varCount.Close();
            }

            variableCountI = System.Convert.ToInt32(variableCount);


            string filename = projectName + variableCountI + ".txt";

            
            StreamWriter varN = File.AppendText("C:\\Users\\" + username + "\\Desktop\\VariableOrganizer\\Variables " + projectName + "\\" + "N" + filename);
            StreamWriter varT = File.AppendText("C:\\Users\\" + username + "\\Desktop\\VariableOrganizer\\Variables " + projectName + "\\" + "T" + filename);
            StreamWriter varS = File.AppendText("C:\\Users\\" + username + "\\Desktop\\VariableOrganizer\\Variables " + projectName + "\\" + "S" + filename);
            StreamWriter varU = File.AppendText("C:\\Users\\" + username + "\\Desktop\\VariableOrganizer\\Variables " + projectName + "\\" + "U" + filename);
            StreamWriter varL = File.AppendText("C:\\Users\\" + username + "\\Desktop\\VariableOrganizer\\Variables " + projectName + "\\" + "L" + filename);
            StreamWriter varC = File.AppendText("C:\\Users\\" + username + "\\Desktop\\VariableOrganizer\\Variables " + projectName + "\\" + "C" + filename);

            varN.WriteLine(name);
            varT.WriteLine(type);
            varS.WriteLine(startingValue);
            varU.WriteLine(use);
            varL.WriteLine(links);
            varC.WriteLine(comments);

            varN.Close();
            varT.Close();
            varS.Close();
            varU.Close();
            varL.Close();
            varC.Close();


            variableCountI = variableCountI + 1;
            
            StreamWriter varCountW = new StreamWriter("C:\\Users\\" + username + "\\Desktop\\VariableOrganizer\\ProgramData " + projectName +  "\\" + "varCount.txt");
            varCountW.WriteLine(variableCountI);
            varCountW.Close();
            MessageBox.Show("Operation completed !");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //create table
            Form f2 = new Form2();
            f2.ShowDialog();

        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
