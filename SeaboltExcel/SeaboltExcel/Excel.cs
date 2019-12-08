// Peyton Seabolt


using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.ComponentModel;
using System.IO;

namespace SeaboltExcel
{
    class Excel
    {
            OpenFileDialog dgOpen;
            StreamReader sr;
            String s = "hello";
           
           
        public Excel(){
            dgOpen = new OpenFileDialog();
            dgOpen.FileOk += new CancelEventHandler(MyOk);
            dgOpen.InitialDirectory = getPath();
            dgOpen.ShowDialog();        
      }

        string getPath()
        {
            string s;
            int i;
            s = Application.StartupPath;
            i = s.IndexOf(Application.ProductName);
            s = s.Substring(0, i + Application.ProductName.Length + 1);
            return s;
        }


        void MyOk(object sender, CancelEventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application oXL;
            Microsoft.Office.Interop.Excel._Workbook oWB;
            Microsoft.Office.Interop.Excel._Worksheet oSheet;
            Microsoft.Office.Interop.Excel.Range oRng;
            object misvalue = System.Reflection.Missing.Value;
            sr = new StreamReader(dgOpen.FileName);
            string[] lines = File.ReadAllLines(dgOpen.FileName);
            try
            {
                //Start Excel and get Application object.
                oXL = new Microsoft.Office.Interop.Excel.Application();
                oXL.Visible = true;

                //Get a new workbook.
                oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
                oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

                String tempb  = "";
                String end = "";
                //Format bold, vertical alignment = center, Horizontal = right for everything but name.
                oSheet.get_Range("A1", "Q1").Font.Bold = true;
                oSheet.get_Range("B1", "Q1").Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                oSheet.get_Range("B1", "Q1").VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                String[] arr = lines[0].Split(' ');
                int temp = 1;

                // adding the format, name, tests, and average
                foreach (string element in arr)
                {
                    if (temp == 1)
                    {
                        oSheet.Cells[1, temp] = "Name";
                        oSheet.get_Range("A1").Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                    }
                    oSheet.Cells[1, temp + 1] = "Test " + (temp);
                    temp++;
                }
                oSheet.Cells[1, temp] = "Average";

                //adding the data into the columns and rows
                while (s != "")
                {          
                    //use sr to find the number of words in line 1.
                         arr = s.Split(' ');
                         char stchar = 'B';
                        int col = 1; int row = 2;
                        for (int i = 0; i < lines.Length + 4; i++) // 2 becuase we tempb at row 2
                        {
                            s = sr.ReadLine();
                            arr = s.Split(' ');
                            int l = 0;
                            for (int j = 0; j <= arr.Length + 4; j++) // 1 becuase we tempb at column 1
                            {
                                oSheet.Cells[row, col] = arr[l];
                                oRng = oSheet.Cells[row, col];
                                oRng.NumberFormat = "0.00";
                                col++;
                                l++;
                                j++;
                            }
                            tempb = stchar.ToString() + row;
                            end = ((char)(stchar + arr.Length - 3)).ToString() + row.ToString();
                            oSheet.Cells[row, col].Formula = "=AVERAGE(" + tempb + ":" + end + ")";
                            col = 1;
                            row = row + 1;
                            i++;
                        }


                        for (int i = 2; i <= arr.Length; i++)
                        {
                            tempb = stchar.ToString() + "2";
                            end = stchar.ToString() + (arr.Length + 1).ToString();
                            oSheet.Cells[row, i].Formula = "=AVERAGE(" + tempb + ":" + end + ")";
                            oRng = oSheet.Cells[row, i];
                            oRng.NumberFormat = "0.00";
                            stchar++;
                        }
                    
                }//while
            }
            catch (Exception b) { Console.WriteLine("Error!" + b ); }
        }//event
    }//class
}
