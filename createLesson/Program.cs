using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.Json;
using Excel = Microsoft.Office.Interop.Excel;       //microsoft Excel 14 object in references-> COM tab

namespace Sandbox
{




    public class Read_From_Excel
    {

        public class Step
        {
            public int step { get; set; }
            public string en { get; set; }
            public string id { get; set; }

        }

        public class Page
        {
            public int page { get; set; }
            //public List<Step>  LstSteps = new List<Step>();   
            public List<Step> steps { get; set; }
        }

        public class Sel
        {
     
            public int choice { get; set; }
            
            public string description { get; set; } 
        }



        public class Quiz
        {
            public int nbr { get; set; }

            public string context { get; set; }

            public string question { get; set; }

            public List<Sel> selections  { get; set; }

            public int answer { get; set; }

            public string reason { get; set; }

        }



        public class Pages
        {
            //   public List<Page> LstPages = new List<Page>();
            public List<Page> pages { get; set; }
        }

        public class Quizes
        {
            //   public List<Page> LstPages = new List<Page>();
            public List<Quiz> quizes { get; set; }
        }



        public static void Main()
        {




            JsonConversation();

           // JsonQuiz();
        }
        
        public static void JsonConversation()
        {

            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\data\inglesGuru\lesson1ver12.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;

            string colValue = "";

            Pages pg = new Pages();
            pg.pages = new List<Page>();


            int cntPages = 0;
            int cntSteps = 0;


            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!
            for (int i = 2; i <= rowCount; i++)
            {
                // for (int j = 1; j <= colCount; j++)
                // {
                //new line
                //     if (j == 1)
                //            Console.Write("\r\n");

                //write the value to the console
                if (xlRange.Cells[i, 1] != null && xlRange.Cells[i, 1].Value2 != null)
                {



                    if (i % 2 == 0)
                    {
                        pg.pages.Add(new Page());
                        cntPages++;
                        pg.pages[cntPages - 1].page = cntPages;
                        pg.pages[cntPages - 1].steps = new List<Step>();


                        var step = new Step();
                        cntSteps++;
                        step.step = cntSteps;
                        step.en = xlRange.Cells[i, 2].Value2.ToString();
                        step.id = xlRange.Cells[i, 3].Value2.ToString();
                        pg.pages[pg.pages.Count - 1].steps.Add(step);
                    }
                    else
                    {
                        var step = new Step();
                        cntSteps++;
                        step.step = cntSteps;
                        step.en = xlRange.Cells[i, 4].Value2.ToString();
                        step.id = xlRange.Cells[i, 5].Value2.ToString();
                        pg.pages[pg.pages.Count - 1].steps.Add(step);

                    }


                }


                    }




       
            var json = JsonSerializer.Serialize(pg);

            File.WriteAllText(@"C:\data\inglesGuru\lesson1.json", json);



            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }

        public static void JsonQuiz()
        {

            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\data\inglesGuru\lesson1ver12.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[2];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;

    

            Quizes qz = new Quizes();
            qz.quizes = new List<Quiz>();

            int cntQuizes = 0;
            

            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!
            for (int i = 1; i <= rowCount; i++)
            {
                //write the value to the console
                if (xlRange.Cells[i, 1] != null && xlRange.Cells[i, 1].Value2 != null)
                {

                    string context = "";
                    string question = "";
                    string answer = "";
                    string option = "";
                    string reason = "";

                    if (xlRange.Cells[i, 1].Value2 != null)
                    {
                        context = xlRange.Cells[i, 1].Value2.ToString();
                    }


                    if (xlRange.Cells[i, 2].Value2 != null)
                    {
                        question = xlRange.Cells[i, 2].Value2.ToString();
                    }

                    if (xlRange.Cells[i, 3].Value2 != null)
                    {
                        option = xlRange.Cells[i, 3].Value2.ToString();
                    }

                    if (xlRange.Cells[i, 4].Value2 != null)
                    {
                        answer = xlRange.Cells[i, 4].Value2.ToString();
                    }

                    if (xlRange.Cells[i, 5].Value2 != null)
                    {
                        reason = xlRange.Cells[i, 5].Value2.ToString();
                    }

                    if (question == "Question")
                    {
                        qz.quizes.Add(new Quiz());
                        
                        cntQuizes++;
                        qz.quizes[cntQuizes - 1].nbr = cntQuizes;
                        qz.quizes[cntQuizes - 1].context = "";
                        qz.quizes[cntQuizes - 1].question = "";
                        qz.quizes[cntQuizes - 1].selections = new List<Sel>();
                        qz.quizes[cntQuizes - 1].answer = 0;
                        qz.quizes[cntQuizes - 1].reason = "";

                    }
                    else
                    {
                        if (context != "")
                        {
                            qz.quizes[cntQuizes - 1].context = qz.quizes[cntQuizes - 1].context + context + "\n";
                        }

                        if (question != "")
                        {
                            qz.quizes[cntQuizes - 1].question = question;
                        }

                        if (option != "")
                       {
                            //option = new Option(1,"");

                            var sel = new Sel();
                            sel.choice = qz.quizes[cntQuizes - 1].selections.Count + 1;
                            sel.description = option;
                            qz.quizes[cntQuizes - 1].selections.Add(sel);

                        }

                        if (answer != "")
                        {
                            qz.quizes[cntQuizes - 1].answer = int.Parse(answer);

                        }

                        if (reason != "")
                        {
                            qz.quizes[cntQuizes - 1].reason = qz.quizes[cntQuizes - 1].reason + reason + "\n";
                        }


                    }



                }


            }





            var json = JsonSerializer.Serialize(qz);

            File.WriteAllText(@"C:\data\inglesGuru\quiz1.json", json);



            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }


    }

}
