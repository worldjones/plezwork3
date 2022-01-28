using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Spire.Xls;
using Spire.Xls.Collections;
using System.IO;
using System.Collections;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Workbook = Spire.Xls.Workbook;
using Worksheet = Spire.Xls.Worksheet;
using System.Data.SqlClient;
using ExcelDataReader;
using CellRange = Spire.Xls.CellRange;

namespace Plex_work_3
{
    static class Program
    {



        static void Main(string[] args)
        {
            {



                string excel = @"C:\Users\jjo\Desktop\Example3.xls";

                //       var result = GetAllWorksheets(excel);



                Workbook workbook = new Workbook();


                workbook.LoadFromFile(excel);

                Worksheet sheet = workbook.Worksheets[0];

                //      var temp = sheet.ExportDataTable();



                Workbook wbToStream = new Workbook();

                Worksheet sheet2 = wbToStream.Worksheets[0];

                //  sheet2.Range["C10"].Text = "The sample demonstrates how to save an Excel workbook to stream.";

                // sheet2.Range["A1"].Text = "xd";


                string filename = "To_stream";

                /*

                Workbook wbFromStream = new Workbook();

                FileStream fileStream = File.OpenRead("sample.xls");

                fileStream.Seek(0, SeekOrigin.Begin);

                wbFromStream.LoadFromStream(fileStream);

                wbFromStream.SaveToFile("From_stream.xls", ExcelVersion.Version97to2003);

                fileStream.Dispose();

                System.Diagnostics.Process.Start("From_stream.xls");


                */





                for (int i = 0; i < 10; i++)
                {



                    int c = i;

                    String f = "A" + c;
                    String f2 = "B" + c;
                    String f3 = "C" + c;

                    int q = 1;

                    int file = 1;

                    int q2 = c;

                    foreach (CellRange cell in sheet.Range[f + ":A999"])

                    {






                        //IMPORTANT finder det nuværende cellenummer
                        System.Diagnostics.Debug.WriteLine(cell.RangeAddress);



                        System.Diagnostics.Debug.WriteLine(cell.Value);




                        // if (cell.Value.Equals("Step:"))

                        if (cell.Style.IncludeBorder == false)
                        {
                            q2++;

                            filename = "To_stream" + file;

                            file++;

                        }


                        if (cell.Style.IncludeBorder == true)
                        {

                            c++;

                            sheet2.Range["A" + q].Text = cell.Value;



                            FileStream file_stream = new FileStream(filename + ".xls", FileMode.Create);

                            wbToStream.SaveToStream(file_stream);


                            String q3 = "B" + q2;

                            String q4 = "C" + q2;

                            String q5 = "D" + q2;

                            String q6 = "E" + q2;

                            String q7 = "F" + q2;

                            String q8 = "G" + q2;


                            i++;

                            foreach (CellRange cell2 in sheet.Range[q3 + ":" + q3])

                            {



                                sheet2.Range["B" + q].Text = cell2.Value;

                                wbToStream.SaveToStream(file_stream);

                            }

                            foreach (CellRange cell2 in sheet.Range[q4 + ":" + q4])

                            {



                                sheet2.Range["C" + q].Text = cell2.Value;

                                wbToStream.SaveToStream(file_stream);

                            }

                            foreach (CellRange cell2 in sheet.Range[q5 + ":" + q5])

                            {



                                sheet2.Range["D" + q].Text = cell2.Value;

                                wbToStream.SaveToStream(file_stream);

                            }


                            foreach (CellRange cell2 in sheet.Range[q6 + ":" + q6])

                            {



                                sheet2.Range["E" + q].Text = cell2.Value;

                                wbToStream.SaveToStream(file_stream);

                            }

                            foreach (CellRange cell2 in sheet.Range[q7 + ":" + q7])

                            {



                                sheet2.Range["F" + q].Text = cell2.Value;

                                wbToStream.SaveToStream(file_stream);

                            }

                            foreach (CellRange cell2 in sheet.Range[q8 + ":" + q8])

                            {



                                sheet2.Range["G" + q].Text = cell2.Value;

                                wbToStream.SaveToStream(file_stream);

                            }



                            q2++;

                            q++;


                            //     sheet2.Range[f2].Text = f2.


                            //    sheet2.Range[f3].Text = cell2.Value;

                            //        sheet2.Range[f3].Text = ;









                            //stop looping through the table if cell does not have border?


                            //  sheet2.Range["A2"].Text = "apoishehexd";


                            wbToStream.SaveToStream(file_stream);

                            file_stream.Close();



                        }



                        /*
                        


                                    //write current cell to database

                                            System.Diagnostics.Debug.WriteLine(cell.Row);

                                            //Write current cell + B to database
                                            //Write current cell + C to database

                                            var range2 = sheet.Range[f];



                                           var temp3 = cell.Row.ToString();
                                           var temp4 = cell.Column.ToString();

                                            //printer current cell row ud
                                            System.Diagnostics.Debug.WriteLine(temp3);

                                            //printer current cell column ud
                                            System.Diagnostics.Debug.WriteLine(temp4);



                                            //Printer current cell text ud 
                                            System.Diagnostics.Debug.WriteLine(cell.Text);


                                        System.Diagnostics.Debug.WriteLine("mashallah");
                                    }



                                }
                            }
                                    //     temp.(26, 1);



                                    CellRange[] ranges = sheet.FindAllString("Step", false, false);

                                foreach (CellRange range in ranges)

                                {



                               //     System.Diagnostics.Debug.WriteLine(ranges[0]);


                                }




                                foreach (Worksheet item in result)

                                {
                             //       System.Diagnostics.Debug.WriteLine(item.Name);
                                }


                            }
                        }



                                public static WorksheetsCollection GetAllWorksheets(string excel)
                            {




                                Workbook workbook = new Workbook();



                                workbook.LoadFromFile(excel);

                                WorksheetsCollection worksheets = workbook.Worksheets;



                                return worksheets;





                            }


                        */








                    }


                }
            }





        }



    }
}

    


    

