using System;
using System.Configuration;
using xlNS = Microsoft.Office.Interop.Excel;
using pptNS = Microsoft.Office.Interop.PowerPoint;
using System.Data.SqlClient;
using System.Data;
using System.IO;
using System.Net.Mail;
using System.Drawing;
using System.Windows.Media;


namespace AOAService
{
    class DataGenerator
    {
        string PresentationPath = ConfigurationManager.AppSettings["TemplatePPTPath"];
        string WorkbookPath = ConfigurationManager.AppSettings["TemplateExcelPath"];

        string DataGenarationPath = ConfigurationManager.AppSettings["DataGenerationPath"];

        public static ADODB.Recordset rs;
        public static xlNS.Worksheet targetSheet = null;
        public static xlNS.Sheets sheet1 = null;

        pptNS.Application powerpointApplication = null;
        pptNS.Presentation pptPresentation = null;
        pptNS.Slide pptSlide = null;
        pptNS.ShapeRange shapeRange = null;
        xlNS.Application excelApplication = null;
        xlNS.Workbook excelWorkBook = null;
       
       
        xlNS.ChartObjects chartObjects = null;
        xlNS.ChartObject existingChartObject = null;
        xlNS.Range tablerange = null;
        Utilities GeneralUtilities = new Utilities();
        
     
        string Time = DateTime.Now.ToString("ddMMMyyyy.hh.m.s tt");
        
        object paramMissing = Type.Missing;
        string filename;

        public void ReportsGenerator(string User_ID,string DB_Name, string Startdate, string EndDate, string Hospital, string UserName, string Email)
        {

            string ProdConn = ConfigurationManager.ConnectionStrings["ProductionConnectionString"].ToString();
            string StgConn = ConfigurationManager.ConnectionStrings["StagingConnectionString"].ToString();


                // Create an instance of PowerPoint.
                try
                {
                    powerpointApplication = new pptNS.Application();

                    // Create an instance Excel.          
                    excelApplication = new xlNS.Application();

                    DateTime dt = new DateTime();
                    dt.ToString("d");
               
                    Console.WriteLine("Enter the DB Name");
                    filename = DB_Name;

                    string currentWorkbookPath = DataGenarationPath + filename + "_" + Time + ".xlsx";

                    File.Copy(WorkbookPath, currentWorkbookPath);

                    //Below scripts adds an instance instead of opening the template
                    excelWorkBook = excelApplication.Workbooks.Open(currentWorkbookPath);
                    excelApplication.Visible = false;
                    excelApplication.DisplayAlerts = false;

                    //Connect to DB Windows Authentication

                    SqlConnection Connection = new SqlConnection(StgConn);

                    try
                    {
                        Connection.Open();

                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.ToString());
                    }

                    SqlCommand myCommand1 = new SqlCommand("[AOA_Readmissions]", Connection);
                    DataTable dt1 = new DataTable();
                    SqlDataAdapter da1 = new SqlDataAdapter(myCommand1);
                    DataSet ds1 = new DataSet();

                    SqlCommand myCommand2 = new SqlCommand("[AOA_SeverityLevelList]", Connection);
                    DataTable dt2 = new DataTable();
                    SqlDataAdapter da2 = new SqlDataAdapter(myCommand2);
                    DataSet ds2 = new DataSet();

                    SqlCommand myCommand3 = new SqlCommand("[AOA_Top5SpecialtiesList]", Connection);
                    DataTable dt3 = new DataTable();
                    SqlDataAdapter da3 = new SqlDataAdapter(myCommand3);
                    DataSet ds3 = new DataSet();

                    SqlCommand myCommand4 = new SqlCommand("[AOA_DischargeDispositionList]", Connection);
                    DataTable dt4 = new DataTable();
                    SqlDataAdapter da4 = new SqlDataAdapter(myCommand4);
                    DataSet ds4 = new DataSet();

                    SqlCommand myCommand5 = new SqlCommand("[AOA_DischargeDayList]", Connection);
                    DataTable dt5 = new DataTable();
                    SqlDataAdapter da5 = new SqlDataAdapter(myCommand5);
                    DataSet ds5 = new DataSet();

                try
                {

                    //Slide 1 - Readmissions - Key Findings by Condition
                        Console.WriteLine("Executing Readmissions - Key Findings..");
                        

                        myCommand1.CommandType = CommandType.StoredProcedure;
                        myCommand1.Parameters.Add("@DBName", SqlDbType.VarChar).Value = DB_Name;
                        myCommand1.Parameters.Add("@StartDate", SqlDbType.VarChar).Value = Startdate;
                        myCommand1.Parameters.Add("@EndDate", SqlDbType.VarChar).Value = EndDate;
                        myCommand1.Parameters.Add("@Hospital", SqlDbType.VarChar).Value = Hospital;
                        myCommand1.CommandTimeout = 1000;

                        Console.WriteLine("Finish");
                    
                        da1.Fill(ds1);
                        dt1 = ds1.Tables[0];
                        Console.WriteLine("Finish");
                       
                        //Paste data to excel                    
                        Excelpaste(excelWorkBook.Worksheets, "Readmissions by Condition", ds1, 3, 1);


                    //Slide 3 - Readmissions - Key Findings By Severity
                        Console.WriteLine("Executing Readmissions - Severity..");
                    
                        myCommand2.CommandType = CommandType.StoredProcedure;
                        myCommand2.Parameters.Add("@DBName", SqlDbType.VarChar).Value = DB_Name;
                        myCommand2.Parameters.Add("@StartDate", SqlDbType.VarChar).Value = Startdate;
                        myCommand2.Parameters.Add("@EndDate", SqlDbType.VarChar).Value = EndDate;
                        myCommand2.Parameters.Add("@Hospital", SqlDbType.VarChar).Value = Hospital;
                        myCommand1.CommandTimeout = 1000;


                    Console.WriteLine("Finish");

                        da2.Fill(ds2);
                        dt2 = ds2.Tables[0];
                        Console.WriteLine("Finish");
                        
                        //Paste data to excel                    
                        Excelpaste(excelWorkBook.Worksheets, "Readmissions by Severity", ds2, 3, 1);

                    //Slide 2 - Readmissions by Specialties

                    Console.WriteLine("Executing Readmissions - Top 5 Specialties..");


                        myCommand3.CommandType = CommandType.StoredProcedure;
                        myCommand3.Parameters.Add("@DBName", SqlDbType.VarChar).Value = DB_Name;
                        myCommand3.Parameters.Add("@StartDate", SqlDbType.VarChar).Value = Startdate;
                        myCommand3.Parameters.Add("@EndDate", SqlDbType.VarChar).Value = EndDate;
                        myCommand3.Parameters.Add("@Hospital", SqlDbType.VarChar).Value = Hospital;
                        myCommand1.CommandTimeout = 1000;


                    Console.WriteLine("Finish");

                        da3.Fill(ds3);
                        dt3 = ds3.Tables[0];
                        Console.WriteLine("Finish");
                        
                        //Paste data to excel                    
                        Excelpaste(excelWorkBook.Worksheets, "Readmissions by Department", ds3, 3, 1);


                        //Slide 4 - Readmissions by Discharge Disposition

                        Console.WriteLine("Executing Readmissions - DischarGe Disposition..");


                        myCommand4.CommandType = CommandType.StoredProcedure;
                        myCommand4.Parameters.Add("@DBName", SqlDbType.VarChar).Value = DB_Name;
                        myCommand4.Parameters.Add("@StartDate", SqlDbType.VarChar).Value = Startdate;
                        myCommand4.Parameters.Add("@EndDate", SqlDbType.VarChar).Value = EndDate;
                        myCommand4.Parameters.Add("@Hospital", SqlDbType.VarChar).Value = Hospital;
                        myCommand1.CommandTimeout = 1000;


                    Console.WriteLine("Finish");

                        da4.Fill(ds4);
                        dt4 = ds4.Tables[0];
                        Console.WriteLine("Finish");
                        
                        //Paste data to excel                    
                        Excelpaste(excelWorkBook.Worksheets, "Readmision-DischargeDisposition", ds4, 3, 1);


                        //Slide 2 - Readmissions by Specialties

                        Console.WriteLine("Executing Readmissions - Day of Discharge..");


                        myCommand5.CommandType = CommandType.StoredProcedure;
                        myCommand5.Parameters.Add("@DBName", SqlDbType.VarChar).Value = DB_Name;
                        myCommand5.Parameters.Add("@StartDate", SqlDbType.VarChar).Value = Startdate;
                        myCommand5.Parameters.Add("@EndDate", SqlDbType.VarChar).Value = EndDate;
                        myCommand5.Parameters.Add("@Hospital", SqlDbType.VarChar).Value = Hospital;
                        myCommand1.CommandTimeout = 1000;

                        Console.WriteLine("Finish");

                        da5.Fill(ds5);
                        dt5 = ds5.Tables[0];
                        Console.WriteLine("Finish");
                        
                        //Paste data to excel                    
                        Excelpaste(excelWorkBook.Worksheets, "Readmissions-DayofDischarge", ds5, 3, 1);
                }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.ToString());
                    }
                    finally {
                        Connection.Close();

                        excelWorkBook.Save();
                        Console.WriteLine("Excel Generated with the name: " + filename + "_" + Time + ".xlsx");

                       
                    }

                    // Create a PowerPoint presentation.
                    //Open with property untitled = True which is like creting a new template

                    Console.WriteLine("Working Before Power Point");

                    string currentPresentationPath = DataGenarationPath + filename + "_" + Time + ".pptx";
                    File.Copy(PresentationPath, currentPresentationPath);

                    Console.WriteLine("1. Working After Copy ");

                    pptPresentation = powerpointApplication.Presentations.Open(currentPresentationPath);


                    Console.WriteLine("2,Working After OPEN ");

                    pptSlide = pptPresentation.Slides[1];


                foreach (pptNS.Shape shape in pptSlide.Shapes)
                    {
                        if (shape.Type.ToString() == "msoTable")
                        {
                            int i = 1,j;
                            
                            foreach (DataRow row in dt1.Rows) {
                                i++; j = 0;
                                Double StdDev = double.Parse(row[4].ToString());
                                foreach (DataColumn column in dt1.Columns) {
                                    j++;
                                if (j == 5) continue;

                                shape.Table.Cell(i, j).Shape.TextFrame.TextRange.Text = row[column].ToString();
                                
                                if (j == 2)
                                {

                                    if (StdDev > 1.0D)
                                    {
                                        shape.Table.Cell(i, j).Shape.TextFrame.TextRange.Font.Color.RGB = Color.FromArgb(0, 0, 222).ToArgb();
                                    }
                                    else if (StdDev >= 0.5D && StdDev <= 1.0D)
                                    {
                                        shape.Table.Cell(i, j).Shape.TextFrame.TextRange.Font.Color.RGB = Color.FromArgb(55, 217, 255).ToArgb();
                                    }
                                    else if (StdDev < 0.5D)
                                    {
                                        shape.Table.Cell(i, j).Shape.TextFrame.TextRange.Font.Color.RGB = Color.FromArgb(0, 153, 0).ToArgb();
                                    }

                                }
                                                        
                                }
                            }
                        }
                    }

                    pptSlide = pptPresentation.Slides[3];
                     

                    
                    foreach (pptNS.Shape shape in pptSlide.Shapes)
                    {
                        if (shape.Type.ToString() == "msoTable")
                        {
                            int i = 1, j;
                            foreach (DataRow row in dt2.Rows)
                            {
                                i++; j = 0;
                                Double StdDev = double.Parse(row[3].ToString());
                                foreach (DataColumn column in dt2.Columns)
                                {
                                    j++;

                                shape.Table.Cell(i, j).Shape.TextFrame.TextRange.Text = row[column].ToString();
                                if (j ==4 )
                                {
                                    
                                    if (StdDev > 1.0D)
                                    {
                                        shape.Table.Cell(i, j).Shape.TextFrame.TextRange.Font.Color.RGB = Color.FromArgb(0, 0, 222).ToArgb();
                                    }
                                    else if (StdDev >= 0.5D && StdDev <= 1.0D)
                                    {
                                        shape.Table.Cell(i, j).Shape.TextFrame.TextRange.Font.Color.RGB = Color.FromArgb(55, 217, 255).ToArgb();
                                    }
                                    else if (StdDev < 0.5D)
                                    {
                                        shape.Table.Cell(i, j).Shape.TextFrame.TextRange.Font.Color.RGB = Color.FromArgb(0, 153, 0).ToArgb();
                                    }

                                }
                                
                               }
                            }
                        }
                    }

                PptPaste("Readmissions by Department", 2, "Readmissions By Top Specialty", 50, 150);
                PptPaste("Readmision-DischargeDisposition", 4, "Readmision-DischargeDisposition", 50, 150);
                PptPaste("Readmissions-DayofDischarge", 5, "Readmissions-DayofDischarge", 50, 150);

                da1.Dispose();
                dt1.Dispose();
                da2.Dispose();
                dt2.Dispose();
                da1.Dispose();

                excelWorkBook.Close();
                excelApplication.Quit();

                Console.WriteLine("4,Working After Paste");
           
                    pptPresentation.SaveAs(currentPresentationPath,
                           pptNS.PpSaveAsFileType.ppSaveAsOpenXMLPresentation,
                           Microsoft.Office.Core.MsoTriState.msoTrue);

                
                    pptPresentation.Close();
                    powerpointApplication.Quit();

                    Console.WriteLine("5,Working After Save");
                    Console.WriteLine("PPT Saved");              

                    Console.WriteLine("PPT Generated with the name: " + filename + "_" + Time + ".pptx");


                    Email_Send(currentPresentationPath, currentWorkbookPath, Email, DB_Name , UserName);

                    GeneralUtilities.SQLQueryExecutor(ProdConn, "UPDATE InputForReportReadmissions SET Status = 1,EndDate = getdate() , FileName =" + "'" + currentPresentationPath + "'" + " where InputUserID = " + User_ID);


                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    GeneralUtilities.SQLQueryExecutor(ProdConn, "UPDATE InputForReportReadmissions SET Status = 2  where InputUserID = " + User_ID);
                    GeneralUtilities.SendExceptionEmail("sathyakr@advisory.com,LoganatD@advisory.com", "AOA Report Failure Message", "The Report failed with the message" + " " + ex.Message + " started by user" + " " + UserName + " " + "The database name is " + DB_Name + ".", "");
                }
                finally
                {

                    Console.WriteLine("Clearing PPT");

                    try
                    {
                        if (shapeRange != null)
                        {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(shapeRange);
                        }
                   
                        shapeRange = null;
                    }
                    catch (Exception)
                    {
                        shapeRange = null;
                    }
                    finally
                    {
                        GC.Collect();
                    }
                    Console.WriteLine("Clearing PPT");
                    try
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(pptSlide);
                        pptSlide = null;
                    }
                    catch (Exception)
                    {
                        pptSlide = null;
                    }
                    finally
                    {
                        GC.Collect();
                    }
                    Console.WriteLine("Clearing PPT");
                    try
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(pptPresentation);
                        pptPresentation = null;
                    }
                    catch (Exception)
                    {
                        pptPresentation = null;
                    }
                    finally
                    {
                        GC.Collect();
                    }
                    Console.WriteLine("Clearing PPT");
                    try
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(powerpointApplication);
                        powerpointApplication = null;
                    }
                    catch (Exception)
                    {
                        powerpointApplication = null;
                    }
                    finally
                    {
                        GC.Collect();
                    }

                  //Close Excel

                    Console.WriteLine("Clearing Excel");
    //                excelWorkBook.Close(false);

                    //~~> Quit the Excel Application
      //              excelApplication.Quit();
                    Console.WriteLine("Clearing Excel");
                    try
                    {
                        if(chartObjects!= null)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(chartObjects);
                        chartObjects = null;
                    }
                    catch (Exception)
                    {
                        chartObjects = null;
                    }
                    finally
                    {
                        GC.Collect();
                    }
                    Console.WriteLine("Clearing Excel");
                    try
                    {
                        if (existingChartObject != null)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(existingChartObject);
                        existingChartObject = null;
                    }
                    catch (Exception)
                    {
                        existingChartObject = null;
                    }
                    finally
                    {
                        GC.Collect();
                    }
                    Console.WriteLine("Clearing Excel");
                    try
                    {
                        if(targetSheet!= null)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(targetSheet);
                        targetSheet = null;
                    }
                    catch (Exception)
                    {
                        targetSheet = null;
                    }
                    finally
                    {
                        GC.Collect();
                    }
                    Console.WriteLine("Clearing Excel");
                    try
                    {
                        if(excelWorkBook!=null)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkBook);
                        excelWorkBook = null;
                    }
                    catch (Exception)
                    {
                        excelWorkBook = null;
                    }
                    finally
                    {
                        GC.Collect();
                    }
                    Console.WriteLine("Clearing Excel");
                    try
                    {
                        if(excelApplication!=null)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApplication);
                        excelApplication = null;
                    }
                    catch (Exception)
                    {
                        excelApplication = null;
                    }
                    finally
                    {
                        GC.Collect();
                    }
                    //nalin end

                    Console.WriteLine("Cleared both Excel and PPT objects");

                }



        }


        public void PptPaste(string slidename,int slidenumber, string chartname, int Lposition, int Tposition)
        {
            // Get the worksheet that contains the chart.
            targetSheet =
                (xlNS.Worksheet)(excelWorkBook.Worksheets[slidename]);
            //targetSheet.Cells[9, 1] = filename;


            // Get the ChartObjects collection for the sheet.
            chartObjects =
                (xlNS.ChartObjects)(targetSheet.ChartObjects(paramMissing));

            // Get the chart to copy.
            existingChartObject = null; //added on 6-Nov
            existingChartObject =
                (xlNS.ChartObject)(chartObjects.Item(chartname));

            pptSlide = pptPresentation.Slides[slidenumber];

                                   
            try
            {
                // Copy the chart from the Excel worksheet to the clipboard.
                existingChartObject.Copy();
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            // Paste the chart into the PowerPoint presentation.
            shapeRange = pptSlide.Shapes.Paste();

            // Position the chart on the slide.
            shapeRange.Left = Lposition;
            shapeRange.Top = Tposition;

            Console.WriteLine("" + slidename + "" + chartname + "Generated Successfully");

           // existingChartObject.Delete();//Added 26-Apr-2016

        }

          public void PptPasteTable(string slidename,int slidenumber, int Lposition, int Tposition, string tab_start, string tab_end)
        {
            // Get the worksheet that contains the chart.
            targetSheet =
                (xlNS.Worksheet)(excelWorkBook.Worksheets[slidename]);
            //targetSheet.Cells[9, 1] = filename;


            //// Get the ChartObjects collection for the sheet.
            //chartObjects =
            //    (xlNS.ChartObjects)(targetSheet.ChartObjects(paramMissing));

            //// Get the chart to copy.
            //existingChartObject = null; //added on 6-Nov
            //existingChartObject =
            //    (xlNS.ChartObject)(chartObjects.Item(chartname));
                      
           // targetSheet = (Microsoft.Office.Interop.Excel.Worksheet)sheet1.get_Item(chartname);

             // xlNS.Range tablerange = null;

              //tablerange = targetSheet.get_Range("A1", "D6");
            tablerange = targetSheet.get_Range(tab_start, tab_end);

              tablerange.CopyPicture();

              
  
              //targetSheet.Copy();

              pptSlide = pptPresentation.Slides[slidenumber];
              pptSlide.Application.Activate();
              
              pptSlide.Design.Application.ActiveWindow.View.Paste();

              //pptSlide.Design.Application.ActiveWindow.View.PasteSpecial(false, false, false, false, false);
              //// Get the chart to copy.             

              //shapeRange = pptSlide.Shapes.Paste();

              
            //destinationRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormulas, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

       

            // Copy the chart from the Excel worksheet to the clipboard.
            //existingChartObject.Copy();
            //tablerange.Copy(Type.Missing);

                 

            // Paste the chart into the PowerPoint presentation.

             
  
            

            // Position the chart on the slide.
            //shapeRange.Left = Lposition;
            //shapeRange.Top = Tposition;

            Console.WriteLine("" + slidename + "" + "Generated Successfully");

        }
      
     

        //To Paste the records set in Excel sheet: parameter Tabname is the Excelsheet name of the current workbook
        public static void Excelpaste(xlNS.Sheets sheet1, string tabname, DataSet dset, int rowstartposition, int ColumnStartposition)
        {
            Microsoft.Office.Interop.Excel.Range range;
            try
            {
                System.Data.DataTable dtable = dset.Tables[0];
                rs = ConvertToRecordset(dtable);

                targetSheet = (Microsoft.Office.Interop.Excel.Worksheet)sheet1.get_Item(tabname);
                range = (xlNS.Range)targetSheet.Cells[rowstartposition, ColumnStartposition];
                range.CopyFromRecordset(rs, dtable.Rows.Count, dtable.Columns.Count);
                //Console.WriteLine(tabname+ "ran");  
                                

            }
            catch (Exception Ex)
            {
                Console.WriteLine(Ex);
            }
            finally
            {
                range = null;
            }

        }

        public static ADODB.Recordset ConvertToRecordset(DataTable inTable)
        {
            ADODB.Recordset result = new ADODB.Recordset();
            result.CursorLocation = ADODB.CursorLocationEnum.adUseClient;

            ADODB.Fields resultFields = result.Fields;
            System.Data.DataColumnCollection inColumns = inTable.Columns;

            foreach (DataColumn inColumn in inColumns)
            {

                resultFields.Append(inColumn.ColumnName
                    , TranslateType(inColumn.DataType)
                    , inColumn.MaxLength
                    , inColumn.AllowDBNull ? ADODB.FieldAttributeEnum.adFldIsNullable :
                                             ADODB.FieldAttributeEnum.adFldUnspecified
                    , null);
            }

            result.Open(System.Reflection.Missing.Value
                    , System.Reflection.Missing.Value
                    , ADODB.CursorTypeEnum.adOpenStatic
                    , ADODB.LockTypeEnum.adLockOptimistic, 0);

            foreach (DataRow dr in inTable.Rows)
            {
                result.AddNew(System.Reflection.Missing.Value,
                              System.Reflection.Missing.Value);

                for (int columnIndex = 0; columnIndex < inColumns.Count; columnIndex++)
                {
                    resultFields[columnIndex].Value = dr[columnIndex];
                }
            }

            return result;
        }

        //To handle the datatype during conversion of dataset into recordset
        static ADODB.DataTypeEnum TranslateType(Type columnType)
        {
            switch (columnType.UnderlyingSystemType.ToString())
            {
                case "System.Boolean":
                    return ADODB.DataTypeEnum.adBoolean;

                case "System.Byte":
                    return ADODB.DataTypeEnum.adUnsignedTinyInt;

                case "System.Char":
                    return ADODB.DataTypeEnum.adChar;

                case "System.DateTime":
                    return ADODB.DataTypeEnum.adDate;

                case "System.Decimal":
                    return ADODB.DataTypeEnum.adCurrency;

                case "System.Double":
                    return ADODB.DataTypeEnum.adDouble;

                case "System.Int16":
                    return ADODB.DataTypeEnum.adSmallInt;

                case "System.Int32":
                    return ADODB.DataTypeEnum.adInteger;

                case "System.Int64":
                    return ADODB.DataTypeEnum.adBigInt;

                case "System.SByte":
                    return ADODB.DataTypeEnum.adTinyInt;

                case "System.Single":
                    return ADODB.DataTypeEnum.adSingle;

                case "System.UInt16":
                    return ADODB.DataTypeEnum.adUnsignedSmallInt;

                case "System.UInt32":
                    return ADODB.DataTypeEnum.adUnsignedInt;

                case "System.UInt64":
                    return ADODB.DataTypeEnum.adUnsignedBigInt;

                case "System.String":
                default:
                    return ADODB.DataTypeEnum.adVarChar;
            }
        }

        //Mail
        public static void Email_Send(string PPTFile,string ExcelFile, string EmailID, string DatabaseName, string User)
        {

            MailMessage msg = new MailMessage();
            msg.To.Add(EmailID.ToString());
            
            msg.From = new MailAddress("Baskaran@advisory.com");
            msg.Subject = "Readmission Analytics - " + DatabaseName.ToString() + " started by user " + User;
            msg.Body = "Readmissions Powerpoint and data worksheet has been successfully generated. Thank you ! :)";
            msg.Attachments.Add(new Attachment(PPTFile));
            msg.Attachments.Add(new Attachment(ExcelFile));
            SmtpClient smclient = new SmtpClient("192.168.17.46");
            smclient.Timeout = 300000;
            smclient.Send(msg);

 
            Console.WriteLine("Mail Sent Successfully!");

        }
    }
}
