using System;
using Quartz;
using Quartz.Impl;
using System.Configuration;
using System.Data.SqlClient;
using System.Data;
using Common.Logging;
using AOAService;

namespace AOAService
{
     public class DataGenerationScheduler
    {

        public void Start()
        {
            // construct a scheduler factory
            ISchedulerFactory schedFact = new StdSchedulerFactory();

            // get a scheduler
            IScheduler sched = schedFact.GetScheduler();
            sched.Start();

            // define the job and tie it to our HelloJob class
            IJobDetail job = JobBuilder.Create<AOAReports>()
                .WithIdentity("AOAv1.0", "Group1")
                .Build();

            // Trigger the job to run now, and then every 30 seconds
            ITrigger trigger = TriggerBuilder.Create()
              .WithIdentity("AOAv1.0", "Group1")
              .StartNow()
              .WithSimpleSchedule(x => x
                  .WithIntervalInSeconds(30)
                  .RepeatForever())
              .Build();

            sched.ScheduleJob(job, trigger);
        }

        public void Stop()
        {

        }
    }

    public class AOAReports : IJob
    {
        public int ISrunning = 0;

        public static string ProdConn = ConfigurationManager.ConnectionStrings["ProductionConnectionString"].ToString();
        public static string StgConn = ConfigurationManager.ConnectionStrings["StagingConnectionString"].ToString();


        public void Execute(IJobExecutionContext context)
        {
            if (ISrunning == 0)
            {
                ISrunning = 1;
                ILog logger = LogManager.GetCurrentClassLogger();

                logger.Info("Creating PPT :)");
                try
                {

                    DataTable dataTable = new DataTable();
                    using (SqlConnection conn = new SqlConnection(ProdConn))
                    {
                        conn.Open();
                      
                        string queryToFetchNewTurboRequest = @"SELECT TOP 1 
                                        InputUserID,
                                        DatabaseName,
                                        Hospital,
                                        Startdate,
                                        EndDate,
                                        Email,
                                        UserName,
                                        [aprdrgw/excludes],
                                        PayerKeys
                                        FROM InputForReportReadmissions (NOLOCK)
                                        WHERE status = '-1' ORDER BY [TimeStamp]";
                        
                        SqlCommand cmd = new SqlCommand(queryToFetchNewTurboRequest, conn);
                    
                        SqlDataAdapter da = new SqlDataAdapter(cmd);

                        da.Fill(dataTable);
                   
                        conn.Close();
                        da.Dispose();

                        conn.Close();
                    }
                    Console.WriteLine(dataTable.Rows.Count.ToString());
                    if (dataTable.Rows.Count == 1)
                    {
                        string InputUserID;
                        string DBName;
                        string Startdate;
                        string EndDate;
                        string Hospital;
                        string UserName;
                        string Email;
                        string aprdrgExclusion;
                        string Payers;

                        InputUserID = dataTable.Rows[0]["InputUserID"].ToString();
                        DBName = dataTable.Rows[0]["DatabaseName"].ToString();
                        Startdate = dataTable.Rows[0]["Startdate"].ToString();
                        EndDate = dataTable.Rows[0]["EndDate"].ToString();
                        Hospital = dataTable.Rows[0]["Hospital"].ToString();
                        UserName = dataTable.Rows[0]["Username"].ToString();
                        Email = dataTable.Rows[0]["Email"].ToString();
                        aprdrgExclusion = dataTable.Rows[0]["aprdrgw/excludes"].ToString();
                        Payers = dataTable.Rows[0]["PayerKeys"].ToString();
                        Console.WriteLine("Executing SPs");

                        Utilities SetStatus = new AOAService.Utilities();

                        //Uncomment the below to change the flag in the table
                        SetStatus.SQLQueryExecutor(ProdConn, @"UPDATE InputForReportReadmissions SET status = 0 WHERE InputUserID =" + InputUserID);

                        Console.WriteLine("Calling PPT Generator");
                        DataGenerator SL = new DataGenerator();
                        SL.ReportsGenerator(InputUserID, DBName, Startdate, EndDate, Hospital, UserName, Email,aprdrgExclusion,Payers);
                    }

                }
                catch (Exception Ex)
                {
                    Console.WriteLine(Ex.Message);
                    logger.Error(Ex);
                }
                finally
                {
                    ISrunning = 0;
                }
            }
        }
    }
}
