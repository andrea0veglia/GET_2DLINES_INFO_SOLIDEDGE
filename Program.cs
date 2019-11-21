using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Get_info_SolidEdge
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.ForegroundColor = ConsoleColor.White;

            Console.Title = "Get info - SE";
            Console.WriteLine("Returns the line information of the 2D model opened in SolidEdge");
            Console.WriteLine("Andrea Veglia");

            Console.WriteLine("Caution may affect the work done in Solid Edge \n");
            Console.WriteLine("Press any key to launch the application. ");
            Console.ReadKey();

            Console.Clear();

            //SET THE LOG FILE NAME
            string NOME_FILE_LOG = Path.Combine("LOGS", "LOG_" + DateTime.UtcNow.ToString("MM'-'dd'-'yyyy'.'HH'-'mm'-'ss") + ".txt");

            //CREATE THE LOG FILE
            CREATE_FILE_LOG(NOME_FILE_LOG);

            bool? se_close = null;

            try
            {
                SolidEdgeFramework.Application seApp = null;

                //I LOOK FOR THE SOLID EDGE PROCESS
                Type type = Type.GetTypeFromProgID("SolidEdge.Application"); 
                Process[] remoteAll = Process.GetProcessesByName("Edge");

                if (remoteAll.Count() > 0)
                {
                    //SOLID EDGE IS ALREADY OPEN 
                    seApp = (SolidEdgeFramework.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("SolidEdge.Application"); 
                    se_close = false;
                }
                else
                {
                    Console.WriteLine("Solid edge is not open.");
                    se_close = true;

                }

                if (se_close == false)
                {
                    //SET THE ACTIVE WINDOW
                    SolidEdgeDraft.SheetWindow win = (SolidEdgeDraft.SheetWindow)seApp.ActiveWindow;
                    win.Fit();

                    var HEIGHT = win.ActiveSheet.SheetSetup.SheetHeight.ToString();
                    var WIDTH = win.ActiveSheet.SheetSetup.SheetWidth.ToString();
                    var DIMENSIONS = win.ActiveSheet.Dimensions.ToString();

                    WRITE_FILE_LOG("HEIGHT: " + HEIGHT, NOME_FILE_LOG, false);
                    WRITE_FILE_LOG("WIDTH: " + WIDTH, NOME_FILE_LOG, false);
                    WRITE_FILE_LOG("DIMENSIONI: " + DIMENSIONS, NOME_FILE_LOG, false);

                    var line = win.ActiveSheet.Lines2d;

                    var FOUND_LINES = line.Count;

                    WRITE_FILE_LOG("FOUND LINES: " + FOUND_LINES.ToString(), NOME_FILE_LOG, false);

                    for (int i = 1; i <= FOUND_LINES; i++)
                    {
                        var line_color = win.ActiveSheet.Lines2d.Item(i).Style.LinearColor;
                        var line_name = win.ActiveSheet.Lines2d.Item(i).Name;
                        var index = win.ActiveSheet.Lines2d.Item(i).Index;
                        var layer = win.ActiveSheet.Lines2d.Item(i).Layer.Trim().ToLower();
                        var angle = win.ActiveSheet.Lines2d.Item(i).Angle.ToString();

                        WRITE_FILE_LOG("LINE NAME: " + line_name + " - INDEX: " + index + " - LINE COLORE: " + line_color + " - LAYER: " + layer + " - ANGLE: " + angle, NOME_FILE_LOG, false);

                    }

                }
                else
                {
                    Console.WriteLine("The SolidEdge process is closed");
                }


            }
            catch (Exception ex)
            {
                WRITE_FILE_LOG("ERROR: " + ex.ToString(), NOME_FILE_LOG, true);
            }

        }

        #region FILE LOG
        static int CREATE_FILE_LOG(string name_file_log)
        {
            int success = 0;

            try
            {

                if (!Directory.Exists("LOGS"))
                {
                    Directory.CreateDirectory("LOGS");
                }

                if (File.Exists(name_file_log))
                {
                    File.Delete(name_file_log);
                }

                using (File.Create(name_file_log)) { };


                File.AppendAllText(name_file_log, "+-----------------------------------+" + Environment.NewLine + "| LOG " + DateTime.Now.ToString() + " - SE |" + Environment.NewLine + "+-----------------------------------+");

                success = 1;
            }
            catch (Exception e)
            {
                Console.WriteLine("Log file creation error: \n" + e.Message.ToString());
                success = 0;
            }

            return success;
        }

        static int WRITE_FILE_LOG(string messaggio_log, string nome_file_log, bool error)
        {
            int success = 0;

            try
            {
                if (File.Exists(nome_file_log))
                {
                    if (error == false)
                    {
                        string NEWLINE = Environment.NewLine;
                        File.AppendAllText(nome_file_log, NEWLINE);
                        File.AppendAllText(nome_file_log, Environment.NewLine + DateTime.Now.ToString() + " - " + messaggio_log);

                        success = 1;
                    }
                    else if (error == true)
                    {
                        string NEWLINE = Environment.NewLine;
                        File.AppendAllText(nome_file_log, NEWLINE + NEWLINE);
                        File.AppendAllText(nome_file_log, "!ERROR!");
                        File.AppendAllText(nome_file_log, Environment.NewLine + DateTime.Now.ToString() + " - " + messaggio_log);
                    }

                }
            }
            catch
            {
                success = 0;
            }

            return success;
        }

        #endregion
    }
}
