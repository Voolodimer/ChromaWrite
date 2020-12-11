using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelDataReader;
using System.Data;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace ChromaWrite
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            while (true)
            {
                try
                {
                    FileStream stream = File.Open(OpenCSVFile(), FileMode.Open, FileAccess.Read);

                    //2. Reading from a OpenXml Excel file (2007 format; *.xlsx)
                    IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);

                    //3. DataSet - The result of each spreadsheet will be created in the result.Tables
                    DataSet result = excelReader.AsDataSet();
                    //excelReader.Close();
                    //stream.Close();
                    string whole_file = "";
                    foreach (DataRow row in result.Tables[0].Rows)
                    {
                        foreach (DataColumn col in result.Tables[0].Columns)
                        {

                            whole_file += row[col.ColumnName].ToString();
                            if (row[col.ColumnName].ToString() == "")
                            {
                                continue;
                            }
                            whole_file += ";";
                        }
                        whole_file += "\n";
                        //Console.WriteLine();
                    }
                    //Console.WriteLine(whole_file);
                    
                    // Разделяем на строки используя символ '\r' (возврат каретки), и удаляем пустые символы*/
                    string[] lines = whole_file.Split(new char[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);

                    // Console.WriteLine(lines[0]);
                    //в эти переменные запишем индекс строк которые начинаются со слов "Название" и "Олефины"
                    int firstEnter = -1, secondEnter = -1, enterNumOne = -1;
                    //порядок в котором нужно вывести компоненты
                    //записываем в переменную string название хроматограммы
                    string titleOfChroma = lines[0].Substring((lines[0].IndexOf(':', 13) + 2), (lines[0].Length - 3 - (lines[0].IndexOf(':', 13))));
                    //получаем номер эксперимента из названия хроматограммы
                    int start_index = titleOfChroma.IndexOf('-') + 1;
                    int end_index = titleOfChroma.IndexOf('-', start_index);
                    //получаем номер эксперимента из названия хроматограммы
                    string numberOfExp = titleOfChroma.Substring(start_index, end_index - start_index);
                    //получаем номер установки
                    string nameOfPlant = titleOfChroma.Substring(0, titleOfChroma.IndexOf('-'));
                    //получаем путь до файла
                    string pathToList = Directory.GetCurrentDirectory() + "\\Список целевых компонентов " + nameOfPlant;
                    OpenExcelComponents(pathToList, out string[,] ListKeyComponents);
                    
                    string[,] printValues = new string[ListKeyComponents.GetLength(0), 2];
                    //получаем индексы строк которые начинаются со слов "Название" и "Олефины"
                    for (int i = 0; i < lines.Length; i++)
                    {
                        if (lines[i].StartsWith("Название"))
                            firstEnter = i;
                        else if (lines[i].StartsWith("Олефины") || lines[i].StartsWith("Оксигенаты") || lines[i].StartsWith("Парафины") || lines[i].StartsWith("Изопарафины"))
                            enterNumOne = i;
                        //Console.WriteLine(lines[i]);
                    }
                    //Ищем начало целевых компонентов
                    for (int j = enterNumOne; j < lines.Length; j++)
                    {
                        if (lines[j].StartsWith("1"))
                        {
                            secondEnter = j;
                            break;
                        }
                    }

                    if (firstEnter < 0 | secondEnter < 0)
                        Console.WriteLine("Неправильно опоределены вхождения слова \"Название\" или \"Олефины\"");

                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine(titleOfChroma); //вывод названия хроматограммы консоль
                    //Console.WriteLine(Directory.GetCurrentDirectory());
                    /*создаём массив размерностью lines[secondEnter].Split(';').Length (столбцы, столько значений в строке)
                    на lines.Length-secondEnter] (строки), количество строк начала целевых данных, до конца массива*/
                    string[,] TargetLines = new string[lines.Length - secondEnter, lines[secondEnter].Split(';').Length + 1];

                    Console.ForegroundColor = ConsoleColor.Red;
                    //записываем целевые значения в новый массив TargetLines. Значения в том виде, в котором они записаны в хроматограмме
                    for (int i = secondEnter, k = 0; i < lines.Length; i++, k++)
                    {
                        try
                        {
                            //разбиваем считанные строки разделенные ;
                            string[] SepValues = lines[i].Split(';');
                            for (int j = 0; j < lines[secondEnter].Split(';').Length; j++)
                            {
                                TargetLines[k, j] = SepValues[j].Replace('"', ' ').Trim();
                                //Console.Write(TargetLines[k,j] +" ");
                            }
                            //Console.WriteLine();
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e.Message);
                        }
                    }

                    // проверка массива со значениями. Test
                    /*for (int i = 0; i < TargetLines.GetLength(0); i++)
                    {
                        for (int j = 0; j <= lines[secondEnter].Split(';').Length - 1; j++)
                            Console.Write(TargetLines[i, j] + " ");
                        Console.WriteLine();
                    }*/
                    /*Если имя компонента из ListKeyComponents совпадает с компонентом из TargetLines
                      то записываем в массив printValues название компонента TargetLines[j, 2] и его
                      массовое содержание TargetLines[j, 4]*/
                    for (int i = 0; i < ListKeyComponents.GetLength(0); i++)
                    {
                        try
                        {
                            for (int j = 0; j < TargetLines.GetLength(0); j++)
                            {
                                if (ListKeyComponents[i, 1] == TargetLines[j, 2])
                                {
                                    //Console.WriteLine(ListKeyComponents[i, 1] == TargetLines[j, 2]);
                                    printValues[i, 0] = TargetLines[j, 2];
                                    printValues[i, 1] = TargetLines[j, 3];
                                    break;
                                }
                                /*else
                                  {
                                  if (ListKeyComponents[i, 1] == TargetLines[j, 3])
                                    {
                                     printValues[i, 0] = TargetLines[j, 3];
                                    printValues[i, 1] = TargetLines[j, 10];
                                     break;
                                    }
                                  }*/
                            }
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e.Message + "Строка 257(поданый на вход массив меньше, чем эталлоный массив)");
                        }
                    }
                    /*если значение в массиве printValues = нулю, то записать в эту ячейку
                      название компонента из эталонного списка ListKeyComponents, и присвоить значение 0,000*/
                    for (int i = 0; i < ListKeyComponents.GetLength(0); i++)
                    {
                        try
                        {
                            if (printValues[i, 0] == null)
                            {
                                printValues[i, 0] = ListKeyComponents[i, 1];
                                printValues[i, 1] = "0,000";
                            }
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e.Message);
                        }
                    }
                    Console.ForegroundColor = ConsoleColor.White;
                    //Для наглядности, вывести данные из массива printValues на экран в виде таблицы
                    Console.WriteLine("---------------------------------------------------------------");
                    for (int i = 0; i < printValues.GetLength(0); i++)
                    {
                        try
                        {
                            for (int j = 0; j < printValues.GetLength(1); j++)
                            {
                                Console.Write("| {0,-28 } ", printValues[i, j]);
                            }
                            Console.Write("|");

                            Console.WriteLine();
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e.Message);
                        }

                    }
                    Console.Write("---------------------------------------------------------------");
                    Console.WriteLine();
                    Console.ForegroundColor = ConsoleColor.Red;
                    //составить путь к файлу использую переменную с номером опыта numberOfExp
                    string path = @"G:\Мой диск\Электронные журналы\" + nameOfPlant + "\\" + nameOfPlant + "-" + numberOfExp + " режимный лист.xlsm";

                    //string path = @"C:\Users\Хроматограф\Desktop\test.xlsx";

                    OpenExcelFile(path, titleOfChroma, printValues);
                    CloseAllExcelProcesses();
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("Конец программы");
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine("Для продолжения нажмите Enter...");

                    Console.ReadLine();
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                    Console.ForegroundColor = ConsoleColor.Blue;
                }
            }
        }
        static string OpenCSVFile()
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Откройте хроматограмму сохранённую в формате Excel2007");
            Console.ForegroundColor = ConsoleColor.Red;
            //Ожидание перед открытием 
            System.Threading.Thread.Sleep(1000);
            OpenFileDialog OFD = new OpenFileDialog();
            //запрет на открытие нескольких диалоговых окон
            OFD.Multiselect = false;
            OFD.Title = "Open Excel2007 Document";
            //фильтр файлов
            OFD.Filter = "Excel2007|*.xlsx";
            OFD.ShowDialog();
            // получаем путь до файла
            string filePath = OFD.FileName;
            return filePath;

        }
        static void OpenExcelFile(string path, string titleOfChroma, string[,] printValues)
        {
            Excel.Application xlApp = new Excel.Application();
            xlApp.DisplayAlerts = true;
            Excel.Workbooks xlWbks = xlApp.Workbooks;
            Excel.Workbook  xlWbk = xlWbks.Open(path, ReadOnly: false);
            //открываем 3й лист - Хроматограммы
            Excel.Sheets xlWrkshts = xlWbk.Sheets;
            int res_ind = 1;
            foreach (Excel.Worksheet sheet in xlWrkshts)
            {
                if (sheet.Name == "Хроматограммы")
                    break;
                //Console.WriteLine(sheet.Name);
                Marshal.ReleaseComObject(sheet);
                res_ind++;
            }

            Excel.Worksheet xlWrkSht = xlWbk.Sheets[res_ind];
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Рабочая вкладка " + xlWrkSht.Name);

            int startId = -1;
            int startWrite = -1;
            
            //Ищем в первом столбце ячейку = name, когда нашли присваиваем startId = i
            for (int i = 1; i < 50; i++)
            {
                try
                {
                    if (xlWrkSht.Cells[i, 1].Value == "name" || xlWrkSht.Cells[i, 1].Value == "Компонент")
                    {
                        startId = i;
                        break;
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                    Console.WriteLine("Название рабочего листа: " + xlWrkSht.Name + "(должно быть \"Хроматограммы\")");
                    //xlWbk.Close(true, Type.Missing, Type.Missing);
                    //xlApp.Quit();
                }
            }
            if (startId > 0)
                startWrite = startId + 1;
            else
                Console.WriteLine("Ошибка индекса поиска Name");

            // (xlWrkSht.Cells[startId, i].Value == null && xlWrkSht.Cells[startWrite, i].Value == null && xlWrkSht.Cells[startId, i - 1].Value != null)
            //if ((String.IsNullOrEmpty(xlWrkSht.Cells[startId, i].Value) | String.IsNullOrEmpty(xlWrkSht.Cells[startWrite, i].Value)) && (!String.IsNullOrEmpty(xlWrkSht.Cells[startId, i - 1].Value) && !String.IsNullOrEmpty(xlWrkSht.Cells[startWrite, i - 1].Value)))
            /*В условии проверяем является ли прошлая ячейка записанной а следующая пустой, если true - записываем сюда*/
            for (int i = 2; i < 1000; i++)
            {
                try
                {
                    if ((String.IsNullOrEmpty(Convert.ToString(xlWrkSht.Cells[startId, i].Value)) | String.IsNullOrEmpty(Convert.ToString(xlWrkSht.Cells[startWrite, i].Value))) && (!String.IsNullOrEmpty(Convert.ToString(xlWrkSht.Cells[startId, i - 1].Value)) && !String.IsNullOrEmpty(Convert.ToString(xlWrkSht.Cells[startWrite, i - 1].Value))))
                    {
                        for (int stWr = startWrite, targRow = 0; targRow < printValues.GetLength(0); stWr++, targRow++)
                        {
                            xlWrkSht.Cells[stWr, i].Value = Convert.ToDouble(printValues[targRow, 1]);
                        }
                        xlWrkSht.Cells[startId, i] = titleOfChroma;
                        break;
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                    xlWbk.Close(true, Type.Missing, Type.Missing);
                    xlApp.Quit();
                }
            }
            //Сохраняем документ и закрываем его. В случае отказа от сохранения сработает блок catch и файл всё равно закроется
            try
            {
                xlWbk.SaveAs(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
                xlWbk.Close(true, Type.Missing, Type.Missing);
                xlApp.Quit();
            }
            catch
            {
                xlWbk.Close(true, Type.Missing, Type.Missing);
                xlApp.Quit();
            }
            finally
            {
                if (xlApp != null) Marshal.ReleaseComObject(xlApp);
                if (xlWbks != null) Marshal.ReleaseComObject(xlWbks);
                if (xlWbk != null) Marshal.ReleaseComObject(xlWbk);
                if (xlWrkshts != null) Marshal.ReleaseComObject(xlWrkshts);
            }
            
        }
        static void OpenExcelComponents(string pathToList, out string[,] ListKeyComponents)
        {
            Excel.Application xlApp = new Excel.Application();
            //xlApp.DisplayAlerts = true;
            Excel.Workbooks xlWbks = xlApp.Workbooks;
            Excel.Workbook xlWbk = xlWbks.Open(pathToList, ReadOnly: false);
            //открываем 1й лист - Хроматограммы
                       
            Excel.Worksheet xlWrkSht = xlWbk.Sheets[1];
            //Excel.Sheets xlShts = xlWbk.Sheets;
            //Excel.Worksheet xlWrkSht = xlShts[1];
            int startId = 1;
            int sizeOfMass = -1;
            
            //определяем сколько строк используется для того, чтобы создать массив
            while (true)
            {
                if (xlWrkSht.Cells[startId, 1].Value != null && xlWrkSht.Cells[startId, 2].Value != null)
                {
                    startId++;
                }
                else
                    break;
            }

            sizeOfMass = startId - 1;
            ListKeyComponents = new string[sizeOfMass, 2];
            for (int i = 1; i <= ListKeyComponents.GetLength(0); i++)
            {
                try
                {
                    ListKeyComponents[i - 1, 0] = xlWrkSht.Cells[i, 1].Value.ToString();
                    ListKeyComponents[i - 1, 1] = xlWrkSht.Cells[i, 2].Value;
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                    xlWbk.Close(true, Type.Missing, Type.Missing);
                    xlApp.Quit();
                }
            }

            //test
            //Console.WriteLine(ListKeyComponents.GetLength(0));
            /*for (int i = 1; i <= ListKeyComponents.GetLength(0); i++)
            {
                Console.WriteLine(ListKeyComponents[i - 1, 1]);
            }*/
            //Закрываем документ и выходим из приложения
            // xlWbk.Close(true, Type.Missing, Type.Missing);

            //xlWrkSht.
            //xlWbk.SaveAs(pathToList, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
            xlWbk.Close();
            xlApp.Quit();
            
        }
        static void CloseAllExcelProcesses()
        {
            List<int> processesbeforegen = GetRunningProcessesExcel();
            List<int> processesaftergen = GetRunningProcessesExcel();
            KillProcesses(processesbeforegen, processesaftergen);
        }
        static List<int> GetRunningProcessesExcel()
        {
            List<int> ProcessIDs = new List<int>();
            //here we're going to get a list of all running processes on
            //the computer
            foreach (Process clsProcess in Process.GetProcesses())
            {
                if (Process.GetCurrentProcess().Id == clsProcess.Id)
                    continue;
                if (clsProcess.ProcessName.Contains("EXCEL"))
                {
                    ProcessIDs.Add(clsProcess.Id);
                }
            }
            return ProcessIDs;
        }
        static private void KillProcesses(List<int> processesbeforegen, List<int> processesaftergen)
        {
            foreach (int pidafter in processesaftergen)
            {
                bool processfound = false;
                foreach (int pidbefore in processesbeforegen)
                {
                    if (pidafter == pidbefore)
                    {
                        processfound = false;
                    }
                }

                if (processfound == false)
                {
                    Process clsProcess = Process.GetProcessById(pidafter);
                    clsProcess.Kill();
                }
            }
        }
    }
}

