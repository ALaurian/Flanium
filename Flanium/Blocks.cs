using System.Data;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Windows.UI.Composition;
using Microsoft.Office.Interop.Excel;
using OpenQA.Selenium.Chrome;
using DataTable = System.Data.DataTable;
using Range = System.Range;

namespace Flanium;

public class Blocks
{
    public class Engine
    {
        private Dictionary<object, object> Output { get; set; }
        private Func<object, object>[] actions = Array.Empty<Func<object, object>>();
        private Dictionary<string, Func<string>[]> TransitionConditions = new();
        private Dictionary<string, bool> ContinueOnError = new();
        private int index { get; set; }
        private int oldIndex { get; set; }
        private bool isRunning { get; set; }


        public Engine()
        {
        }

        public Engine(Func<object, object>[] _actions,
            Dictionary<string, Func<string>[]> _TransitionConditions = null,
            Dictionary<string, bool> _ContinueOnError = null)
        {
            actions = _actions;
            TransitionConditions = _TransitionConditions;
            ContinueOnError = _ContinueOnError;

            var machine = new Engine(new Func<object, object>[]
            {
                (myAction) => WebEvents.Action.Click(new ChromeDriver(), "")
            });
            
            machine.AddCondition("myAction", () => "myAction");
        }

        public Dictionary<object, object> GetDictionary()
        {
            return Output;
        }

        public object GetOutput(string actionName)
        {
            return Output[actionName];
        }

        public Engine AddAction(Func<object, object> action)
        {
            if (isRunning)
            {
                throw new Exception("Cannot add action while state machine is running.");
            }

            var actionsList = actions.ToList();
            actionsList.Add(action);
            actions = actionsList.ToArray();

            return this;
        }

        public Engine AddActions(Func<object, object>[] _actions)
        {
            if (isRunning)
            {
                throw new Exception("Cannot add actions while state machine is running.");
            }

            var actionsList = actions.ToList();
            actionsList.AddRange(_actions);
            actions = actionsList.ToArray();

            return this;
        }

        public Engine RemoveAction(string actionName)
        {
            if (isRunning)
            {
                throw new Exception("Cannot remove action while state machine is running.");
            }

            var actionsList = actions.ToList();
            actionsList.RemoveAll(x => x.Method.GetParameters()[0].Name == actionName);
            actions = actionsList.ToArray();

            return this;
        }

        public Engine RemoveActions(string[] actionNames)
        {
            if (isRunning)
            {
                throw new Exception("Cannot remove actions while state machine is running.");
            }

            var actionsList = actions.ToList();
            actionsList.RemoveAll(x => actionNames.Contains(x.Method.GetParameters()[0].Name));
            actions = actionsList.ToArray();

            return this;
        }

        public Engine AddCondition(string actionName, Func<string> condition)
        {
            if (isRunning)
            {
                throw new Exception("Cannot add condition while state machine is running.");
            }

            TransitionConditions.Add(actionName, new[] {condition});

            return this;
        }

        public Engine AddConditions(string[] actionNames, Func<string>[] conditions)
        {
            if (isRunning)
            {
                throw new Exception("Cannot add conditions while state machine is running.");
            }

            for (var i = 0; i < actionNames.Length; i++)
            {
                TransitionConditions.Add(actionNames[i], new[] {conditions[i]});
            }

            return this;
        }

        public Engine RemoveCondition(string actionName)
        {
            if (isRunning)
            {
                throw new Exception("Cannot remove condition while state machine is running.");
            }

            TransitionConditions.Remove(actionName);

            return this;
        }

        public Engine RemoveConditions(string[] actionNames)
        {
            if (isRunning)
            {
                throw new Exception("Cannot remove conditions while state machine is running.");
            }

            foreach (var actionName in actionNames)
            {
                TransitionConditions.Remove(actionName);
            }

            return this;
        }

        public Engine AddContinueOnError(string actionName, bool continueOnError)
        {
            if (isRunning)
            {
                throw new Exception("Cannot add continue on error while state machine is running.");
            }

            ContinueOnError.Add(actionName, continueOnError);

            return this;
        }

        public Engine AddContinueOnErrors(string[] actionNames, bool[] continueOnErrors)
        {
            if (isRunning)
            {
                throw new Exception("Cannot add continue on errors while state machine is running.");
            }

            for (var i = 0; i < actionNames.Length; i++)
            {
                ContinueOnError.Add(actionNames[i], continueOnErrors[i]);
            }

            return this;
        }

        public Engine RemoveContinueOnError(string actionName)
        {
            if (isRunning)
            {
                throw new Exception("Cannot remove continue on error while state machine is running.");
            }
            
            var engine = new Engine(new Func<object, object>[]
                {
                    (myAction) => WebEvents.Action.Click(new ChromeDriver(), "")
                }, new Dictionary<string, Func<string>[]>()
                {
                    ["myAction"] = new Func<string>[] {() => "myAction"}
                },
                new Dictionary<string, bool>()
                {
                    {"myAction", true}
                });
            
            ContinueOnError.Remove(actionName);
            return this;



        }

        public Engine RemoveContinueOnErrors(string[] actionNames)
        {
            if (isRunning)
            {
                throw new Exception("Cannot remove continue on errors while state machine is running.");
            }

            foreach (var actionName in actionNames)
            {
                ContinueOnError.Remove(actionName);
            }

            return this;
        }

        public bool Stop()
        {
            oldIndex = index;
            index = actions.Length + 1;
            return true;
        }

        public bool Resume()
        {
            index = oldIndex;
            Execute();
            return true;
        }

        public string JumpTo(string actionName)
        {
            index = Array.FindIndex(actions, x => x.Method.GetParameters()[0].Name == actionName);

            return "Jumped to " + actionName;
        }

        public Engine Execute()
        {
            if (actions.Length == 0)
                throw new Exception("No actions to execute");
            if (isRunning)
                throw new Exception("State machine is already running");

            isRunning = true;
            Output = new Dictionary<object, object>();
            foreach (var action in actions)
            {
                Output.Add(action.Method.GetParameters()[0].Name, null);
            }

            for (var executeIndex = index; executeIndex < actions.Length; executeIndex++)
            {
                stateGoto:
                var a = actions[executeIndex];

                try
                {
                    var funcOutput = a.Invoke(a);
                    Output[a.Method.GetParameters()[0].Name] = funcOutput;
                    try
                    {
                        if (TransitionConditions[a.Method.GetParameters()[0].Name] != null)
                        {
                            foreach (var condition in TransitionConditions[a.Method.GetParameters()[0].Name])
                            {
                                var c = condition.Invoke();
                                if (c != null)
                                {
                                    executeIndex = Output.Keys.ToList().IndexOf(c);
                                    goto stateGoto;
                                }
                            }
                        }
                    }
                    catch
                    {
                    }
                }
                catch (Exception e)
                {
                    Output[a.Method.GetParameters()[0].Name] = e;
                    if (ContinueOnError != null)
                    {
                        try
                        {
                            if (ContinueOnError[a.Method.GetParameters()[0].Name] == false)
                            {
                                throw;
                            }
                        }
                        catch
                        {
                            // ignored
                        }
                    }
                }
            }

            isRunning = false;
            return this;
        }
    }

    public class ExcelEngine
    {
        private Application excelApp { get; set; }
        private Workbook workbook { get; set; }

        public ExcelEngine(string filePath)
        {
            try
            {
                excelApp = new()
                {
                    Visible = true
                };
                try
                {
                    workbook.Close();
                }
                catch
                {
                    // ignored
                }

                workbook = excelApp.Workbooks.Open(filePath);
            }
            catch (Exception e)
            {
                throw new Exception("Error occurred while opening excel file.", e);
            }
        }

        public object GetCellValue(object sheet, int row, int column)
        {
            return workbook.Worksheets[sheet].Cells[row, column].Value2;
        }

        public object GetRangeValue(object sheet, string range)
        {
            if(range.Contains(":"))
                return (workbook.Worksheets[sheet].Range[range].Value2 as object[,]).Cast<object>().ToList();
            
            return (workbook.Worksheets[sheet].Range[range].Value2 as object[,]).Cast<object>().First();
        }

        public object GetRowValue(object sheet, int index)
        {
            return (workbook.Worksheets[sheet].Rows[index].Value2 as object[,]).Cast<object>().ToList();
        }

        public object GetColumnValue(object sheet, int index)
        {
            return (workbook.Worksheets[sheet].Columns[index].Value2 as object[,]).Cast<object>().ToList();
        }

        public string EditCell(object sheet, int row, int column, string value)
        {
            workbook.Worksheets[sheet].Cells[row, column].Value2 = value;
            return "Edited cell on row " + row + " and column " + column + " to " + value + ".";
        }

        public string EditRange(object sheet, string range, string value)
        {
            workbook.Worksheets[sheet].Range[range].Value2 = value;
            return "Edited range " + range + " to " + value + ".";
        }

        public string DeleteCell(object sheet, int row, int column, XlDeleteShiftDirection direction)
        {
            workbook.Worksheets[sheet].Cells[row, column].Delete(direction);
            return "Deleted cell on row " + row + " and column " + column + ".";
        }

        public string DeleteRange(object sheet, string range, XlDeleteShiftDirection direction)
        {
            workbook.Worksheets[sheet].Range[range].Delete(direction);
            return "Deleted range " + range + ".";
        }

        public string DeleteRow(object sheet, int row, XlDeleteShiftDirection direction)
        {
            workbook.Worksheets[sheet].Rows[row].EntireRow.Delete(direction);
            return "Deleted row " + row + ".";
        }

        public string DeleteColumn(object sheet, int column, XlDeleteShiftDirection direction)
        {
            workbook.Worksheets[sheet].Columns[column].EntireColumn.Delete(direction);
            return "Deleted column " + column + ".";
        }

        public string InsertRow(object sheet, int index)
        {
            var line = workbook.Worksheets[sheet].Rows[index];
            line.Insert();
            return "Inserted new row at" + index + ".";
        }

        public string InsertColumn(object sheet, int index)
        {
            var line = workbook.Worksheets[sheet].Columns[index];
            line.Insert();
            return "Inserted new column at" + index + ".";
        }

        public DataTable ToDataTable(object sheet, int headerAt = 1)
        {
            var range = workbook.Worksheets[sheet].UsedRange;

            System.Data.DataTable dt = new System.Data.DataTable();
            var headerRow = range?.Rows[headerAt].Value2;

            var placeholderIndex = 0;
            foreach (object item in headerRow)
            {
                if (item != null)
                {
                    dt.Columns.Add(item.ToString());
                }
                else
                {
                    dt.Columns.Add($"Placeholder_{placeholderIndex}");
                    placeholderIndex++;
                }
            }

            var rowsCount = range.Rows.Count;
            for (int i = headerAt + 1; i <= rowsCount; i++)
            {
                var rowValues = range.Rows[i].Value2;
                List<string> row = new List<string>();

                foreach (object v in rowValues)
                {
                    row.Add(v != null ? v.ToString() : "");
                }

                dt.Rows.Add(row.ToArray());
            }

            return dt;
        }


        public string Close()
        {
            workbook.Close();
            excelApp.Quit();
            
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(excelApp);
            
            return "Closed Excel.";
        }

        public string SaveAs(string filePath, XlFileFormat format)
        {
            workbook.SaveAs(filePath, format);
            return "Saved Excel as " + filePath + " with format" + format + ".";
        }
    }

    public class DataTableEngine
    {
        private Dictionary<object, object> Output { get; set; }
        private Func<object, object>[] actions = Array.Empty<Func<object, object>>();
        private Dictionary<string, Func<string>[]> TransitionConditions = new();
        private Dictionary<string, bool> ContinueOnError = new();
        private int index { get; set; }
        private int oldIndex { get; set; }
        private bool isRunning { get; set; }

        private DataTable dataTable { get; set; }

        public string Cells(int row, object column, string value)
        {
            if (column is int)
                dataTable.Rows[row][(int)column] = value;
            else
                dataTable.Rows[row][column.ToString()] = value;
            return "Edited cell on row " + row + " and column " + column + " to " + value + ".";
        }
    
        public string Cells(int row, object column)
        {
            if (column is int)
                return dataTable.Rows[row][(int) column].ToString();
            else
                return dataTable.Rows[row][column.ToString()].ToString();
        }

        public DataRow Row(int index)
        {
            return dataTable.Rows[index];
        }

        public string UpdateRow(int index, DataRow row)
        {
            dataTable.Rows[index].ItemArray = row.ItemArray;
            return "Updated row at " + index + ".";
        }
    
        public string DeleteRow(int index)
        {
            dataTable.Rows[index].Delete();
            return "Deleted row at " + index + ".";
        }
        
        public DataTableEngine(DataTable _dataTable)
        {
            dataTable = _dataTable;
        }
    }
}