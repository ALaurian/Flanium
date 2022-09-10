using System.Data;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using FlaUI.Core.AutomationElements;
using FlaUI.UIA3;
using Interop.UIAutomationClient;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using OpenQA.Selenium.Chrome;
using Application = Microsoft.Office.Interop.Excel.Application;
using Clipboard = System.Windows.Forms.Clipboard;
using DataColumn = System.Data.DataColumn;
using DataTable = System.Data.DataTable;
using Range = Microsoft.Office.Interop.Excel.Range;
using Window = FlaUI.Core.AutomationElements.Window;

#pragma warning disable CS0168

namespace Flanium;

[SuppressMessage("ReSharper", "AssignNullToNotNullAttribute")]
public class Blocks
{
    public class Engine
    {
        private Dictionary<object, object> Output { get; set; }
        private Func<object, object>[] _actions = Array.Empty<Func<object, object>>();
        private Dictionary<string, Func<string>[]> _transitionConditions = new();
        private Dictionary<string, bool> _continueOnError = new();
        private Dictionary<string, object> _parameters = new();
        private int Index { get; set; }
        private int OldIndex { get; set; }
        private bool IsRunning { get; set; }


        public Engine()
        {
        }

        [Obsolete("Engine constructor is obsolete due to the fact that you would not have access to GetOutput during runtime.",true)]
        public Engine(Func<object, object>[] actions,
            Dictionary<string, Func<string>[]> transitionConditions = null,
            Dictionary<string, bool> continueOnError = null)
        {
            _actions = actions;
            _transitionConditions = transitionConditions;
            _continueOnError = continueOnError;

        }

        public Dictionary<object, object> GetDictionary()
        {
            return Output;
        }

        public bool SetOutput(string actionName, object value)
        {
            Output[actionName] = value;
            return true;
        }
        public object GetOutput(string actionName)
        {
            return Output[actionName];
        }

        public Engine AddAction(Func<object, object> action)
        {
            if (IsRunning)
            {
                throw new Exception("Cannot add action while state machine is running.");
            }

            var actionsList = _actions.ToList();
            actionsList.Add(action);
            _actions = actionsList.ToArray();

            return this;
        }

        public Engine AddActions(Func<object, object>[] _actions)
        {
            if (IsRunning)
            {
                throw new Exception("Cannot add actions while state machine is running.");
            }

            var actionsList = this._actions.ToList();
            actionsList.AddRange(_actions);
            this._actions = actionsList.ToArray();

            return this;
        }

        public Engine RemoveAction(string actionName)
        {
            if (IsRunning)
            {
                throw new Exception("Cannot remove action while state machine is running.");
            }

            var actionsList = _actions.ToList();
            actionsList.RemoveAll(x => x.Method.GetParameters()[0].Name == actionName);
            _actions = actionsList.ToArray();

            return this;
        }

        public Engine RemoveActions(string[] actionNames)
        {
            if (IsRunning)
            {
                throw new Exception("Cannot remove actions while state machine is running.");
            }

            var actionsList = _actions.ToList();
            actionsList.RemoveAll(x => actionNames.Contains(x.Method.GetParameters()[0].Name));
            _actions = actionsList.ToArray();

            return this;
        }
        
        [Obsolete("AddCondition is obsolete due to the fact that you would not have access to GetOutput during runtime.",true)]
        public Engine AddCondition(string actionName, Func<string> condition)
        {
            if (IsRunning)
            {
                throw new Exception("Cannot add condition while state machine is running.");
            }

            _transitionConditions.Add(actionName, new[] {condition});

            return this;
        }
        
        [Obsolete("AddConditions is obsolete due to the fact that you would not have access to GetOutput during runtime.",true)]
        public Engine AddConditions(string[] actionNames, Func<string>[] conditions)
        {
            if (IsRunning)
            {
                throw new Exception("Cannot add conditions while state machine is running.");
            }

            for (var i = 0; i < actionNames.Length; i++)
            {
                _transitionConditions.Add(actionNames[i], new[] {conditions[i]});
            }

            return this;
        }

        public Engine RemoveCondition(string actionName)
        {
            if (IsRunning)
            {
                throw new Exception("Cannot remove condition while state machine is running.");
            }

            _transitionConditions.Remove(actionName);

            return this;
        }

        public Engine RemoveConditions(string[] actionNames)
        {
            if (IsRunning)
            {
                throw new Exception("Cannot remove conditions while state machine is running.");
            }

            foreach (var actionName in actionNames)
            {
                _transitionConditions.Remove(actionName);
            }

            return this;
        }

        public Engine AddContinueOnError(string actionName, bool continueOnError)
        {
            if (IsRunning)
            {
                throw new Exception("Cannot add continue on error while state machine is running.");
            }

            _continueOnError.Add(actionName, continueOnError);

            return this;
        }

        public Engine AddContinueOnErrors(string[] actionNames, bool[] continueOnErrors)
        {
            if (IsRunning)
            {
                throw new Exception("Cannot add continue on errors while state machine is running.");
            }

            for (var i = 0; i < actionNames.Length; i++)
            {
                _continueOnError.Add(actionNames[i], continueOnErrors[i]);
            }

            return this;
        }

        public Engine RemoveContinueOnError(string actionName)
        {
            if (IsRunning)
            {
                throw new Exception("Cannot remove continue on error while state machine is running.");
            }

            _continueOnError.Remove(actionName);
            return this;
            
        }

        public Engine RemoveContinueOnErrors(string[] actionNames)
        {
            if (IsRunning)
            {
                throw new Exception("Cannot remove continue on errors while state machine is running.");
            }

            foreach (var actionName in actionNames)
            {
                _continueOnError.Remove(actionName);
            }

            return this;
        }

        public object GetArgument(string keyName)
        {
            return _parameters[keyName];
        }
        
        public Engine SetArguments(Dictionary<string,object> arguments)
        {
            _parameters = arguments;
            return this;
        }

        public Engine ResetArguments()
        {
            _parameters = new Dictionary<string, object>();
            return this;
        }

        public Engine Stop()
        {
            IsRunning = false;
            OldIndex = Index;
            Index = _actions.Length + 1;
            return this;
        }

        public Engine Resume()
        {
            Index = OldIndex;
            Execute();
            return this;
        }

        public bool JumpTo(string actionName)
        {
            Stop();
            Index = Array.FindIndex(_actions, x => x.Method.GetParameters()[0].Name == actionName);
            Execute();
            return true;
        }

        public Engine Reset()
        {
            Index = 0;
            return this;
        }

        public Engine Execute()
        {
            if (_actions.Length == 0)
                throw new Exception("No actions to execute");
            if (IsRunning)
                throw new Exception("State machine is already running");

            IsRunning = true;
            Output = new Dictionary<object, object>();
            foreach (var action in _actions)
            {
                Output.Add(action.Method.GetParameters()[0].Name, null);
            }

            for (var executeIndex = Index; executeIndex < _actions.Length; executeIndex++)
            {
                stateGoto:
                var a = _actions[executeIndex];

                try
                {
                    var funcOutput = a.Invoke(a);
                    Output[a.Method.GetParameters()[0].Name] = funcOutput;
                    try
                    {
                        if (_transitionConditions[a.Method.GetParameters()[0].Name] != null)
                        {
                            foreach (var condition in _transitionConditions[a.Method.GetParameters()[0].Name])
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
                    if (_continueOnError != null)
                    {
                        try
                        {
                            if (_continueOnError[a.Method.GetParameters()[0].Name] == false)
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

            IsRunning = false;
            return this;
        }
    }

    public class ExcelEngine
    {
        private Application excelApp { get; set; }
        private Workbook workbook { get; set; }

        public Workbook GetWorkbook()
        {
            return workbook;
        }
        public ExcelEngine(string filePath)
        {
            try
            {
                excelApp = new()
                {
                    Visible = true,
                    DisplayAlerts = false
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
            return workbook.Worksheets[sheet].Range[range].Value2;
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

        public void InsertDataTable(object sheet, int row, DataTable dataTable, bool deleteEntireSheet = true, bool dataTableHeader = true)
        {
            workbook.Worksheets[sheet].Activate();
            //Delete all cells
            if (deleteEntireSheet)
            {
                workbook.Worksheets[sheet].Cells.Clear();
            }
            
            var columns = dataTable.Columns.Cast<DataColumn>();
            //	
            if (dataTableHeader)
            {
                var columnNames = columns.Select(column => column.ColumnName).ToArray();
                var joinedColumnNames = Strings.Join(columnNames,"	");

                Clipboard.SetText(joinedColumnNames);
                workbook.Worksheets[sheet].Cells[row, 1].Select();
                workbook.Worksheets[sheet].Paste();
                row++;
            }
            
            var rows = dataTable.Rows.Cast<DataRow>();
            var itemArrayJoiner = rows.Select(r => r.ItemArray).Aggregate("",
                (current, itemArray_original) => current +
                                                 (Strings.Join(
                                                     itemArray_original.Cast<string>()
                                                         .Select(item =>
                                                             item.ReplaceLineEndings("")
                                                                 .Replace("\t", ""))
                                                     .ToArray(), "	") + "\n"));
            Clipboard.SetText(itemArrayJoiner);
            workbook.Worksheets[sheet].Cells[row,1].Select();
            workbook.Worksheets[sheet].Paste();
        }

        public int GetLastRow(object sheet, int row, int column)
        {
            workbook.Worksheets[sheet].Cells[row, column].Select();
            excelApp.Selection.End(XlDirection.xlDown).Select();
            return excelApp.Selection.Row;
        }

        public bool CopyPaste(object sheet, string range, string destinationRange, XlPasteType pasteType)
        {
            workbook.Worksheets[sheet].Range(range).Copy();
            workbook.Worksheets[sheet].Range(destinationRange).PasteSpecial(pasteType);

            return true;
        }
        public bool AutoFill(object sheet, string range, int lastRow)
        {

            //split range by ":"
            var rangeArray = range.Split(':')[1];
            //replace all numbers in range
            var rangeArrayWithoutNumbers = Regex.Replace(rangeArray, @"\d", "");


            workbook.Worksheets[sheet].Range(range).AutoFill(workbook.Worksheets[sheet].Range(range.Split(':')[0] + ":" + rangeArrayWithoutNumbers + lastRow), XlAutoFillType.xlFillCopy);
            return true;
        }
        
        public DataTable ToDataTable(object sheet, int headerAt = 1)
        {            
            workbook.Worksheets[sheet].Activate();
            workbook.Worksheets[sheet].Cells[headerAt, 1].Select();
            while (excelApp.Selection.Value2 != null)
            {
                excelApp.Selection.End(XlDirection.xlToRight).Select();
            }

            excelApp.Selection.End(XlDirection.xlToLeft).Select();
            var lastColumnWithValue = excelApp.Selection.Column;

            var newDataTable = new DataTable();
            var headerRange = workbook.Worksheets[sheet].Range[workbook.Worksheets[sheet].Cells[headerAt, 1],
                workbook.Worksheets[sheet].Cells[headerAt, lastColumnWithValue]].Value2;
            var placeholderIndex = 0;
            
            foreach (string item in headerRange)
            {
                try
                {
                    var newColumn = new DataColumn(item,typeof(string));
                    newDataTable.Columns.Add(newColumn);
                }
                catch (Exception e)
                {
                    var newColumn = new DataColumn("Blank" + placeholderIndex,typeof(string));
                    newDataTable.Columns.Add(newColumn);
                }
            }

            workbook.Worksheets[sheet].Range[workbook.Worksheets[sheet].Cells[headerAt, 1],
                workbook.Worksheets[sheet].Cells[headerAt, lastColumnWithValue]].Select();
            workbook.Worksheets[sheet].Range[excelApp.Selection,excelApp.Selection.End(XlDirection.xlDown)].Select();
            excelApp.Selection.Offset(1, 0).Select();
            excelApp.Selection.Resize[excelApp.Selection.Rows.Count-1, excelApp.Selection.Columns.Count].Select();
            var selectionValue = ((object[,]) excelApp.Selection.Value2).GetEnumerator();

            selectionValue.MoveNext();
            try
            {
                while (selectionValue.Current != null)
                {
                    var newRow = newDataTable.NewRow();
                    for (var i = 0; i < lastColumnWithValue; i++)
                    {
                        
                        if(selectionValue.Current!= null)
                            newRow[i] = selectionValue.Current.ToString();

                        selectionValue.MoveNext();
                    }

                    newRow.ItemArray = newRow.ItemArray.Select(x =>
                    {
                        if (x == DBNull.Value)
                        {
                            x = "";
                        }

                        return x;
                    }).ToArray();
                    
                    newDataTable.Rows.Add(newRow);
                }
            }
            catch
            {

            }
            
            return newDataTable;
        }


        public string Close()
        {
            Retry:
            try
            {
                if (excelApp == null && workbook == null)
                {
                    goto FinishLine;
                }
                
                workbook.Close(true);
                excelApp.Quit();
                
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(excelApp);
                

            }
            catch
            {
                goto Retry;
            }

            FinishLine:
            return "Closed Excel.";
        }
        
        public string SaveAs(string filePath, XlFileFormat format)
        {

            
            workbook.SaveAs(filePath, format);


            return "Saved Excel as " + filePath + " with format" + format + ".";
        }

        public string Save()
        {
            var window = WinEvents.Search.GetWindow("*[contains(@Name,'" + workbook.Name + "')]");
            var saveButton = WinEvents.Search.FindElement(window, "//*[@Name='Save']");
            WinEvents.Action.Click(saveButton, true);
            
            return "Saved Excel file.";
        }
    }

    public class DataTableEngine
    {
        private DataTable _dataTable { get; set; }

        public string Cells(int row, object column, string value)
        {
            if (column is int)
                _dataTable.Rows[row][(int)column] = value;
            else
                _dataTable.Rows[row][column.ToString()] = value;
            return "Edited cell on row " + row + " and column " + column + " to " + value + ".";
        }
    
        public string Cells(int row, object column)
        {
            if (column is int)
                return _dataTable.Rows[row][(int) column].ToString();
            else
                return _dataTable.Rows[row][column.ToString()].ToString();
        }

        public DataRow Row(int index)
        {
            return _dataTable.Rows[index];
        }

        public string UpdateRow(int index, DataRow row)
        {
            _dataTable.Rows[index].ItemArray = row.ItemArray;
            return "Updated row at " + index + ".";
        }
    
        public string DeleteRow(int index)
        {
            _dataTable.Rows[index].Delete();
            return "Deleted row at " + index + ".";
        }
        
        public DataTableEngine(DataTable dataTable)
        {
            _dataTable = dataTable;
        }

        public DataTableEngine Filter(Func<DataRow, bool> linQFunction)
        {
            var filteredDataTableRows = _dataTable.Rows.Cast<DataRow>()
                .Where(linQFunction);

            var newDataTable = new DataTable();
            foreach (var c in _dataTable.Columns.Cast<DataColumn>())
            {
                newDataTable.Columns.Add(c.ColumnName);
            }

            foreach (var row in filteredDataTableRows)
            {
                var newRow = newDataTable.NewRow();
                newRow.ItemArray = row.ItemArray;
                newDataTable.Rows.Add(newRow);
            }

            _dataTable = null;
            _dataTable = newDataTable;
            return this;
        }
        
        public DataTableEngine ForEach(Func<DataRow, DataRow> linQFunction)
        {
            var filteredDataTableRows = _dataTable.Rows.Cast<DataRow>()
                .Select(linQFunction);

            var newDataTable = new DataTable();
            foreach (var c in _dataTable.Columns.Cast<DataColumn>())
            {
                newDataTable.Columns.Add(c.ColumnName);
            }

            foreach (var row in filteredDataTableRows)
            {
                var newRow = newDataTable.NewRow();
                newRow.ItemArray = row.ItemArray;
                newDataTable.Rows.Add(newRow);
            }

            _dataTable = null;
            _dataTable = newDataTable;
            return this;
        }

        public DataTable GetDataTable()
        {
            return _dataTable;
        }
    }
}