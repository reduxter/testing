
Sub TraceFormulaForRange()
    Dim rng As Range
    Dim cell As Range
    Dim traceSheet As Worksheet
    Dim logRow As Long

    ' Set up the trace log sheet
    On Error Resume Next
    Set traceSheet = Worksheets("Trace Log")
    If traceSheet Is Nothing Then
        Set traceSheet = ThisWorkbook.Worksheets.Add
        traceSheet.Name = "Trace Log"
    End If
    On Error GoTo 0
    traceSheet.Cells.Clear
    traceSheet.Range("A1:F1").Value = Array("Workbook", "Sheet", "Cell", "Formula", "Value", "File Path")
    logRow = 2

    ' Select the range to process
    Set rng = Application.InputBox("Select the range to trace formulas:", Type:=8)

    ' Loop through each cell in the selected range
    For Each cell In rng
        If cell.HasFormula Then
            Call TraceCellRecursive(cell, traceSheet, logRow, rng)
        End If
    Next cell

    MsgBox "Tracing completed. Check the 'Trace Log' sheet for details.", vbInformation
End Sub

Sub TraceCellRecursive(ByVal targetCell As Range, ByVal traceSheet As Worksheet, ByRef logRow As Long, ByVal selectedRange As Range)
    Dim precedent As Range
    Dim precedentsRange As Range
    Dim externalWorkbook As Workbook
    Dim externalSheet As Worksheet
    Dim currentWorkbook As String, currentSheet As String, currentCell As String
    Dim currentFormula As String, currentValue As String
    Dim externalFilePath As String, externalSheetName As String, externalCellAddress As String
    Dim retryAttempts As Integer
    Dim fileOpened As Boolean
    Dim linkSources As Variant

    ' Log the current cell details
    currentWorkbook = targetCell.Parent.Parent.Name
    currentSheet = targetCell.Parent.Name
    currentCell = targetCell.Address
    currentFormula = Replace(targetCell.Formula, "=", "", 1, 1) ' Remove the "=" sign
    currentValue = targetCell.Value

    traceSheet.Cells(logRow, 1).Value = currentWorkbook
    traceSheet.Cells(logRow, 2).Value = currentSheet
    traceSheet.Cells(logRow, 3).Value = currentCell
    traceSheet.Cells(logRow, 4).Value = currentFormula
    traceSheet.Cells(logRow, 5).Value = currentValue
    traceSheet.Cells(logRow, 6).Value = "N/A"
    logRow = logRow + 1

    ' Stop tracing if the cell contains a plain value
    If Not targetCell.HasFormula Then Exit Sub

    ' Get the precedents for this cell (internal references)
    On Error Resume Next
    Set precedentsRange = targetCell.Precedents
    On Error GoTo 0

    If Not precedentsRange Is Nothing Then
        ' Loop through all precedents and recursively trace them
        For Each precedent In precedentsRange
            If Not Intersect(precedent, selectedRange) Is Nothing Then
                Call TraceCellRecursive(precedent, traceSheet, logRow, selectedRange)
            Else
                Call TraceExternalReference(precedent, traceSheet, logRow)
            End If
        Next precedent
    End If
End Sub




--------------------------------------

Sub TraceFormulaForRange()
    Dim rng As Range
    Dim cell As Range
    Dim traceSheet As Worksheet
    Dim logRow As Long

    ' Set up the trace log sheet
    On Error Resume Next
    Set traceSheet = Worksheets("Trace Log")
    If traceSheet Is Nothing Then
        Set traceSheet = ThisWorkbook.Worksheets.Add
        traceSheet.Name = "Trace Log"
    End If
    On Error GoTo 0
    traceSheet.Cells.Clear
    traceSheet.Range("A1:F1").Value = Array("Workbook", "Sheet", "Cell", "Formula", "Value", "File Path")
    logRow = 2

    ' Select the range to process
    Set rng = Application.InputBox("Select the range to trace formulas:", Type:=8)

    ' Loop through each cell in the selected range
    For Each cell In rng
        If cell.HasFormula Then
            Call TraceCellRecursive(cell, traceSheet, logRow, rng)
        End If
    Next cell

    MsgBox "Tracing completed. Check the 'Trace Log' sheet for details.", vbInformation
End Sub

Sub TraceCellRecursive(ByVal targetCell As Range, ByVal traceSheet As Worksheet, ByRef logRow As Long, ByVal selectedRange As Range)
    Dim precedent As Range
    Dim precedentsRange As Range
    Dim externalWorkbook As Workbook
    Dim externalSheet As Worksheet
    Dim currentWorkbook As String, currentSheet As String, currentCell As String
    Dim currentFormula As String, currentValue As String
    Dim externalFilePath As String, externalSheetName As String, externalCellAddress As String
    Dim retryAttempts As Integer
    Dim fileOpened As Boolean
    Dim linkSources As Variant

    ' Log the current cell details
    currentWorkbook = targetCell.Parent.Parent.Name
    currentSheet = targetCell.Parent.Name
    currentCell = targetCell.Address
    currentFormula = targetCell.Formula
    currentValue = targetCell.Value

    traceSheet.Cells(logRow, 1).Value = currentWorkbook
    traceSheet.Cells(logRow, 2).Value = currentSheet
    traceSheet.Cells(logRow, 3).Value = currentCell
    traceSheet.Cells(logRow, 4).Value = currentFormula
    traceSheet.Cells(logRow, 5).Value = currentValue
    traceSheet.Cells(logRow, 6).Value = "N/A"
    logRow = logRow + 1

    ' Stop tracing if the cell contains a plain value
    If Not targetCell.HasFormula Then Exit Sub

    ' Get the precedents for this cell (internal references)
    On Error Resume Next
    Set precedentsRange = targetCell.Precedents
    On Error GoTo 0

    If Not precedentsRange Is Nothing Then
        ' Loop through all precedents and recursively trace them
        For Each precedent In precedentsRange
            If Not Intersect(precedent, selectedRange) Is Nothing Then
                Call TraceCellRecursive(precedent, traceSheet, logRow, selectedRange)
            Else
                Call TraceExternalReference(precedent, traceSheet, logRow)
            End If
        Next precedent
    End If
End Sub

Sub TraceExternalReference(ByVal targetCell As Range, ByVal traceSheet As Worksheet, ByRef logRow As Long)
    Dim externalWorkbook As Workbook
    Dim externalSheet As Worksheet
    Dim externalFilePath As String, externalSheetName As String, externalCellAddress As String
    Dim retryAttempts As Integer
    Dim fileOpened As Boolean

    ' Extract external reference details
    ExtractExternalReference targetCell.Formula, externalFilePath, externalSheetName, externalCellAddress
    externalFilePath = SanitizeFilePath(externalFilePath)

    ' Check if the external workbook is already open
    On Error Resume Next
    Set externalWorkbook = Workbooks(externalFilePath)
    On Error GoTo 0

    ' Attempt to open the external workbook if it is not already open
    If externalWorkbook Is Nothing Then
        retryAttempts = 0
        fileOpened = False

        Do While retryAttempts < 5 And Not fileOpened
            retryAttempts = retryAttempts + 1
            On Error Resume Next
            Set externalWorkbook = Workbooks.Open(externalFilePath, ReadOnly:=True)
            On Error GoTo 0

            If Not externalWorkbook Is Nothing Then
                fileOpened = True
                traceSheet.Cells(logRow - 1, 6).Value = externalFilePath
            Else
                Application.Wait Now + TimeValue("00:00:05")
            End If
        Loop

        ' Log error if unable to open the file
        If Not fileOpened Then
            traceSheet.Cells(logRow - 1, 5).Value = "Error: Unable to open external workbook"
            traceSheet.Cells(logRow - 1, 6).Value = externalFilePath
            Exit Sub
        End If
    End If

    ' Access the external sheet and recursively trace the referenced cell
    Set externalSheet = externalWorkbook.Sheets(externalSheetName)
    Dim nextCell As Range
    Set nextCell = externalSheet.Range(externalCellAddress)

    ' Continue tracing, including internal references within the external workbook
    Call TraceCellRecursive(nextCell, traceSheet, logRow, externalSheet.UsedRange)
End Sub


---------------


import win32com.client as win32
import openpyxl
import time
import sys

def trace_formula_in_range(excel_app, selected_range):
    trace_log = []
    visited_cells = set()  # Track visited cells to avoid infinite loops

    # Iterate through each cell in the selected range using an iterative approach
    stack = [(cell, excel_app.ActiveWorkbook.Name, excel_app.ActiveSheet.Name) for cell in selected_range if cell.HasFormula]

    while stack:
        target_cell, workbook_name, sheet_name = stack.pop()
        cell_address = target_cell.Address
        formula = target_cell.Formula[1:]  # Remove the "=" sign
        value = target_cell.Value

        log_entry = [workbook_name, sheet_name, cell_address, formula, value, "N/A"]
        trace_log.append(log_entry)

        # Skip if the cell has already been visited
        cell_key = (workbook_name, sheet_name, cell_address)
        if cell_key in visited_cells:
            continue
        visited_cells.add(cell_key)

        # If the cell has no formula, stop tracing
        if not target_cell.HasFormula:
            continue

        # Get the precedents for internal references
        try:
            precedents = target_cell.Precedents
        except Exception:
            precedents = None

        if precedents is not None:
            # Add internal precedents to the stack
            for precedent in precedents:
                stack.append((precedent, workbook_name, sheet_name))
        else:
            # Handle external workbook references
            try:
                external_file_path, external_sheet_name, external_cell_address = extract_external_reference(formula)
                external_file_path = resolve_full_path(excel_app, external_file_path)

                # Attempt to open the external workbook if not already open
                try:
                    external_workbook = excel_app.Workbooks(external_file_path)
                except Exception:
                    external_workbook = None

                if external_workbook is None:
                    for attempt in range(5):
                        try:
                            external_workbook = excel_app.Workbooks.Open(external_file_path, ReadOnly=True)
                            time.sleep(2)  # Increase wait time
                            break
                        except Exception as e:
                            print(f"Attempt {attempt + 1}: Unable to open {external_file_path} - {str(e)}")
                            time.sleep(5)

                if external_workbook is not None:
                    external_sheet = external_workbook.Sheets(external_sheet_name)
                    next_cell = external_sheet.Range(external_cell_address)
                    stack.append((next_cell, external_workbook.Name, external_sheet.Name))
                else:
                    log_entry[5] = f"Error: Unable to open external workbook {external_file_path}"
            except Exception as e:
                log_entry[5] = f"Error tracing external reference: {str(e)}"

    # Write the trace log to a new Excel file
    write_trace_log(trace_log)

def extract_external_reference(formula):
    try:
        start_pos = formula.index("'[") + 2
        end_pos = formula.index("]")
        file_path = formula[start_pos:end_pos]

        start_pos = end_pos + 2
        end_pos = formula.index("'", start_pos)
        sheet_name = formula[start_pos:end_pos]

        cell_address = formula.split("!")[-1]

        return file_path, sheet_name, cell_address
    except Exception as e:
        print(f"Error extracting reference: {str(e)}")
        return "", "", ""

def resolve_full_path(excel_app, partial_path):
    try:
        link_sources = excel_app.ActiveWorkbook.LinkSources(1)  # xlExcelLinks = 1
        if link_sources:
            for full_path in link_sources:
                if partial_path in full_path:
                    return full_path
        return partial_path
    except Exception as e:
        print(f"Error resolving full path: {str(e)}")
        return partial_path

def write_trace_log(trace_log):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Trace Log"

    headers = ["Workbook", "Sheet", "Cell", "Formula", "Value", "File Path"]
    ws.append(headers)

    for row in trace_log:
        ws.append(row)

    wb.save("Trace_Log.xlsx")
    print("Trace log saved as 'Trace_Log.xlsx'.")

if __name__ == "__main__":
    try:
        excel_app = win32.Dispatch("Excel.Application")
        excel_app.Visible = True

        active_workbook = excel_app.ActiveWorkbook
        selected_range = excel_app.Selection

        trace_formula_in_range(excel_app, selected_range)
    except Exception as e:
        print(f"Error: {str(e)}")
        sys.exit(1)
    finally:
        excel_app.Quit()
------------------


import win32com.client as win32
import openpyxl
import time
import sys

# Function to trace formulas in the selected range
def trace_formula_in_range(excel_app, selected_range):
    trace_log = []
    visited_cells = set()  # Track visited cells to avoid infinite loops

    # Iterate through each cell in the selected range using an iterative approach
    stack = [(cell, excel_app.ActiveWorkbook.Name, excel_app.ActiveSheet.Name) for cell in selected_range if cell.HasFormula]

    while stack:
        target_cell, workbook_name, sheet_name = stack.pop()
        cell_address = target_cell.Address
        formula = target_cell.Formula[1:]  # Remove the "=" sign
        value = target_cell.Value

        log_entry = [workbook_name, sheet_name, cell_address, formula, value, "N/A"]
        trace_log.append(log_entry)

        # Skip if the cell has already been visited
        cell_key = (workbook_name, sheet_name, cell_address)
        if cell_key in visited_cells:
            continue
        visited_cells.add(cell_key)

        # If the cell has no formula, stop tracing
        if not target_cell.HasFormula:
            continue

        # Get the precedents for internal references
        try:
            precedents = target_cell.Precedents
        except Exception:
            precedents = None

        if precedents is not None:
            # Add internal precedents to the stack
            for precedent in precedents:
                stack.append((precedent, workbook_name, sheet_name))
        else:
            # Handle external workbook references
            try:
                external_file_path, external_sheet_name, external_cell_address = extract_external_reference(formula)
                external_file_path = resolve_full_path(excel_app, external_file_path)

                # Attempt to open the external workbook if not already open
                try:
                    external_workbook = excel_app.Workbooks(external_file_path)
                except Exception:
                    external_workbook = None

                if external_workbook is None:
                    for attempt in range(5):
                        try:
                            external_workbook = excel_app.Workbooks.Open(external_file_path, ReadOnly=True)
                            time.sleep(2)  # Increase wait time
                            break
                        except Exception as e:
                            print(f"Attempt {attempt + 1}: Unable to open {external_file_path} - {str(e)}")
                            time.sleep(5)

                if external_workbook is not None:
                    external_sheet = external_workbook.Sheets(external_sheet_name)
                    next_cell = external_sheet.Range(external_cell_address)
                    stack.append((next_cell, external_workbook.Name, external_sheet.Name))
                else:
                    log_entry[5] = f"Error: Unable to open external workbook {external_file_path}"
            except Exception as e:
                log_entry[5] = f"Error tracing external reference: {str(e)}"

    # Write the trace log to a new Excel file
    write_trace_log(trace_log)

# Function to extract external reference details from a formula
def extract_external_reference(formula):
    try:
        start_pos = formula.index("'[") + 2
        end_pos = formula.index("]")
        file_path = formula[start_pos:end_pos]

        start_pos = end_pos + 2
        end_pos = formula.index("'", start_pos)
        sheet_name = formula[start_pos:end_pos]

        cell_address = formula.split("!")[-1]

        return file_path, sheet_name, cell_address
    except Exception as e:
        print(f"Error extracting reference: {str(e)}")
        return "", "", ""

# Function to resolve the full path of an external workbook
def resolve_full_path(excel_app, partial_path):
    try:
        link_sources = excel_app.ActiveWorkbook.LinkSources(1)  # xlExcelLinks = 1
        if link_sources:
            for full_path in link_sources:
                if partial_path in full_path:
                    return full_path
        return partial_path
    except Exception as e:
        print(f"Error resolving full path: {str(e)}")
        return partial_path

# Function to write the trace log to an Excel file
def write_trace_log(trace_log):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Trace Log"

    headers = ["Workbook", "Sheet", "Cell", "Formula", "Value", "File Path"]
    ws.append(headers)

    for row in trace_log:
        ws.append(row)

    wb.save("Trace_Log.xlsx")
    print("Trace log saved as 'Trace_Log.xlsx'.")

# Script execution starts here
try:
    # Launch Excel application
    excel_app = win32.Dispatch("Excel.Application")
    excel_app.Visible = True

    # Get the active workbook and selected range
    active_workbook = excel_app.ActiveWorkbook
    selected_range = excel_app.Selection

    # Trace the formulas in the selected range
    trace_formula_in_range(excel_app, selected_range)

except Exception as e:
    print(f"Error: {str(e)}")
    sys.exit(1)

finally:
    # Quit Excel application
    excel_app.Quit()


-------------

import win32com.client as win32
import openpyxl
import time
import sys

# Function to trace formulas in the specified range
def trace_formula_in_range(excel_app, workbook_path, sheet_name, cell_range):
    trace_log = []
    visited_cells = set()  # Track visited cells to avoid infinite loops

    try:
        # Open the specified workbook
        workbook = excel_app.Workbooks.Open(workbook_path)
        sheet = workbook.Sheets(sheet_name)
        selected_range = sheet.Range(cell_range)

        # Iterate through each cell in the specified range using an iterative approach
        stack = [(cell, workbook.Name, sheet.Name) for cell in selected_range if cell.HasFormula]

        while stack:
            target_cell, workbook_name, sheet_name = stack.pop()
            cell_address = target_cell.Address
            formula = target_cell.Formula[1:]  # Remove the "=" sign
            value = target_cell.Value

            log_entry = [workbook_name, sheet_name, cell_address, formula, value, "N/A"]
            trace_log.append(log_entry)

            # Skip if the cell has already been visited
            cell_key = (workbook_name, sheet_name, cell_address)
            if cell_key in visited_cells:
                continue
            visited_cells.add(cell_key)

            # If the cell has no formula, stop tracing
            if not target_cell.HasFormula:
                continue

            # Get the precedents for internal references
            try:
                precedents = target_cell.Precedents
            except Exception:
                precedents = None

            if precedents is not None:
                # Add internal precedents to the stack
                for precedent in precedents:
                    stack.append((precedent, workbook_name, sheet_name))
            else:
                # Handle external workbook references
                try:
                    external_file_path, external_sheet_name, external_cell_address = extract_external_reference(formula)
                    external_file_path = resolve_full_path(excel_app, external_file_path)

                    # Attempt to open the external workbook if not already open
                    try:
                        external_workbook = excel_app.Workbooks(external_file_path)
                    except Exception:
                        external_workbook = None

                    if external_workbook is None:
                        for attempt in range(5):
                            try:
                                external_workbook = excel_app.Workbooks.Open(external_file_path, ReadOnly=True)
                                time.sleep(2)  # Increase wait time
                                break
                            except Exception as e:
                                print(f"Attempt {attempt + 1}: Unable to open {external_file_path} - {str(e)}")
                                time.sleep(5)

                    if external_workbook is not None:
                        external_sheet = external_workbook.Sheets(external_sheet_name)
                        next_cell = external_sheet.Range(external_cell_address)
                        stack.append((next_cell, external_workbook.Name, external_sheet.Name))
                    else:
                        log_entry[5] = f"Error: Unable to open external workbook {external_file_path}"
                except Exception as e:
                    log_entry[5] = f"Error tracing external reference: {str(e)}"

        # Write the trace log to a new Excel file
        write_trace_log(trace_log)
        print("Tracing completed. Check 'Trace_Log.xlsx' for details.")

    except Exception as e:
        print(f"Error processing the specified range: {str(e)}")

# Function to extract external reference details from a formula
def extract_external_reference(formula):
    try:
        start_pos = formula.index("'[") + 2
        end_pos = formula.index("]")
        file_path = formula[start_pos:end_pos]

        start_pos = end_pos + 2
        end_pos = formula.index("'", start_pos)
        sheet_name = formula[start_pos:end_pos]

        cell_address = formula.split("!")[-1]

        return file_path, sheet_name, cell_address
    except Exception as e:
        print(f"Error extracting reference: {str(e)}")
        return "", "", ""

# Function to resolve the full path of an external workbook
def resolve_full_path(excel_app, partial_path):
    try:
        link_sources = excel_app.ActiveWorkbook.LinkSources(1)  # xlExcelLinks = 1
        if link_sources:
            for full_path in link_sources:
                if partial_path in full_path:
                    return full_path
        return partial_path
    except Exception as e:
        print(f"Error resolving full path: {str(e)}")
        return partial_path

# Function to write the trace log to an Excel file
def write_trace_log(trace_log):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Trace Log"

    headers = ["Workbook", "Sheet", "Cell", "Formula", "Value", "File Path"]
    ws.append(headers)

    for row in trace_log:
        ws.append(row)

    wb.save("Trace_Log.xlsx")
    print("Trace log saved as 'Trace_Log.xlsx'.")

# Script execution starts here
try:
    # Launch Excel application
    excel_app = win32.Dispatch("Excel.Application")
    excel_app.Visible = True

    # Get user input for workbook path, sheet name, and range
    workbook_path = input("Enter the full path of the workbook: ")
    sheet_name = input("Enter the sheet name: ")
    cell_range = input("Enter the range of cells (e.g., A1:B10): ")

    # Trace the formulas in the specified range
    trace_formula_in_range(excel_app, workbook_path, sheet_name, cell_range)

except Exception as e:
    print(f"Error: {str(e)}")
    sys.exit(1)

finally:
    # Quit Excel application
    excel_app.Quit()

