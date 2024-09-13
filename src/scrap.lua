require('luacom')

-- helper functions

local function throwError(message, functionName)
    error('\n\nError in scrap.lua (made by Alexander) caused by ' .. functionName .. '\n -- ' .. message .. ' -- \n')
end

local function parseInt(inputString)
    if inputString == "" then
        return 1
    end

    local number = tonumber(inputString)

    if number then
        return math.floor(number)
    else
        return inputString
    end
end

local function getFileName(path)
    return path:match("[^\\/]+$")
end

local function getWorkSheetFromWorkbook(workbook, sheetName)
    if type(sheetName) == "number" then
        return workbook.Worksheets:Item(1)
    else
        return workbook.Worksheets(sheetName)
    end
end

local function getWorksheet(path, sheet)
    local sheetName = parseInt(sheet)
    local excel = luacom.GetObject("Excel.Application")
    local fileName = getFileName(path)

    if excel ~= nil and excel.Workbooks ~= nil then
        for i = 1, excel.Workbooks.Count do
            local workbook = excel.Workbooks(i)
            if getFileName(workbook.FullName) == fileName then
                print('An instance of "' .. fileName .. '" is already initiated and will be used.')
                return excel, getWorkSheetFromWorkbook(workbook, sheetName), false
            end
        end
    end

    excel = luacom.CreateObject("Excel.Application")
    excel.Visible = true

    local workbook = excel.Workbooks:Open(path)

    return excel, getWorkSheetFromWorkbook(workbook, sheetName), true
end

local function getPath(pathIndex)
    local path = pathList[pathIndex] or pathList[1]

    if not path then
        throwError('No path provided and pathList is empty.', 'getCellValue')
    end

    return path
end

local function getRangeValues(startRow, startCol, endRow, endCol, worksheet)
    local values = {}
    for row = startRow, endRow do
        local rowValues = {}
        for col = startCol, endCol do
            local cellValue = worksheet.Cells(row, col).Value2
            if cellValue == nil then
                cellValue = ""
            end
            table.insert(rowValues, tostring(cellValue))
        end
        table.insert(values, rowValues)
    end
    return values
end

local function getExcelIndex(input)
    local num = tonumber(input)

    if num ~= nil then
        return num
    else
        local index = 0
        for i = 1, #input do
            local char = input:sub(i, i)
            local value = string.byte(char) - string.byte("A") + 1
            index = index * 26 + value
        end
        return index
    end
end

-- methods which can be accessed from file

pathList = {}

local function addPath(path)
    table.insert(pathList, path)
end

local function getPathAt(i)
    return tex.sprint(pathList[i])
end

local function getAllPaths()
    for i = 1, #pathList do
        tex.sprint(pathList[i])
    end
end

-- This function retrieves and prints the value of a specific cell in an Excel worksheet.
--
-- The function takes the following inputs:
-- - `pathIndex`: Index of the path in `pathList` used to open the Excel file. If not provided, the first path in `pathList` is used.
-- - `sheet`: The name or index of the worksheet from which to retrieve the cell value. If not provided, the first worksheet is used.
-- - `row`: The row number of the cell to retrieve.
-- - `column`: The column number of the cell to retrieve.
--
-- The function checks if a valid path is provided; otherwise, it throws an error. It also parses
-- the worksheet name or index and retrieves the cell value from the specified row and column.
-- Finally, it prints the cell value and ensures that the Excel application and workbook are closed properly.
local function getCellValue(pathIndex, sheet, row, column)
    column = getExcelIndex(column)

    local path = getPath(pathIndex)
    local excel, worksheet, shouldClose = getWorksheet(path, sheet)
    local cellValue = worksheet.Cells(row, column).Value2

    tex.sprint(cellValue)

    if shouldClose then
        excel:Quit()
        excel = nil
    end
end

-- This function extracts values from a specified range of cells in an Excel worksheet
-- and formats them with customizable separators and end-of-line characters.
--
-- The function takes the following inputs:
-- - `pathIndex`: Index of the path in `pathList` used to open the Excel file.
-- - `startRow`, `startCol`: Coordinates of the starting cell in the range.
-- - `endRow`, `endCol`: Coordinates of the ending cell in the range.
-- - `separator`: Character(s) to use between values within the same row (e.g., `", "`).
-- - `endOfLine`: Character(s) to append at the end of each row (e.g., `"\\ \hline"`).
--
-- If the `separator` is `"%"`, it is replaced with `" % "` for formatting. The function
-- retrieves all values within the specified cell range and prints them in a formatted manner,
-- with the chosen separator and end-of-line character.
local function getCellValues(pathIndex, sheet, startRow, startCol, endRow, endCol, separator, endOfLine)
    startCol = getExcelIndex(startCol)
    endCol = getExcelIndex(endCol)
    local path = getPath(pathIndex)
    local excel, worksheet, shouldClose = getWorksheet(path, sheet)

    for row = startRow, endRow do
        local rowValues = {}
        for col = startCol, endCol do
            local cellValue = worksheet.Cells(row, col).Value2
            if cellValue == nil then
                cellValue = ""
            end
            table.insert(rowValues, tostring(cellValue))
        end
        tex.sprint(table.concat(rowValues, separator) .. ' ' .. endOfLine)
        print(table.concat(rowValues, separator) .. ' ' .. endOfLine)
    end

    if shouldClose then
        excel:Quit()
        excel = nil
    end
end

local function getCellValuesTwice(pathIndex, sheet, firstStartRow, firstStartCol, firstEndRow, firstEndCol,
                                  secondStartRow, secondStartCol, secondEndRow, secondEndCol, separator, endOfLine)
    firstStartCol = getExcelIndex(firstStartCol)
    firstEndCol = getExcelIndex(firstEndCol)
    secondStartCol = getExcelIndex(secondStartCol)
    secondEndCol = getExcelIndex(secondEndCol)

    local path = getPath(pathIndex)
    local excel, worksheet, shouldClose = getWorksheet(path, sheet)

    -- Ensure both areas have the same number of rows
    if (firstEndRow - firstStartRow) ~= (secondEndRow - secondStartRow) then
        throwError('The two areas do not have the same amount of rows', 'getCellValuesTwice')
    end

    local firstRangeValues = getRangeValues(firstStartRow, firstStartCol, firstEndRow, firstEndCol, worksheet)
    local secondRangeValues = getRangeValues(secondStartRow, secondStartCol, secondEndRow, secondEndCol, worksheet)

    for row = 1, #firstRangeValues do
        local combinedRowValues = {}

        for _, value in ipairs(firstRangeValues[row]) do
            table.insert(combinedRowValues, value)
        end

        for _, value in ipairs(secondRangeValues[row]) do
            table.insert(combinedRowValues, value)
        end

        tex.sprint(table.concat(combinedRowValues, separator) .. endOfLine)
        print(table.concat(combinedRowValues, separator) .. endOfLine)
    end

    if shouldClose then
        excel:Quit()
        excel = nil
    end
end

-- all latex outputs will be printed out in the consol. Therefore if there should be some errors 
-- according to the way the values, take a look there. 
-- to insert columns as strings, declare them so e.g. "'AB'" and not "AB"
-- paths can be inserted with / or with a \, keep in mind that lua as many other programming
-- languages give \ a special function in strings (e.g. \n = new line). Therefore to add one single
-- \ two \\ need to be inserted.
return {
    addPath = addPath,                      -- adds path to list
    getPathAt = getPathAt,                  -- returns path at index
    getCellValue = getCellValue,            -- return values at certain cell value
    getAllPaths = getAllPaths,              -- returns all paths
    getCellValues = getCellValues,          -- return all values in area
    getCellValuesTwice = getCellValuesTwice -- returns all values in two areas
}
