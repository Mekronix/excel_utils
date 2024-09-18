local function loadLuacom(pathOfCallerDir)
    package.loadlib(pathOfCallerDir .. "/../lib/luacom.dll", "luacom_openlib")()
    -- or use require('luacom') 
end

-- helper functions

ExcelVisibility = false
ShouldCreateOutput = true

local function trim(s)
    return tostring(s):match("^%s*(.-)%s*$")
end

local function tableContains(table, element)
    for _, value in pairs(table) do
        if value == element then
            return true
        end
    end
    return false
end

local function extractOptionalValues(value)
    if value == nil then
        return '', '', '', ''
    end

    local allowedWords = { 'path', 'worksheet'}

    value = trim(value)

    local keyValuePairs = {}

    for pair in string.gmatch(value, "([^,]+)") do
        local key, val = string.match(pair, "%s*(%w+)%s*=%s*(.+)%s*")
        
        key = key and key:match("^%s*(.-)%s*$")
        val = val and val:match("^%s*(.-)%s*$")
    
        if key and tableContains(allowedWords, key) and val then
            keyValuePairs[key] = val
        end
    end
    return keyValuePairs['path'] or "", keyValuePairs['worksheet'] or ""
end

local function setExcelVisible()
    ExcelVisibility = true
end

local function setNoOutput()
    ShouldCreateOutput = false
end

local function logExcelError(message, functionName)
    error('\n\nError in excel_utils (made by Mekronix) caused by ' .. functionName .. '\n -- ' .. message .. ' -- \n')
end

-- only returns a number if the input string contains a number or is empty
-- else the same string gets returned
local function parseNumber(inputString)
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

local function extractBracketsContent(inputString)
    local content = string.match(inputString, "!(.-)!")
    
    if content then
        local modifiedString = string.gsub(inputString, "!.-!", "", 1)
        return content, modifiedString
    else
        return "", inputString
    end
end


local function extractFileNameFromPath(path)
    return path:match("[^\\/]+$")
end

local function getWorkSheetFromWorkbook(workbook, sheetName, functionName)
    if type(sheetName) == "number" then
        if sheetName < 1 or sheetName > workbook.Worksheets.Count then
            logExcelError('The given workbook index is out of range', functionName)
        end
        return workbook.Worksheets:Item(sheetName)
    end

    local success, result = pcall(function()
        return workbook.Worksheets(sheetName)
    end)

    if success then
        return result
    else
        logExcelError('The given workbook does not exist', functionName)
    end
end

-- if the worksheet is already opened it is going to use it, as it is a lot faster
-- but it is not able to distinguish two .xlsx files with the same name.
local function getWorksheet(path, sheet, functionName)
    local sheetName = parseNumber(sheet)
    local excel = luacom.GetObject("Excel.Application")
    local fileName = extractFileNameFromPath(path)

    if excel ~= nil and excel.Workbooks ~= nil then
        for i = 1, excel.Workbooks.Count do
            local workbook = excel.Workbooks(i)
            if extractFileNameFromPath(workbook.FullName) == fileName then
                print('An instance of "' .. fileName .. '" is already initiated and will be used.')
                return excel, getWorkSheetFromWorkbook(workbook, sheetName, functionName), false
            end
        end
    end

    excel = luacom.CreateObject("Excel.Application")
    excel.Visible = ExcelVisibility

    local workbook = excel.Workbooks:Open(path)

    return excel, getWorkSheetFromWorkbook(workbook, sheetName, functionName), true
end

local function getPath(pathIndex, functionName)
    local index = parseNumber(pathIndex)
    if type(index) == 'number' then
        if index < 1 or index > #pathList then
            logExcelError('The given workbook index is out of range', functionName)
        end
        return pathList[index]
    end

    return index
end

local function getExcelIndex(input, functionName)
    if input == nil then
        logExcelError('A column index is missing or wrong', functionName)
    end

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
    tex.sprint(pathList[i])
end

local function getAllPaths()
    for i = 1, #pathList do
        tex.sprint(pathList[i] .. '\\\\')
    end
end

local function getCellValue(column, row, option)
    if not ShouldCreateOutput then return end

    local pathIndex, sheet = extractOptionalValues(option)

    column = getExcelIndex(column, 'getCellValue')

    local path = getPath(pathIndex, 'getCellValue')
    local excel, worksheet, shouldClose = getWorksheet(path, sheet, 'getCellValue')
    local cellValue = worksheet.Cells(row, column).Value2

    if cellValue == nil then
        cellValue = ""
    else
        cellValue = trim(cellValue)
    end

    tex.sprint(cellValue)
    print('Cell Value: ' .. cellValue)

    if shouldClose then
        excel:Quit()
        excel = nil
    end
end

local function getCellValues(startCol, startRow, endCol, endRow, option, tableOrPlot)
    if not ShouldCreateOutput then return end
    
    local plotOption
    plotOption, option = extractBracketsContent(option)
    local pathIndex, sheet = extractOptionalValues(option)
    
    startCol = getExcelIndex(startCol, 'getCellValues')
    endCol = getExcelIndex(endCol, 'getCellValues')

    local path = getPath(pathIndex, 'getCellValues')
    local excel, worksheet, shouldClose = getWorksheet(path, sheet, 'getCellValues')

    local separator
    local rowEnd

    if tableOrPlot == 1 then
        separator = ' & '
        rowEnd = '\\\\ \\hline \n'
    elseif tableOrPlot == 2 then
        separator = ','
        rowEnd = ' '
        tex.sprint('\\addplot ' .. '[' .. plotOption .. ']' .. 'coordinates {')
        print('\\addplot ' .. '[' .. plotOption .. ']' .. 'coordinates {')
    end

    local allRows = {}
    local i = 0
    for row = startRow, endRow do
        local rowValues = {}
        for col = startCol, endCol do
            local cellValue = worksheet.Cells(row, col).Value2
            if cellValue == nil then
                cellValue = ""
            else
                cellValue = trim(cellValue)
            end
            if tableOrPlot == 2 and startCol == endCol then
                if cellValue == "" then
                    goto continue
                end
                i = i + 1
                table.insert(rowValues, tostring(i) .. ',' .. tostring(cellValue))
            else
                table.insert(rowValues, tostring(cellValue))
            end
        end
        if #rowValues > 0 then
            if tableOrPlot == 1 then
                table.insert(allRows, table.concat(rowValues, separator))
            elseif tableOrPlot == 2 then
                table.insert(allRows, '(' .. table.concat(rowValues, separator) .. ')')
            end
        end
        ::continue::
    end

    local finalOutput = table.concat(allRows, rowEnd)

    print(finalOutput)
    tex.sprint(finalOutput)

    if tableOrPlot == 2 then
        tex.sprint('};')
        print('};')
    end

    if shouldClose then
        excel:Quit()
        excel = nil
    end
end

local function getCellValuesTwice(firstStartCol, firstStartRow, firstEndCol, firstEndRow, secondStartCol, secondStartRow, secondEndCol, secondEndRow, option, tableOrPlot)
    if not ShouldCreateOutput then return end
    
    local plotOption
    plotOption, option = extractBracketsContent(option)
    local pathIndex, sheet = extractOptionalValues(option)

    firstStartCol = getExcelIndex(firstStartCol, 'getCellValuesTwice')
    firstEndCol = getExcelIndex(firstEndCol, 'getCellValuesTwice')
    secondStartCol = getExcelIndex(secondStartCol, 'getCellValuesTwice')
    secondEndCol = getExcelIndex(secondEndCol, 'getCellValuesTwice')

    local path = getPath(pathIndex, 'getCellValuesTwice')
    local excel, worksheet, shouldClose = getWorksheet(path, sheet, 'getCellValuesTwice')

    local separator
    local rowEnd

    if tableOrPlot == 1 then
        separator = ' & '
        rowEnd = '\\\\ \\hline \n'
    elseif tableOrPlot == 2 then
        separator = ','
        rowEnd = ' '
        tex.sprint('\\addplot ' .. '[' .. plotOption .. ']' .. 'coordinates {')
        print('\\addplot ' .. '[' .. plotOption .. ']' .. 'coordinates {')
    end
    
    -- Ensure both areas have the same number of rows
    if (firstEndRow - firstStartRow) ~= (secondEndRow - secondStartRow) then
        logExcelError('The two areas do not have the same amount of rows', 'getCellValuesTwice')
    end

    local allRows = {}
    for row = firstStartRow, firstEndRow do
        local rowValues = {}
        for col = firstStartCol, firstEndCol do
            local cellValue = worksheet.Cells(row, col).Value2
            if cellValue == nil then
                if tableOrPlot == 2 then
                    goto continue
                end
                cellValue = ""
            else
                cellValue = trim(cellValue)
            end
            table.insert(rowValues, tostring(cellValue))
        end
        for col = secondStartCol, secondEndCol do
            local cellValue = worksheet.Cells(row, col).Value2
            if cellValue == nil then
                if tableOrPlot == 2 then
                    goto continue
                end
                cellValue = ""
            else
                cellValue = trim(cellValue)
            end
            table.insert(rowValues, tostring(cellValue))
        end
        if #rowValues > 0 then
            if tableOrPlot == 1 then
                table.insert(allRows, table.concat(rowValues, separator))
            elseif tableOrPlot == 2 then
                table.insert(allRows, '(' .. table.concat(rowValues, separator) .. ')')
            end
        end
        ::continue::
    end

    local finalOutput = table.concat(allRows, rowEnd) -- here make change

    tex.sprint(finalOutput)
    print(finalOutput)

    if tableOrPlot == 2 then
        tex.sprint('};')
        print('};')
    end

    if shouldClose then
        excel:Quit()
        excel = nil
    end
end

return {
    loadLuacom = loadLuacom,                -- loads luacom with certain path
    setNoOutput = setNoOutput,              -- excel_utils wont create output
    setExcelVisible = setExcelVisible,      -- make the opening of the excel files visible
    addPath = addPath,                      -- adds path to list
    getPathAt = getPathAt,                  -- returns path at index
    getCellValue = getCellValue,            -- return values at certain cell value
    getAllPaths = getAllPaths,              -- returns all paths
    getCellValues = getCellValues,          -- return all values in area
    getCellValuesTwice = getCellValuesTwice -- returns all values in two areas
}
