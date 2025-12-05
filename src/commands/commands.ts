async function fillColor(cell: string, color: string) {
  // @ts-ignore
  await Excel.run(async (context) => {
    context.workbook.worksheets.getActiveWorksheet().getRange(cell).format.fill.color = color;
    await context.sync();
  })
}

async function enterValue(cell: string, value: string) {
  // @ts-ignore
  await Excel.run(async (context) => {
    context.workbook.worksheets.getActiveWorksheet().getRange(cell).values = [[value]];
    await context.sync();
  })
}

async function createTable(startCell: string, headers: string, rows: string) {
  // @ts-ignore
  await Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getActiveWorksheet();
    
    // Parse headers and rows, trimming whitespace
    const headerArray = headers.split(',').map(h => h.trim()).filter(h => h.length > 0);
    const rowsArray = rows.split('|')
      .filter(row => row.trim().length > 0)
      .map(row => row.split(',').map(cell => cell.trim()));
    
    // Validate that all rows have the same number of columns as headers
    const numCols = headerArray.length;
    const validRows = rowsArray.filter(row => row.length === numCols);
    
    // Combine headers and rows
    const tableData = [headerArray, ...validRows];
    const numRows = tableData.length;
    
    // Calculate the range address
    // Parse start cell to get column and row
    const cellMatch = startCell.match(/^([A-Z]+)(\d+)$/i);
    if (!cellMatch) {
      throw new Error(`Invalid cell address: ${startCell}`);
    }
    
    const startCol = cellMatch[1].toUpperCase();
    const startRow = parseInt(cellMatch[2]);
    
    // Calculate end column
    const startColCode = startCol.charCodeAt(0);
    const endColCode = startColCode + numCols - 1;
    const endCol = String.fromCharCode(endColCode);
    const endRow = startRow + numRows - 1;
    
    const rangeAddress = `${startCol}${startRow}:${endCol}${endRow}`;
    
    // Set the data
    const range = worksheet.getRange(rangeAddress);
    range.values = tableData;
    
    // Format the header row
    const headerRangeAddress = `${startCol}${startRow}:${endCol}${startRow}`;
    const headerRange = worksheet.getRange(headerRangeAddress);
    headerRange.format.fill.color = "#4472C4";
    headerRange.format.font.color = "white";
    headerRange.format.font.bold = true;
    
    await context.sync();
  });
}

// @ts-ignore
Office.onReady((info) => {
  // @ts-ignore
  // The first parameter of the associate method must exactly match both the extensions.runtimes.actions.id property 
  // in the manifest and the functions.name property in the API plugins JSON.
  // The message parameter is an object passed by the Copilot runtime to the JavaScript runtime in Office. 
  // It's an object that contains the cell address and color parameters as the user specified in a natural language prompt, 
  // such as "Set cell C4 to green."
  Office.actions.associate("FillColor", async (message) => {
    const { Cell: cell, Color: color } = JSON.parse(message);
    await fillColor(cell, color);
    return "Cell color changed.";
  })

  // @ts-ignore
  Office.actions.associate("EnterValue", async (message) => {
    const { Cell: cell, Value: value } = JSON.parse(message);
    await enterValue(cell, value);
    return "Cell value updated.";
  })

  // @ts-ignore
  Office.actions.associate("CreateTable", async (message) => {
    console.log("=== CreateTable called ===");
    console.log("Raw message:", message);
    
    try {
      const params = JSON.parse(message);
      console.log("Parsed params:", params);
      
      const startCell = params.StartCell || "A1";
      const headers = params.Headers || "";
      const rows = params.Rows || "";
      
      console.log("StartCell:", startCell);
      console.log("Headers:", headers);
      console.log("Rows:", rows);
      
      if (!headers || !rows) {
        console.error("Missing headers or rows");
        return "Error: Both headers and rows are required to create a table.";
      }
      
      await createTable(startCell, headers, rows);
      
      const numCols = headers.split(',').length;
      const numRows = rows.split('|').length;
      console.log(`✅ Table created: ${numCols} columns, ${numRows} rows`);
      
      return `Table created successfully at ${startCell} with ${numCols} columns and ${numRows} rows.`;
    } catch (error) {
      console.error("❌ Error in CreateTable:", error);
      console.error("Stack:", error.stack);
      return `Error creating table: ${error.message}`;
    }
  })
});