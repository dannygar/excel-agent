async function fillColor(cell: string, color: string) {
  // @ts-ignore
  await Excel.run(async (context) => {
    context.workbook.worksheets.getActiveWorksheet().getRange(cell).format.fill.color = color;
    await context.sync();
  })
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
});