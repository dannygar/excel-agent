# Practical Debugging Guide for Excel Agent

The WebView2 attach debugger in VS Code is notoriously unreliable for Office Add-ins. Here's what actually works.

## Method 1: Using F12 Developer Tools (Recommended & Most Reliable)

This is the most straightforward way to debug Office Add-ins in Excel Desktop.

### Steps:

1. **Start your dev server:**
   ```powershell
   yarn dev-server
   ```

2. **Launch Excel with your add-in:**
   ```powershell
   yarn start:desktop:excel
   ```
   
   Or press F5 → "Debug Copilot in Excel Desktop"

3. **Open Developer Tools:**
   
   **For Runtime Add-ins (like this agent), use Edge DevTools:**
   
   Since your add-in runs as a runtime without a visible taskpane:
   
   1. Make sure Excel is running with your add-in
   2. Open **Microsoft Edge** browser
   3. Navigate to: `edge://inspect/#devices`
   4. Under "Other" or "Remote Target", find entries with `localhost:3000`
   5. Click the blue **"inspect"** link
   6. DevTools opens connected to your add-in
   
   **Alternative - Enable Registry DevTools (One-time setup):**
   
   To make DevTools automatically open:
   
   1. Open **Registry Editor** (Win+R, type `regedit`)
   2. Navigate to: `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer`
   3. Create DWORD (32-bit) value: `EnableWebViewDevTools`
   4. Set value to `1`
   5. Restart Excel
   6. DevTools will auto-open when add-in runs

4. **Navigate to your code:**
   
   In the Developer Tools that open:
   
   - Go to the **Sources** tab
   - Expand the file tree on the left
   - Find `webpack://` → your TypeScript files
   - Open `src/commands/commands.ts`

5. **Set breakpoints:**
   
   - Click on the line numbers to set breakpoints
   - You can also use `debugger;` statements in your code

6. **Trigger your agent:**
   
   - In Excel, open Copilot
   - Type your prompt: "Create a table starting at A1 with headers Name, Age, City and rows John, 25, NYC | Jane, 30, LA"
   - Your breakpoints will hit

7. **Debug:**
   
   - Step through code (F10, F11)
   - Inspect variables in the right panel
   - Use the Console tab to run JavaScript
   - View network requests in the Network tab

### Advantages:

✅ Always works - no connection issues  
✅ Full DevTools functionality  
✅ Network panel for API/MCP debugging  
✅ Console for real-time testing  
✅ Performance profiling  
✅ No VS Code admin requirements

### Disadvantages:

❌ Not integrated with VS Code  
❌ Separate window  
❌ Must re-set breakpoints if you reload

## Method 2: Console Logging (Simplest)

Add strategic console logs to your code:

```typescript
Office.actions.associate("CreateTable", async (message) => {
  console.log("=== CreateTable called ===");
  console.log("Raw message:", message);
  
  try {
    const params = JSON.parse(message);
    console.log("Parsed params:", params);
    console.log("StartCell:", params.StartCell);
    console.log("Headers:", params.Headers);
    console.log("Rows:", params.Rows);
    
    await createTable(params.StartCell || "A1", params.Headers || "", params.Rows || "");
    
    console.log("✅ Table created successfully");
    return `Table created successfully`;
  } catch (error) {
    console.error("❌ Error in CreateTable:", error);
    console.error("Stack:", error.stack);
    return `Error: ${error.message}`;
  }
});
```

View logs:
- Press F12 in Excel
- Go to Console tab
- See all your logs in real-time

## Method 3: Debug with Edge DevTools (Recommended for Runtime Add-ins)

This is the most reliable method for agent/runtime-only add-ins:

1. **Start your add-in in Excel:**
   
   ```powershell
   yarn dev-server
   yarn start:desktop:excel
   ```

2. **Open Edge DevTools:**
   
   - Open **Microsoft Edge** browser (separate window)
   - Navigate to: `edge://inspect/#devices`
   - You'll see a list of "Inspectable pages"
   
3. **Find your add-in:**
   
   - Look under "Other" or "Remote Target" sections
   - Find entries showing `localhost:3000` or your add-in's URL
   - There may be multiple entries - look for one with `commands.html` or similar
   
4. **Click "inspect":**
   
   - Click the blue **"inspect"** link next to your add-in
   - DevTools opens in a new window, connected to your add-in
   
5. **Debug as normal:**
   
   - Go to Sources tab → find your TypeScript files under `webpack://`
   - Set breakpoints by clicking line numbers
   - Go to Console to see console.log output
   - Trigger your agent in Excel - breakpoints will hit

### Advantages of Edge DevTools:

✅ Works for runtime-only add-ins (no taskpane needed)  
✅ More reliable than in-app debugging  
✅ Full Chrome DevTools features  
✅ Can keep open across Excel restarts  
✅ Better for long debugging sessions

## Method 4: Using Alerts (Last Resort)

For quick debugging without DevTools:

```typescript
Office.actions.associate("CreateTable", async (message) => {
  try {
    const params = JSON.parse(message);
    
    // Show what was received
    alert(`Received:\nCell: ${params.StartCell}\nHeaders: ${params.Headers}\nRows: ${params.Rows}`);
    
    await createTable(params.StartCell || "A1", params.Headers || "", params.Rows || "");
    
    return "Table created successfully";
  } catch (error) {
    alert(`Error: ${error.message}`);
    return `Error: ${error.message}`;
  }
});
```

## Testing Without Copilot

To test your functions directly without going through Copilot:

1. **Add a test button in commands.html:**

   ```html
   <button onclick="testCreateTable()">Test CreateTable</button>
   <script>
     function testCreateTable() {
       const message = JSON.stringify({
         StartCell: "A1",
         Headers: "Name,Age,City",
         Rows: "John,25,NYC|Jane,30,LA"
       });
       
       Office.actions.getByIdAsync("CreateTable", (result) => {
         if (result.status === Office.AsyncResultStatus.Succeeded) {
           result.value.invoke(message);
         }
       });
     }
   </script>
   ```

2. **Or call directly from DevTools Console:**

   ```javascript
   // In F12 Console
   const testMessage = JSON.stringify({
     StartCell: "A1",
     Headers: "Name,Age,City",
     Rows: "John,25,NYC|Jane,30,LA"
   });
   
   Office.actions.getByIdAsync("CreateTable", (result) => {
     result.value.invoke(testMessage).then(console.log);
   });
   ```

## Debugging Remote MCP Server Issues

To debug the Microsoft Docs MCP Server function:

1. **Check if the function is being called:**
   
   ```typescript
   Office.actions.associate("SearchMicrosoftDocs", async (message) => {
     console.log("SearchMicrosoftDocs called with:", message);
     const params = JSON.parse(message);
     console.log("Query:", params.query);
     // ... rest of code
   });
   ```

2. **Use the Network tab in F12 DevTools:**
   
   - Open F12 → Network tab
   - Trigger the search function
   - Look for requests to `mcp.microsoft.com`
   - Check response status and data

3. **Test the MCP server directly:**
   
   You can test if the MCP server is reachable from the Console:
   
   ```javascript
   fetch('https://mcp.microsoft.com/microsoft-docs/health')
     .then(r => r.json())
     .then(console.log)
     .catch(console.error);
   ```

## Common Issues & Solutions

### Issue: "Office.actions is undefined"

**Solution:** Office.js hasn't loaded yet. Ensure your code is in the `Office.onReady()` callback:

```typescript
Office.onReady((info) => {
  // All your Office.actions.associate calls here
});
```

### Issue: Function not being called from Copilot

**Check:**

1. Function name matches in all three places:
   - `manifest.json` → actions → id
   - `Office-API-local-plugin.json` → functions → name
   - `commands.ts` → Office.actions.associate

2. Plugin is correctly referenced in `declarativeAgent.json`:
   ```json
   "actions": [
     {
       "id": "ExcelActions",
       "file": "Office-API-local-plugin.json"
     }
   ]
   ```

3. Add-in is properly loaded (check Excel ribbon/add-ins menu)

### Issue: Parameters are wrong or undefined

**Debug the message:**

```typescript
Office.actions.associate("CreateTable", async (message) => {
  console.log("Raw message type:", typeof message);
  console.log("Raw message:", message);
  console.log("Message length:", message.length);
  
  // Try to parse
  try {
    const params = JSON.parse(message);
    console.log("Parsed successfully:", params);
  } catch (e) {
    console.error("Parse failed:", e);
    console.log("Message content:", message);
  }
});
```

## VS Code Debugging (When It Works)

The "Excel Desktop (Edge Chromium)" configuration may work sometimes:

1. Close all Excel instances
2. Run VS Code as Administrator
3. Press F5
4. Set breakpoints
5. Trigger function

If it fails with connection errors, use F12 DevTools instead.

## Summary: Best Approach for Runtime Add-ins

For reliable debugging of agent/runtime add-ins like yours:

### **Recommended Method (Do This Now):**

1. **Use Edge DevTools (`edge://inspect`)** - Most reliable for runtime add-ins
   - Start Excel with your add-in running
   - Open Edge browser → `edge://inspect/#devices`
   - Find your `localhost:3000` entry under "Other"
   - Click "inspect"
   - Full debugging capabilities available

2. **View Console Logs** - I've already added comprehensive logging
   - In Edge DevTools → Console tab
   - See all the logs when you trigger CreateTable
   - Shows raw messages, parsed params, success/error

3. **Set Breakpoints** - For step-by-step debugging
   - In Edge DevTools → Sources tab
   - Navigate to `webpack://` → `src/commands/commands.ts`
   - Click line numbers to set breakpoints

### **Alternative Methods:**

- **Registry DevTools** - For automatic popup (one-time setup)
- **Console.log everywhere** - For production monitoring
- **VS Code debugging** - Not recommended for runtime add-ins (unreliable)

### **Quick Start Right Now:**

```powershell
# Terminal 1: Start dev server
yarn dev-server

# Terminal 2: Launch Excel
yarn start:desktop:excel
```

Then:
1. Open Edge browser
2. Go to `edge://inspect/#devices`
3. Find your add-in and click "inspect"
4. Trigger CreateTable in Excel Copilot
5. Watch the Console logs in DevTools

This method works perfectly for runtime-only add-ins where there's no visible taskpane to right-click on.
