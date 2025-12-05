# Debugging Excel Agent in Desktop

This guide explains how to debug your Excel Agent, including event-based handlers and agent functions, when running in Excel Desktop.

## Prerequisites

### 1. Install Required Extension

You need one of these VS Code extensions:

- **Microsoft Office Add-in Debugger** (recommended) - `msoffice.microsoft-office-add-in-debugger`
- **Debugger for Microsoft Edge** - `msjsdiag.debugger-for-edge`

The recommended extension should be suggested automatically when you open this workspace. Click "Install" when prompted, or install manually:

1. Open VS Code Extensions (Ctrl+Shift+X)
2. Search for "Microsoft Office Add-in Debugger"
3. Click Install

### 2. Run VS Code as Administrator

**Important:** On Windows, you must run VS Code as Administrator to attach to WebView2 processes.

- Right-click VS Code icon → "Run as Administrator"
- Or run from command line with admin privileges

### 3. Verify Debug Configuration

Your workspace already has the necessary debug configurations in `.vscode/launch.json`:

- `Excel Desktop (Edge Chromium)` - For debugging taskpane and commands
- `Attach to Office Add-in (WebView2)` - For debugging event-based handlers

## Debugging Workflow

### Option 1: Debug Agent with Automatic Excel Launch

This is the easiest method for general debugging:

1. **Set breakpoints** in your code:
   - `src/commands/commands.ts` - Your agent function handlers
   - Any other TypeScript files you want to debug

2. **Start debugging:**
   - Press `F5` or go to Run and Debug (Ctrl+Shift+D)
   - Select "Debug Copilot in Excel Desktop" from the dropdown
   - Click the green play button

3. **What happens:**
   - VS Code builds your add-in
   - Excel Desktop launches automatically
   - Your add-in loads
   - Debugger attaches

4. **Trigger your agent:**
   - In Excel, open Copilot
   - Type a prompt like "Create a table starting at A1 with headers Name, Age, City and rows John, 25, NYC | Jane, 30, LA"
   - Your breakpoints should hit when the function executes

### Option 2: Debug with Process Picker (Recommended for Connection Issues)

If you encounter "Cannot connect to localhost:0" errors, use this method:

1. **Start your add-in:**
   - Run task: "Start Agent Locally" (from Terminal → Run Task)
   - Or manually: `yarn dev-server` then `yarn start:desktop:excel`

2. **Load the add-in in Excel:**
   - Excel should open automatically
   - Make sure your add-in is loaded and visible
   - You can open the taskpane or just verify it's in the ribbon

3. **Attach using Process Picker:**
   - Go to Run and Debug (Ctrl+Shift+D)
   - Select: **"Attach to Office Add-in (Process Picker)"**
   - Press F5
   - A dropdown will appear showing available Edge processes
   - Select the one that shows `localhost:3000` in its title

4. **Set breakpoints and test:**
   - Set breakpoints in `src/commands/commands.ts`
   - In Excel Copilot, trigger your agent function
   - Breakpoints should now hit

### Option 3: Debug Event-Based Handlers (Manual Attach)

When you see the "Debug Event-based handler" dialog in Excel, use this method:

1. **Start your add-in:**
   - Run task: "Start Agent Locally" (from Terminal → Run Task)
   - Or press F5 to use "Debug Copilot in Excel Desktop"

2. **Trigger the agent in Excel:**
   - Open Copilot in Excel
   - Type a prompt that calls your agent function
   - You'll see a dialog: **"Debug Event-based handler"**
   - **DON'T CLICK OK YET!**

3. **Attach VS Code debugger:**
   - Go to Run and Debug (Ctrl+Shift+D)
   - Select: **"Attach to Office Add-in (Process Picker)"** (recommended)
   - Or: **"Attach to Office Add-in (WebView2)"**
   - Press F5 to attach
   - If using Process Picker, select the localhost:3000 process
   - Wait for "Debugger attached" message in VS Code

4. **Resume execution:**
   - Go back to Excel
   - Click **OK** in the debug dialog
   - Your breakpoints will now hit

5. **Debug:**
   - Step through code with F10/F11
   - Inspect variables
   - View console output
   - Check network calls to MCP server

### Option 3: Debug Without Auto-Launch

If you want more control over when Excel starts:

1. **Start the dev server manually:**
   ```powershell
   yarn dev-server
   ```

2. **Launch Excel separately:**
   ```powershell
   yarn start:desktop:excel
   ```

3. **Attach debugger:**
   - In VS Code, go to Run and Debug
   - Select "Attach to Office Add-in (WebView2)"
   - Press F5

## Troubleshooting

### Breakpoints Not Hitting

**Check source maps:**
- Ensure your build includes source maps (check `webpack.config.js`)
- Verify the `sourceMapPathOverrides` in `launch.json` match your build output

**Verify file paths:**
- The manifest's runtime section must reference the correct JS file
- Ensure you're setting breakpoints in the actual source file, not a copy
- For event-based handlers, breakpoints must be in the same file specified in the manifest

**Rebuild the project:**
```powershell
yarn build:dev
```

### "Cannot connect to runtime process" or "Cannot connect to the target at localhost:0" Error

This error occurs when the debugger can't find or connect to the WebView2 process.

**Solutions:**

1. **Use the Process Picker method instead:**
   - Select "Attach to Office Add-in (Process Picker)" from debug dropdown
   - VS Code will show a list of available Edge processes
   - Select the one that shows your localhost:3000 URL

2. **Ensure VS Code is running as Administrator:**
   - On Windows, WebView2 debugging requires admin privileges
   - Right-click VS Code → "Run as Administrator"

3. **Make sure Excel is actually running with your add-in loaded:**
   - Verify the dev server is running (`yarn dev-server`)
   - Check that Excel has loaded the add-in (look for your add-in in the ribbon)
   - Open the taskpane or trigger a function before attaching

4. **Close all Excel instances and restart:**
   - Close all Excel windows completely
   - End any Excel.exe processes in Task Manager
   - Restart the debug session

5. **Verify WebView2 is being used:**
   - Check that Excel is using Edge WebView2 (not IE or legacy WebView)
   - Newer Office versions use WebView2 by default
   - Update Office if necessary

6. **Try attaching after triggering the function:**
   - Start Excel with your add-in
   - Type your Copilot prompt (but don't press Enter yet)
   - Start the debugger attachment
   - Once attached, press Enter in Copilot

7. **Enable trace logging:**
   - The WebView2 config has `trace: true` enabled
   - Check the Debug Console for detailed connection attempts
   - Look for errors about pipe names or ports

### Debug Dialog Doesn't Appear

The "Debug Event-based handler" dialog only appears for event-based handlers (like launch events, AppSource handlers, etc.). 

For standard agent function calls:
- Use Option 1 (automatic debugging)
- Breakpoints should hit without the dialog

### Agent Functions Not Being Called

**Verify function registration:**
1. Check `manifest.json` - action IDs must match
2. Check `Office-API-local-plugin.json` - function names must match
3. Check `commands.ts` - `Office.actions.associate()` names must match

**Example:**
```typescript
// All three must match exactly:
// manifest.json: "id": "CreateTable"
// Office-API-local-plugin.json: "name": "CreateTable"
// commands.ts: Office.actions.associate("CreateTable", ...)
```

### Remote MCP Server Issues

If the Microsoft Docs MCP Server isn't responding:

1. **Check the URL in `Office-API-local-plugin.json`:**
   ```json
   {
     "type": "RemoteMCPServer",
     "spec": {
       "url": "https://mcp.microsoft.com/microsoft-docs"
     }
   }
   ```

2. **Test the endpoint** separately (if possible)

3. **Check network logs** in the debugger Console

### Console Logging

Add console logs to help debug:

```typescript
Office.actions.associate("CreateTable", async (message) => {
  console.log("CreateTable called with:", message);
  try {
    const params = JSON.parse(message);
    console.log("Parsed params:", params);
    // ... rest of code
  } catch (error) {
    console.error("Error in CreateTable:", error);
    return `Error: ${error.message}`;
  }
});
```

View logs in VS Code Debug Console when debugger is attached.

## Debug Configuration Details

### Excel Desktop (Edge Chromium)

```jsonc
{
  "name": "Excel Desktop (Edge Chromium)",
  "type": "msedge",
  "request": "attach",
  "url": "https://localhost:3000/*",
  "preLaunchTask": "Debug: Excel Desktop",
  "postDebugTask": "Stop Debug"
}
```

- Auto-launches Excel
- Attaches to taskpane/commands
- Runs build and cleanup tasks

### Attach to Office Add-in (WebView2)

```jsonc
{
  "name": "Attach to Office Add-in (WebView2)",
  "type": "msedge",
  "request": "attach",
  "useWebView": true
}
```

- Manual attach to running Excel
- Required for event-based handler debugging
- Use when you see "Debug Event-based handler" dialog

## Additional Resources

- [Microsoft Learn: Debug Office Add-ins](https://learn.microsoft.com/office/dev/add-ins/testing/debug-add-ins-overview)
- [Debug event-based handlers](https://learn.microsoft.com/office/dev/add-ins/testing/debug-function-command)
- [Attach debugger to WebView2](https://learn.microsoft.com/microsoft-edge/webview2/how-to/debug-devtools)

## Quick Reference

| Task | Command |
|------|---------|
| Start debugging (with Excel launch) | F5 → "Debug Copilot in Excel Desktop" |
| Attach to running Excel | F5 → "Attach to Office Add-in (WebView2)" |
| Start dev server only | `yarn dev-server` |
| Launch Excel only | `yarn start:desktop:excel` |
| Build for development | `yarn build:dev` |
| Stop debug | Stop button in VS Code or `yarn stop` |

## Debugging Remote Deployment

For debugging the remotely deployed agent:

1. Select "Debug Copilot in Excel Desktop (Remote)"
2. Uses dev tunnel URL instead of localhost
3. Same debugging workflow as local

Remember: Always run VS Code as Administrator for Excel Desktop debugging on Windows.
