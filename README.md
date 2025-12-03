# Excel Agent - Office JavaScript Library API Plugin

This project demonstrates how to build an **API plugin for Microsoft 365 Copilot** that leverages the **Office JavaScript Library** to perform read and write operations on Excel documents. By combining a declarative agent with API plugins that call Office JavaScript APIs, you can enable Copilot to interact with Office documents in precise, error-free ways that go beyond natural language generation.

## Overview

This declarative agent extends Microsoft 365 Copilot to work directly with Excel cells through a natural language interface. The agent uses API plugins to call Office JavaScript Library APIs, enabling scenarios such as:

- **Content Analysis**: Analyze spreadsheet content and take action based on what's found
- **Trusted Data Insertion**: Insert data unchanged from trusted sources into Excel
- **Natural Language Interface**: Provide a conversational way to interact with Excel functionality
- **Parameterized Actions**: Execute complex operations with parameters passed at runtime through natural language

### Key Features

- Execute Office JavaScript Library APIs through Copilot chat
- Change cell colors and formatting through natural language commands
- Handle user prompts and convert them into precise Excel operations
- Preview feature showcasing the future of Office extensibility

## Get Started

### Prerequisites

Before you begin, ensure you have the following installed and configured:

- **[Node.js](https://nodejs.org/)** - Supported versions: 18, 20, or 22
- **[Visual Studio Code](https://code.visualstudio.com/)**
- **[Microsoft 365 Agents Toolkit](https://aka.ms/M365AgentsToolkit)** - VS Code extension version 5.0.0 or higher
- **[Microsoft 365 account for development](https://docs.microsoft.com/microsoftteams/platform/toolkit/accounts)** with:
  - [Microsoft 365 Copilot license](https://learn.microsoft.com/microsoft-365-copilot/extensibility/prerequisites#prerequisites)
  - Permission to upload custom apps (sideloading enabled)

> **Note**: During the preview period, this feature is only available for Excel, PowerPoint, and Word on Windows and on the web. Mac support is coming soon.

### Local Development Setup

Follow these steps to run the Excel Agent locally:

#### 1. Install Dependencies

Open a terminal in the project root and run:

```bash
yarn install
```

#### 2. Sign In to Microsoft 365

1. Open the project in Visual Studio Code
2. Select the **Microsoft 365 Agents Toolkit** icon in the VS Code activity bar
3. In the **ACCOUNTS** section, sign in with your Microsoft 365 account that has Copilot enabled

#### 3. Provision the Local Environment

Use the Microsoft 365 Agents Toolkit to provision resources:

- In the **LIFECYCLE** pane, click **Provision**

Alternatively, run from the terminal:

```bash
yarn provision
```

This creates the necessary build artifacts in the `appPackage/build` folder.

#### 4. Start the Development Server

Start the local web server to serve the plugin files:

```bash
yarn dev-server
```

> **Important**: If prompted to delete an old certificate or install a new one, agree to both prompts for proper HTTPS setup.

Wait until you see a message indicating successful compilation (e.g., "webpack compiled successfully").

#### 5. Test in Excel

##### Option A: Using VS Code Tasks

You can use the pre-configured VS Code task:

- Press `F5` or select **Run > Start Debugging**
- Choose **"Start Agent Locally"** from the launch configuration

##### Option B: Manual Testing

1. **Open Excel**:
   - **Windows**: Open Excel desktop application
   - **Web**: Navigate to [https://excel.cloud.microsoft](https://excel.cloud.microsoft)

2. Open or create a workbook

3. **Open Copilot**:
   - Click the **Copilot** button on the ribbon, or
   - If you see a Copilot dropdown menu, select **App Skills**

4. **Select the Agent**:
   - Click the hamburger menu (☰) in the Copilot pane
   - Find **Excel Agent** in the list (you may need to select "See more")
   - If the agent isn't listed, wait a few minutes and reload Copilot (Ctrl+R)

5. **Try it out**:
   - Select the **"Change cell color"** conversation starter, or
   - Type a command like: "Change the color of cell B2 to orange"
   - When prompted, click **Confirm** to allow the action
   - The cell color should change!

### Remote Development Setup (Dev/Production)

For deploying to a remote environment with dev tunnels:

#### 1. Set Up Dev Tunnel

If you haven't already, log in and create a dev tunnel:

```bash
yarn tunnel:login
yarn tunnel:create
yarn tunnel:port
```

See [DEVTUNNEL-SETUP.md](./DEVTUNNEL-SETUP.md) for detailed tunnel configuration instructions.

#### 2. Provision for Dev Environment

```bash
yarn provision:dev
```

This provisions resources for the `dev` environment (uses `env/.env.dev` configuration).

#### 3. Start the Agent Remotely

Use the VS Code task **"Start Agent Remotely"** or manually:

1. Start the dev tunnel:

   ```bash
   yarn tunnel:host
   ```

2. In a separate terminal, start the dev server:

   ```bash
   yarn dev-server
   ```

3. The agent will be accessible through the dev tunnel URL

#### 4. Build for Production

To create a production build:

```bash
yarn build
```

This creates optimized bundles in the `dist` folder.

### Making Changes During Development

Live reloading is not supported during the preview period. To test changes:

1. **Stop the server**:
   - Press `Ctrl+C` in the terminal running `dev-server`, or run:

     ```bash
     yarn stop
     ```

2. **Clear Office cache** following the [official instructions](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/clear-cache#manually-clear-the-cache)

3. **Uninstall the agent**:
   - Open Microsoft Teams
   - Navigate to **Apps > Manage your apps**
   - Find **Excel Agentdev** in the list
   - Click the trash icon and select **Remove**

4. **Make your code changes**

5. **Restart the development process** (Provision and start dev-server again)

## Project Structure

### Folders

| Folder       | Contents                                                                                 |
| ------------ | ---------------------------------------------------------------------------------------- |
| `.vscode`    | VS Code configuration files for debugging and tasks                                      |
| `appPackage` | App package templates including manifests, plugin definitions, and assets               |
| `env`        | Environment-specific configuration files                                                 |
| `src`        | Source code including TypeScript functions that call Office JavaScript APIs              |

### Key Configuration Files

The following files define the agent's behavior and can be customized:

| File                                      | Purpose                                                                                           |
| ----------------------------------------- | ------------------------------------------------------------------------------------------------- |
| `appPackage/manifest.json`                | Microsoft 365 app manifest defining metadata, permissions, and extensions configuration           |
| `appPackage/declarativeAgent.json`        | Declarative agent configuration including instructions, conversation starters, and action references |
| `appPackage/Office-API-local-plugin.json` | API plugin manifest specifying functions, parameters, and the connection to Office JavaScript Library |
| `src/commands/commands.ts`                | TypeScript implementation of functions that call Office JavaScript APIs                           |
| `src/commands/commands.html`              | HTML runtime loader for the Office JavaScript Library                                             |
| `m365agents.yml`                          | Microsoft 365 Agents Toolkit project configuration for provisioning and deployment                |

### Important Manifest Properties

**manifest.json**:

- `authorization.permissions.resourceSpecific`: Grants permission to read/write Office documents
- `extensions.runtimes`: Configures the JavaScript runtime and maps action IDs to functions

**declarativeAgent.json**:

- `instructions`: Defines how the agent should behave and interpret user requests
- `conversation_starters`: Provides example prompts to help users get started
- `actions`: References the API plugin configuration file

**Office-API-local-plugin.json**:

- `functions`: Defines each action with parameters, descriptions, and instructions
- `runtimes.spec.local_endpoint`: Set to `"Microsoft.Office.Addin"` to indicate Office JavaScript Library usage

## Extending This Agent

This project demonstrates a basic implementation of changing cell colors. You can extend it by:

### Adding More Office JavaScript API Functions

1. Add new actions to `Office-API-local-plugin.json` with their function definitions
2. Implement the corresponding TypeScript functions in `src/commands/commands.ts`
3. Map the action IDs in both `manifest.json` and using `Office.actions.associate()`
4. Update the declarative agent instructions in `appPackage/declarativeAgent.json`

### Example Extensions

- **Data manipulation**: Insert, update, or format data in cells and ranges
- **Chart creation**: Generate charts based on spreadsheet data
- **Formula insertion**: Add formulas to cells programmatically
- **Worksheet management**: Create, rename, or delete worksheets
- **Table operations**: Create and manipulate Excel tables

### Enhancing the Declarative Agent

- **[Add conversation starters](https://learn.microsoft.com/microsoft-365-copilot/extensibility/build-declarative-agents?tabs=ttk&tutorial-step=3)**: Provide more example prompts to guide users
- **[Add knowledge sources](https://learn.microsoft.com/microsoft-365-copilot/extensibility/build-declarative-agents?tabs=ttk&tutorial-step=5)**: Ground the agent with OneDrive, SharePoint, or web content
- **Refine instructions**: Improve how the agent interprets and responds to user requests
- **[Add more API plugins](https://learn.microsoft.com/microsoft-365-copilot/extensibility/build-declarative-agents?tabs=ttk&tutorial-step=7)**: Integrate with REST APIs alongside Office JavaScript APIs

## Additional Resources

### Documentation

- **[Build API plugins for Microsoft 365 Copilot with Office JavaScript Library](https://learn.microsoft.com/en-us/microsoft-365-copilot/extensibility/build-api-plugins-local-office-api)** - Official tutorial this project is based on
- **[Office JavaScript API Reference](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/understanding-the-javascript-api-for-office)** - Complete API documentation
- **[Declarative Agents for Microsoft 365](https://aka.ms/teams-toolkit-declarative-agent)** - Declarative agent overview
- **[Microsoft 365 App Manifest Schema](https://learn.microsoft.com/en-us/microsoft-365/extensibility/schema/)** - Manifest reference documentation
- **[API Plugin Manifest Schema](https://learn.microsoft.com/en-us/microsoft-365-copilot/extensibility/api-plugin-manifest-2.4)** - Plugin configuration reference
- **[Declarative Agent Schema](https://learn.microsoft.com/en-us/microsoft-365-copilot/extensibility/declarative-agent-manifest-1.6)** - Agent configuration reference

### Tools

- **[Microsoft 365 Agents Toolkit](https://aka.ms/M365AgentsToolkit)** - VS Code extension for building agents
- **[Microsoft 365 Agents Toolkit CLI](https://aka.ms/teamsfx-toolkit-cli)** - Command-line interface

### Support & Feedback

This feature is currently in **preview**. Your feedback is valuable:

- **Limitations**: Currently supports Excel, PowerPoint, and Word on Windows and web (Mac support coming soon)
- **Future improvements**: Templates and tooling are being developed to streamline this type of project

## Troubleshooting

### Agent doesn't appear in Copilot

- Wait a few minutes after provisioning
- Press `Ctrl+R` with the Copilot pane focused
- Verify you're signed in with the correct Microsoft 365 account
- Check that your account has Copilot enabled and sideloading permissions

### Certificate Errors

- Accept prompts to install development certificates when running `yarn dev-server`
- On Windows, you may need to restart VS Code after certificate installation

### Changes Not Reflected

- Remember to stop the server, clear the Office cache, uninstall the agent, and re-provision after making changes
- Live reload is not supported during the preview period

### Plugin Function Errors

- Check browser console (F12) when testing in Office on the web
- Verify action IDs match across `manifest.json`, `Office-API-local-plugin.json`, and `commands.ts`
- Ensure parameter names and types match between the plugin manifest and TypeScript implementation

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
