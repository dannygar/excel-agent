# Copilot Chat Developer Commands for Declarative Agents

This document lists all available developer commands for debugging and tracing declarative agents in Microsoft 365 Copilot Chat across Excel, Word, PowerPoint, Outlook, and other M365 applications.

> **Note:** These commands are not publicly documented in one place (yet), so this is essentially your **developer cheat-sheet**.

## Table of Contents

- [Copilot Chat Developer Commands for Declarative Agents](#copilot-chat-developer-commands-for-declarative-agents)
  - [Table of Contents](#table-of-contents)
  - [Quick Start - Best Debugging Combo](#quick-start---best-debugging-combo)
  - [1. Core Developer Mode Commands](#1-core-developer-mode-commands)
    - [How to Use](#how-to-use)
  - [2. Declarative Agent Commands](#2-declarative-agent-commands)
  - [3. Tool Invocation Debugging](#3-tool-invocation-debugging)
  - [4. Conversation Debug \& System Messages](#4-conversation-debug--system-messages)
  - [5. Environment \& Versioning](#5-environment--versioning)
  - [6. Frontier Mode (Excel Agent Mode) Commands](#6-frontier-mode-excel-agent-mode-commands)
  - [7. Excel-specific Developer Commands](#7-excel-specific-developer-commands)
  - [8. Authentication \& Permissions](#8-authentication--permissions)
  - [9. Logging, Diagnostics \& State Tools](#9-logging-diagnostics--state-tools)
  - [10. Hidden / Internal Commands](#10-hidden--internal-commands)
  - [Debug Information Card](#debug-information-card)
    - [Agent Metadata Section](#agent-metadata-section)
    - [Capabilities Section](#capabilities-section)
    - [Actions Section](#actions-section)
    - [Execution Section](#execution-section)
  - [Using Developer Mode in VS Code](#using-developer-mode-in-vs-code)
    - [Quick Start](#quick-start)
    - [Debug Panel Information (VS Code)](#debug-panel-information-vs-code)
  - [Availability](#availability)
  - [Troubleshooting Tips](#troubleshooting-tips)
    - [Common Issues](#common-issues)
    - [Licensing Requirements](#licensing-requirements)
    - [Best Practices for Debugging](#best-practices-for-debugging)
  - [References](#references)

---

## Quick Start - Best Debugging Combo

⭐ **If your agent isn't firing in Excel, run these 6 commands in order:**

```text
-developer on
-frontier on
-agent info
-tools list
-tools log
-plan
```

This shows:

- ✅ Whether your agent was selected
- ✅ Whether your intent routing matched
- ✅ What arguments were passed
- ✅ What the tool returned
- ✅ Whether JSON validation failed
- ✅ Whether Excel runtime rejected the call

---

## 1. Core Developer Mode Commands

These commands can be typed directly in the Copilot chat interface when testing your declarative agent.

| Command | Description |
|---------|-------------|
| `-developer on` | **Enables developer mode** - Turns on developer logging, debug information, and exposes system messages. Shows debug information cards with detailed execution data. |
| `-developer off` | **Disables developer mode** - Returns to normal operation without debug information cards. |
| `-developer status` | **Show developer mode status** - Displays whether developer mode is currently enabled. |

### How to Use

1. Open Microsoft 365 Copilot Chat (in Excel, Word, Teams, or any M365 app)
2. Select your declarative agent from the agent picker
3. Type `-developer on` in the chat and press Enter
4. A confirmation message will appear indicating developer mode is enabled
5. Continue testing your agent - debug cards will appear with each response
6. Type `-developer off` to disable when done

![Developer Mode On](https://github.com/MicrosoftDocs/m365copilot-docs/blob/main/docs/assets/images/developer-mode-on.png)

---

## 2. Declarative Agent Commands

These commands relate to your custom declarative agents or agent add-ins (Frontier mode + add-ins).

| Command | Description |
|---------|-------------|
| `-agents list` | **List all available agents** - Shows all agents installed in your tenant. |
| `-agent info` | **Show active agent info** - Displays the agent currently invoked for the prompt (including its manifest and tool list). |
| `-agent reload` | **Reload agent manifest** - Force-reloads your declarative agent or add-in manifest. Extremely useful while iterating. |
| `-agent manifest` | **Show manifest for an agent** - Prints the merged manifest Copilot is using after applying internal validation. |
| `-agent test "<prompt>"` | **Test intent matching** - Simulates how Copilot routes your prompt to intents/functions. |

---

## 3. Tool Invocation Debugging

These are critical when your MCP/Azure Function tool isn't firing.

| Command | Description |
|---------|-------------|
| `-tools list` | **Show tools registered** - Lists all tools registered for the active agent. |
| `-tools log` | **Show tool invocation log** - Displays what tool Copilot attempted to call, the arguments sent, and the returned JSON. |
| `-tools clear` | **Clear tool invocation log** - Clears the current tool invocation log. |

---

## 4. Conversation Debug & System Messages

| Command | Description |
|---------|-------------|
| `-system` | **Show system messages** - Dumps system prompts, orchestration messages, planning steps, etc. |
| `-plan` | **Show detailed reasoning** - Shows planner output and reasoning steps. |
| `-trace on` | **Show full LLM trace** - Works only in certain rings/Frontier-enabled tenants. |
| `-trace off` | **Disable LLM trace** - Turns off the full LLM trace. |
| `-context` | **Show session variables/context** - Displays current session variables and context. |

---

## 5. Environment & Versioning

| Command | Description |
|---------|-------------|
| `-version` | **Show Copilot build version** - Displays the current Copilot build version. |
| `-config` | **Show manifest version/active config** - Shows the active configuration and manifest version. |
| `-reset` | **Clear the current session** - Resets the current conversation session. |

---

## 6. Frontier Mode (Excel Agent Mode) Commands

If you're using the new Frontier (Agent Mode for Excel Desktop):

| Command | Description |
|---------|-------------|
| `-frontier on` | **Enable Frontier mode** - Activates the Excel Agent Mode. |
| `-frontier off` | **Disable Frontier mode** - Deactivates the Excel Agent Mode. |
| `-frontier capabilities` | **List all Frontier capabilities** - Shows all available Frontier capabilities. |
| `-frontier skills` | **Show active Frontier skills** - Displays currently active Frontier skills. |

---

## 7. Excel-specific Developer Commands

These commands work **only in Excel Desktop**:

| Command | Description |
|---------|-------------|
| `-excel context` | **Dump workbook context** - Shows the current workbook context Copilot is using. |
| `-excel data` | **List data regions** - Lists data regions Copilot detects in the workbook. |
| `-excel entities` | **Show named entities** - Shows all named entities Copilot identified. |
| `-excel analyze` | **Trigger re-analysis** - Forces Copilot to re-analyze the current sheet. |
| `-excel refresh` | **Force reload data models** - Reloads data models from the workbook. |

---

## 8. Authentication & Permissions

Useful when testing your Azure Function-based MCP server or other authenticated tools:

| Command | Description |
|---------|-------------|
| `-auth status` | **Show authentication status** - Displays current authentication state. |
| `-auth reset` | **Clear auth cache** - Clears the authentication cache. |
| `-auth login` | **Force re-authenticate** - Forces a new authentication flow. |

---

## 9. Logging, Diagnostics & State Tools

| Command | Description |
|---------|-------------|
| `-log on` | **Enable detailed logs** - Turns on detailed logging. |
| `-log off` | **Disable logs** - Turns off detailed logging. |
| `-log dump` | **Dump session log** - Outputs the current session log. |

---

## 10. Hidden / Internal Commands

> ⚠️ These show up only on tenants with dev flags enabled (internal/dev rings):

| Command | Description |
|---------|-------------|
| `-help` | **Show all debug commands** - Lists all available debug commands. |
| `-raw` | **Dump raw LLM messages** - Shows raw messages sent to/from the LLM. |
| `-thoughts` | **Show planner reasoning** - Displays internal planner reasoning (only in internal/dev rings). |

---

## Debug Information Card

When developer mode is enabled, each response includes a debug information card with the following sections:

### Agent Metadata Section

| Field | Description |
|-------|-------------|
| **Agent ID** | Unique identifier for the agent (includes title ID and manifest ID) |
| **Agent version** | The version number of the agent currently in use |
| **Conversation ID** | Identifies the active chat session |
| **Request ID** | Identifies the specific prompt within the conversation |

### Capabilities Section

| Field | Description |
|-------|-------------|
| **Capabilities** | List of capabilities configured for the agent |
| **Executed capabilities** | Status and response stats for capabilities that were executed |

### Actions Section

| Field | Description |
|-------|-------------|
| **Actions** | List of actions configured for the agent |
| **Matched functions** | Status of functions matched in the runtime app index lookup |
| **Selected functions for execution** | Functions selected for invocation based on orchestrator reasoning |
| **Executed actions** | Request and response execution status for actions |

### Execution Section

| Field | Description |
|-------|-------------|
| **Execution** | List of executed capabilities and actions for the prompt |

---

## Using Developer Mode in VS Code

When using Microsoft 365 Agents Toolkit (formerly Teams Toolkit) in VS Code:

### Quick Start

1. Open your agent project in VS Code
2. Press **F5** or select **"Preview your app (F5)"** in the Agents Toolkit pane
3. This launches your agent in a browser-based Copilot Chat experience
4. Select your agent in the Copilot interface
5. Type `-developer on` to enable developer mode
6. Debug information appears in both:
   - The Copilot Chat interface (as debug cards)
   - The **Debug panel** in VS Code Agents Toolkit

### Debug Panel Information (VS Code)

The Debug panel in Agents Toolkit displays:

- Agent metadata (identifiers for agent and conversation)
- Capabilities with execution status and response statistics
- Actions configured for the agent
- Matched functions from runtime app index lookup
- Selected functions for execution based on orchestrator reasoning

---

## Availability

| Platform | Developer Mode Support |
|----------|----------------------|
| Microsoft 365 Copilot Chat (Web) | ✅ Supported |
| Microsoft Teams | ✅ Supported |
| Excel Desktop | ✅ Supported |
| Excel Online | ✅ Supported |
| Word Desktop | ✅ Supported |
| Word Online | ✅ Supported |
| PowerPoint Desktop | ✅ Supported |
| PowerPoint Online | ✅ Supported |
| Outlook Desktop | ✅ Supported |
| Outlook Web | ✅ Supported |

> **Important:** Developer mode is only available within **Microsoft 365 Copilot (Copilot for Work)** experiences. A Microsoft 365 Copilot license or tenant with Copilot Chat metering enabled is required.

---

## Troubleshooting Tips

### Common Issues

| Issue | Solution |
|-------|----------|
| Developer mode doesn't enable | Ensure you have a Microsoft 365 Copilot license or your tenant has metering enabled |
| No debug cards appearing | Make sure the agent is using enterprise knowledge, capabilities, or actions |
| Agent not visible after provisioning | Wait 1-24 hours for backend synchronization; check in Teams Admin Center |
| "Sorry, I can't chat about this" error | Content may be blocked by Responsible AI filters; review API response content |
| Processing errors with API actions | Verify OpenAPI spec has valid server URL and all required properties |

### Licensing Requirements

To use developer mode and debug declarative agents:

1. **Microsoft 365 Copilot License** - Required for the account testing the agent
2. **OR Tenant with Copilot Chat Metering** - Alternative to individual licenses

Without these, the agent backend may execute but responses won't appear in the Copilot UI.

### Best Practices for Debugging

1. **Enable developer mode first** before testing new prompts
2. **Check Agent Metadata** to verify the correct agent version is loaded
3. **Review Matched Functions** to see which functions the orchestrator considered
4. **Compare Selected vs Matched** functions to understand orchestrator reasoning
5. **Examine Execution status** for any errors in capability or action execution
6. **Use VS Code Debug Panel** for more detailed logging when using Agents Toolkit

---

## References

- [Use developer mode in Microsoft 365 Copilot to test and debug agents](https://learn.microsoft.com/microsoft-365-copilot/extensibility/debugging-agents-copilot-studio)
- [Use developer mode to test and debug agents in Microsoft 365 Agents Toolkit](https://learn.microsoft.com/microsoft-365-copilot/extensibility/debugging-agents-vscode)
- [Set up your development environment for Microsoft 365 Copilot](https://learn.microsoft.com/microsoft-365-copilot/extensibility/prerequisites)
- [Copilot extensibility in the Microsoft 365 ecosystem](https://learn.microsoft.com/microsoft-365-copilot/extensibility/ecosystem)

---

Last updated: December 2024
