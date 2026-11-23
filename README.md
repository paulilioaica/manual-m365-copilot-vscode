# Manual M365 Copilot for VS Code

Build the VSIX with the command below. Install it with **Extensions: Install from VSIX…** in desktop VS Code 1.104 or later. Run **Manual M365 Copilot: Paste Token** in the Command Palette. In VS Code Chat, open the model picker and choose **M365 Copilot (Manual Token)**. If needed, open **Chat: Manage Language Models** and look under **Manual M365 Copilot**. The provider's management action also opens token entry.

Paste a Microsoft 365 business BizChat token. The extension accepts a raw token, a token prefixed with `Bearer`, or a Microsoft BizChat WebSocket URL. It holds the credential in memory until VS Code exits, reloads the window, or you run **Manual M365 Copilot: Forget Token**. Use **Choose License Route** to match an included or paid Microsoft 365 Copilot license. The native picker also offers Quick Response and Think Deeper, subject to account availability.

For the Copilot harness, VS Code may require **Chat › Agent Host › Byok Models: Enabled** (`chat.agentHost.byokModels.enabled`) and may require GitHub sign-in. An organization can control whether third-party providers are available. See [Microsoft's BYOK documentation](https://github.com/microsoft/vscode-docs/blob/main/docs/agent-customization/language-models.md#bring-your-own-language-model-key).

This extension uses the local VS Code runtime and the source files packaged in the VSIX. Replies come from Microsoft 365 over the network. The extension has no third-party package dependencies, webview, chat participant, token renewal, background calls, or automatic retries.

In the VS Code 1.147 Copilot harness, the extension emits an opaque continuation marker after a completed model response. When VS Code returns the marker, the extension keeps the same temporary Microsoft conversation and sends only new text or the tool result. System messages are skipped. The tool catalog is repeated only if the offered tools or required-tool mode change. A new chat, model, license route, or account cannot reuse another chat's context.

If the host omits the marker, the extension uses a fresh temporary conversation with up to 5,000 characters of recent complete turns. Stale markers, including ones from before an editor restart or from a turn that already advanced, require a new VS Code chat. Continuations stay in memory and expire after 45 minutes idle or when the local store is full. They never contain the credential and are not sent to Microsoft. Microsoft can expire its own conversation independently.

VS Code sees a 128,000-token input budget for planning. The outbound cap remains 12,000 characters for each Microsoft request, including any tools or replay. A prompt that does not fit is rejected locally. VS Code may include editor context inside its user messages. The extension does not read workspace files on its own.

The extension sends one BizChat request per model invocation. VS Code owns tool approvals and execution, and a tool result can trigger a new model invocation. **Manual M365 Copilot: Show Request Diagnostics** shows the latest context mode, new message and replayed context characters, total chat invocations, and any allowance counters returned by Microsoft. A continued turn with no replay shows `contextMode: continued` and `replayedContextChars: 0`. Smaller payloads alone do not establish a reduction in Microsoft's daily allowance debits.

Live availability depends on your account and Microsoft's unofficial BizChat endpoint. The JavaScript source and AGPL license are included in the VSIX under `extension/`.

## Develop and package

The extension requires no package installation. Run the local mock tests with Node.js 22 or later:

```sh
node --test test/*.test.js
```

Build the VSIX using the Python 3 standard library:

```sh
python3 scripts/package.py
```

The archive is written to `dist/`. The tests use fake credentials and a local mock Microsoft server.
