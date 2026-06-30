# Patent Tools for Microsoft Word v. 0.1.0

Patent Tools is a Microsoft Word `.dotm` add-in for connecting Microsoft Word to a locally deployed OpenAI compatible OpenAI-compatible language model endpoint.

At the present time, the tool supports inserting patent reference signs into claim text. Unlike existing solutions, it supports any language supported by the model and will happily deal with inflected languages such as German, French and others. It is packaged as a self-contained Word template add-in with a custom Ribbon tab.

This add-in has been validated to work reasonably well and fast in English as well as non-English languages when using **gemma-4 26b** in **non-thinking mode** via an OpenAI-compatible server endpoint provided by llama.cpp running on a DGX Spark.

All changes applied by the plugin, i.e. the reference signs inserted, are marked up in track changes mode so you can be sure the model does not modify the claims in any unintended way.

Further features may be added in the future.

## What it does

Patent Tools is designed to:

- Read the current selection, or the whole document if nothing is selected.
- Let you enter a list of reference signs (extracting reference signs is currently out of scope, you could just upload the description into a chatbot and ask it to prepare a list) in free format along with any custom prompts.
- Use any OpenAI-compatible chat-completions endpoint to produce a version of the claims from the current selection or whole documents with reference signs inserted.
- Insert the returned reference signs into the Word document using Track Changes without altering anything else.
- Keep formatting and punctuation intact as far as possible through conservative alignment logic.

Typical use case:

- Patent claim drafting in Word.
- Semi-automated insertion of reference signs in response to a first office action.
- Local/private model workflows using `llama.cpp` or another OpenAI-compatible inference server.

## Ribbon interface

After installation, Word shows a custom Ribbon tab named **Patent Tools**. The tab contains:

- **Insert reference signs** — runs the main claim-processing workflow.
- **Settings** — opens the model configuration dialog.

## Settings dialog

The settings dialog lets the user configure the model connection and request behavior. The current version includes these fields:

- **OpenAI compatible URL** — base URL including host and port, for example `http://127.0.0.1:8080`.
- **API Key** — optional; can be left empty.
- **Model name** — entered manually, use exactly as advertised by your llama.cpp under /v1/models. If in doubt to a "curl $URL/v1/models and look for the advertised "id". Auto-discovery is planned for a follow up release.
- **Temperature** — floating-point value using a dot, for claim insertion use a low one such as `0.2`. Accepts digits plus one dot.
- **Timeout** — integer value in seconds. In non-thinking mode 120 should be sufficient, raise if you plan to use thinking mode to up to 600 (10 minutes). Accepts only integer values.
- **Max. Tokens** — integer token limit.
- **Thinking** — checkbox that controls whether `chat_template_kwargs` enables thinking mode. In my experience this is not required for sufficiently smart models. Thinking mode will considerably slow down the process.

The URL is normalized before use. If the user enters a URL ending in `/v1` or `/v1/chat/completions`, that suffix is removed so the add-in can append the correct endpoint path itself.

These settings are stored per user and persist after Word is closed and reopened through VBA settings storage.

## Privacy and security

Patent claims can be highly sensitive before publication. Users should **not** point Patent Tools at the public OpenAI API service if the claims are not yet published or otherwise cleared for external disclosure.

For unpublished or confidential patent material, use a **local or self-controlled OpenAI-compatible endpoint**, for example a private `llama.cpp` server running on the user’s own machine or local network.

Recommended safe usage:

- Use a local model endpoint for confidential drafts.
- Keep the API key blank when using a local server that does not require authentication.
- Only use external hosted APIs when the document content may legally and contractually be sent there.

## Validated model setup

This add-in has been validated with:

- **Model:** gemma-4 26b
- **Mode:** non-thinking mode
- **Endpoint type:** OpenAI-compatible chat completions API

Thinking mode is available as a checkbox in the settings dialog, but the validated baseline configuration is **non-thinking mode**.

## Installation

To make the Ribbon and macros available in all Word windows, the `.dotm` must be installed as a **global Word add-in**, not merely opened like a normal template. Word automatically loads `.dotm` templates placed in the Word **Startup** folder at launch.

## Manual installation

1. Close all Word windows.
2. Locate Word’s Startup folder.
3. Copy `PatentTools.dotm` into that folder.
4. Reopen Word.
5. Confirm that a **Patent Tools** tab appears in the Ribbon when you open a Word document.

A common Startup path on Windows is:

```
C:\Users\<YourUserName>\AppData\Roaming\Microsoft\Word\STARTUP
```

The exact folder can be checked in Word under:

```
File -> Options -> Advanced -> General -> File Locations -> Startup
```

If the file was downloaded from the internet, Windows may block it. In that case:

1. Right-click `PatentTools.dotm`.
2. Choose **Properties**.
3. If shown, click **Unblock**.
4. Click **Apply** and **OK**.

## How to use

1. Open Word.
2. Open the **Patent Tools** tab.
3. Click **Settings** and configure the endpoint.
4. Prepare a reference-sign table in any format that would be understood by our model.
5. Select the relevant claim text, or leave nothing selected to process the full document.
6. Click **Insert reference signs**.
7. Paste the previously prepared reference sign table into the dialog. You may also add any special prompts in difficult cases where the model should pay attention on how to do things the way you want.
8. The status bar (windows button shows elapsed time to indicate work in progress).
9. Review the inserted changes in Track Changes mode.

## Endpoint requirements

Patent Tools expects an **OpenAI-compatible** chat-completions API. The macro currently posts to:

```
<base-url>/v1/chat/completions
```

The request includes:

- `model`
- `temperature`
- `max_tokens`
- `response_format` with `json_object`
- `chat_template_kwargs.enable_thinking` depending on the Thinking checkbox
- standard `messages` content

If an API key is provided, the macro sends:

```
Authorization: Bearer <API key>
```

If the API key field is blank, no Authorization header is sent.

## Current behavior and limitations

- The model name must currently be entered manually.
- The add-in is optimized for preserving original claim wording and only inserting reference signs.
- It relies on conservative word-by-word alignment, so major rewrites by the model can cause a paragraph to be skipped safely rather than applied incorrectly. However, the model is not supposed to re-write the claims, so this is expected behaviour.
- Best results are obtained when the model follows the instruction to reproduce the paragraph text exactly and only add parenthesized reference signs.

## Recommended local setup

A practical confidential setup is:

- Microsoft Word with `PatentTools.dotm` in the Startup folder.
- A local OpenAI-compatible server, such as `llama.cpp`.
- A model validated for this workflow, such as gemma-4 26b in non-thinking mode.

Example base URL:

```
http://127.0.0.1:8080
```

The add-in will normalize the base URL and call the chat-completions endpoint automatically.

## Troubleshooting

### The Patent Tools tab only appears when the template is opened directly

The template is being opened as a document/template, not loaded as a global add-in. Copy it to Word’s Startup folder and restart Word.

### The settings are lost after restarting Word

The settings should persist per user through VBA settings storage. If they do not, confirm that the settings dialog was closed with **OK** rather than **Cancel**, because only valid saved values are persisted.

### The model connection fails

Check:

- the base URL,
- whether the local server is running,
- whether the model name is correct,
- whether the timeout is long enough,
- whether an API key is required by the endpoint.

### Error message "Done, but these paragraphs were skipped because the word sequence could not be aligned safely"

This error message occur when the model outputs extra words or otherwise reorders the claim wording. Consider lowering temperature, using a better model, or add even more explicit instructions to the prompt. The Goolge Gemma-4 models are a good reference point. If you keep seeing this issue with a reasonble precise model, contact the developer.

### Repository contents

Typical repository structure:

```
textPatentTools/
├─ PatentTools.dotm
├─ README.md
└─ docs/
   └─ settings-dialog.png
```

## Status

This project currently provides a working Word add-in with:

- a custom Ribbon tab,
- a settings dialog,
- persistent user configuration,
- and a claim-processing workflow validated with gemma-4 26b in non-thinking mode.

## Todo
This was a project for my own needs and is considered finished for as long as I am happy with it. However, I am also  to develop it further and help others if I see there is demand. Buy me a coffee to incrase my motiviation! The following things would probably be worth improving:

- autodectection of available models rahter than having to type in the model name by hand,
- other API types, in particular Ollama,
- a feature to automatically extract the list of reference signs from the description in case it is available (at present, you need do this by hand with your chatbot).

## Support
You can support the developer by buying him a coffee. One-time and regular donations are very welcome: [Buy me a coffee](buymeacoffee.com/tftranslate).

## License
This project is licensed under the [MIT License](LICENSE).

---

MIT License

Copyright (c) 2026 Tobias Friedrich Ernst

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
