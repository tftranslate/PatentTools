# Patent Tools for Microsoft Word v. 0.2.0-beta

See [CHANGELOG.md](CHANGELOG.md) for release history.

Patent Tools is a Microsoft Word `.dotm` add-in for connecting Microsoft Word to a locally or remotely deployed OpenAI compatible OpenAI-compatible language model endpoint.

At the present time, the tool supports compiling reference sign lists and smart insertion of the reference signs into claim text. Unlike existing solutions, it supports any language supported by the model and will happily deal with inflected languages such as German, French and others, as well as with difficult cases where reference signs depend on context. It is packaged as a self-contained Word template add-in with a custom Ribbon tab.

Flexible programmatic checks make sure that any model hallucinations, omissions or additions other than reference signs are not carried over into the claim set. All changes applied by the plugin, i.e. the reference signs inserted, are marked up in track changes mode so you can be sure the model does not modify the claims in any unintended way.

This add-in has been validated to work reasonably well and fast in English as well as non-English languages with sufficiently smart edge models hosted by llama.cpp. The target hardware are unified memory systems such as DGX Spark, AMD Strix Halo or a Mac Ultra. However this add-in can of course also be connected to a public or private cloud-based service that provides OpenAI compatible API endpoints.

Further features may be added in the future.

![Screenshot of Patent Tools](docs/screenshot.png)

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
- **Edit reference signs** — lets you populate the reference sign list from the description and manually edit it before insertion.
- **Settings** — opens the model and prompt configuration dialog.

## Settings dialog

The settings dialog lets the user configure the model connection and request behavior. These settings are stored per user and persist after Word is closed and reopened through VBA settings storage (Windows registry).

### Model Settings ###

- **OpenAI compatible URL** — base URL including host and port, for example `http://127.0.0.1:11434` if you are hosting with Ollama, or`https://api.openai.com/v1` for a ChatGPT API subscription.

- **API Key** — optional; can usually be left empty if you use Ollama or llama.cpp for self-hosting.

- **Fetch model list** — verifies the connection to the model and automatically discovers which models the API offers.

- **Model name** — offers a list of models offer by the API, choose the one you want to use.

- **Temperature** — floating-point value using a dot, for reference sign insertion use a low one such as `0.2`. Accepts digits plus one dot.

- **Timeout** — integer value in seconds. In non-thinking mode 120 should be sufficient, raise if you plan to use thinking mode to up to 600 (10 minutes). Accepts only integer values.

- **Thinking** — checkbox that controls whether `chat_template_kwargs` enables thinking mode. In my experience this is not required for sufficiently smart models. Thinking mode will considerably slow down the process. Also, for gpt-oss, which does not honour ```chat_template_kwargs``, reasoning_effort``` will be set to ```low``` when thinking is off and to default (```medium```) when thinking is on.

  It is noted that temperature, timeout and thinking mode can be configured individually for reference sign insertion (use a very low temperature here) and reference sign list population (a higher temperature helps the model discover cases of ambiguity of reference sign usage here)

- **Max. Tokens** — integer token limit. Upper limit is your model's context window size.

The URL is normalized before use. If the user enters a URL ending in `/v1` or `/v1/chat/completions`, that suffix is removed so the add-in can append the correct endpoint path itself. If llama.cpp is discovered, a native ```completions`` path is used that offers more feedback while the model is working.

### Prompt settings

The two prompt settings pages in the settings dialog allow to edit large parts of the system prompt that the code sends to the model in each case. These are considered expert settings, do not fiddle with them unless you know what you are doing.

The insertion prompt supports the ```{PARAGRAPH_COUNT}``` placeholder, which will be filled with the number of claim paragraphs that are sent to the model. It is vital that the model have this information, otherwise the later matching with the original claims will likely fail.

If you want to play with prompts, start with the population prompt that generates the reference sign list. This one is less brittle than the insertion prompt and may be tuned to your liking as to the level of detail of the so-called "further observations" that will be added below the actual list. 

## Privacy and security

Patent claims can be highly sensitive before publication. Although technically possible, users should **not** point Patent Tools at the public OpenAI API service if the claims are not yet published or otherwise cleared for external disclosure.

For unpublished or confidential patent material, use a **local or self-controlled OpenAI-compatible endpoint**, for example a private `llama.cpp` server running on the user’s own machine or local network or on a hosted server for which a privacy agreement is in place.

Recommended safe usage:

- Use a local model endpoint for confidential drafts.
- Keep the API key blank when using a local server that does not require authentication.
- Only use external hosted APIs when the document content may legally and contractually be sent there.

## Validated model setup

This add-in has been validated with:

- **Model:** gemma-4 31b, qwen-3.8 27b, and gpt-oss-120b

  Tests with gemma-4 26b-a4b were promising and it might work for you, but it is not as stable as gemma-4 31b.
  
  Note: For gemma and qwen, turn ```thinking```off for claim insertion, but on for population. For gpt-oss, turn ```thinking```off for both.
- **Mode:** non-thinking mode / low reasoning effort
- **Endpoint type:** OpenAI-compatible chat completions API

## Installation

To make the Ribbon and macros available in all Word windows, the `.dotm` must be installed as a **global Word add-in**, not merely opened like a normal template. Word automatically loads `.dotm` templates placed in the Word **Startup** folder at launch.

### Manual installation

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
5. Select the relevant claim text in your currently open document, or leave nothing selected to process the full document (not recommended)
6. Click **Insert reference signs** in the Patent Tools ribbon.
7. Paste the previously prepared reference sign table into the dialog. **Please note**: If you use a smart model, you may also add any special prompts in difficult cases where the model should pay attention on how to do things the way you want.
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
- `chat_template_kwargs.enable_thinking` and ```reasoning_effort low``` depending on the Thinking checkbox
- standard `messages` content

If an API key is provided, the macro sends:

```
Authorization: Bearer <API key>
```

If the API key field is blank, no Authorization header is sent.

## Recommended local setup

A practical confidential setup is:

- Microsoft Word with `PatentTools.dotm` in the Startup folder.
- A local OpenAI-compatible server, such as `llama.cpp` or Ollama.
- A model validated for this workflow, such as gemma-4 26b in non-thinking mode.

Example base URL (this is for ollama on your own computer):

```
http://127.0.0.1:11434
```

The add-in will normalize the base URL and call the chat-completions endpoint at `/v1` automatically.

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

This error message occurs when the model outputs extra words or otherwise reorders the claim wording. Consider lowering temperature, using a better model, or add even more explicit instructions to the prompt. The Google Gemma-4 31 b dense model is a good reference point. If you keep seeing this issue with a reasonable precise model, contact the developer.

### During processing I see "Model is thinking", but I disabled "Thinking" in the settings

This should not happen with the validated Gemma and Qwen models. The reason is probably that you are using a thinking-only model or a model that dues not support disabling thinking using ```chat_template_kwargs```. Note that when you disable thinking in the settings, PatentTools will also pass ```reasoning_effort: low```. Some models, such as gpt-oss, will be cause by this to "think less" when thinking is disabled in settings, and to "think more" when thinking is enabled. Compare the time required for thinking between both scenarios to see whether ticking and unticking "thinking" has this effect on your model.

### Repository contents

Current repository structure:

```
.github/
  FUNDING.yml
.gitignore
AI_USAGE.md
BUILD.md
Build-PatentTools.cmd
CHANGELOG.md
LICENSE
README.md
docs/
  screenshot.png
scripts/
  Build-PatentTools.ps1
source/
  package/
    [Content_Types].xml
    _rels/
      .rels
  ribbon/
    customUI/
      _rels/
        customUI14.xml.rels
      customUI14.xml
      images/
        refsignicon.png
  vba/
    frmPatentToolsSettings.frm
    frmPatentToolsSettings.frx
    frmPromptPreview.frm
    frmPromptPreview.frx
    frmRefList.frm
    frmRefList.frx
    modHttpStream.bas
    modJsonHelper.bas
    modPatentToolsconfig.bas
    modRefSigns.bas
    modRibbon.bas
template/
  PatentTools.base.dotm
```

| Directory / File | Purpose |
|---|---|
| `source/vba/` | VBA module, form, and frx files (macros) |
| `source/ribbon/` | Custom UI XML defining the Ribbon tab and icons |
| `source/package/` | OPC package structure for rebuilding `.dotm` from source |
| `template/` | The base template (`PatentTools.base.dotm`) used to build `.dotm` |
| `scripts/` | PowerShell helper scripts for building the add-in |
| `docs/` | Screenshots and images |
| `.github/FUNDING.yml` | Donation configuration for "Buy me a coffee" link |
| `AI_USAGE.md`, `BUILD.md` | Project documentation for AI usage and build process |

## Status

This project currently provides a working Word add-in with:

- a custom Ribbon tab,
- a settings dialog,
- persistent user configuration,
- and a claim-processing workflow validated with gemma-4 31b, qwen-3.8 27b and gpt-oss-120b

## Todo
This was a project for my own needs and is considered finished for as long as I am happy with it. However, I am also  to develop it further and help others if I see there is demand. Buy me a coffee to incrase my motiviation!

## Support
You can support the developer by buying him a coffee. One-time and regular donations are very welcome: [Buy me a coffee](buymeacoffee.com/tftranslate).

You can contact me at t.ernst@tf-translate.net.

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

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
