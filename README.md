# Google Sheets AI Functions

🤖 **Simple AI integration for Google Sheets**

Use multiple AI providers (Gemini, OpenAI, Claude, DeepSeek) directly in Google Sheets cells with custom functions.

## ⚡ Quick Setup

### 1. Create Apps Script Project

1. Go to [script.google.com](https://script.google.com/)
2. Create new project
3. Replace the default code with the content from `Code.gs`

### 2. Configure API Keys

Go to **Project Settings** (⚙️) → **Script Properties** and add your API keys:

| Property           | Description           | Required          |
| ------------------ | --------------------- | ----------------- |
| `GEMINI_API_KEY`   | Your Gemini API key   | If using Gemini   |
| `OPENAI_API_KEY`   | Your OpenAI API key   | If using OpenAI   |
| `CLAUDE_API_KEY`   | Your Claude API key   | If using Claude   |
| `DEEPSEEK_API_KEY` | Your DeepSeek API key | If using DeepSeek |

### 3. Get API Keys

- **Gemini**: [aistudio.google.com](https://aistudio.google.com/app/apikey) (Free tier available)
- **OpenAI**: [platform.openai.com](https://platform.openai.com/api-keys)
- **Claude**: [console.anthropic.com](https://console.anthropic.com/)
- **DeepSeek**: [platform.deepseek.com](https://platform.deepseek.com/) (Free tier available)

### 4. Save and Use

1. Save your Apps Script project
2. Refresh your Google Sheet
3. Start using the AI functions in cells

## 🚀 Available Functions

Use these functions directly in any Google Sheets cell:

```excel
=GEMINI("Your question or prompt here")
=OPENAI("Your question or prompt here")
=CLAUDE("Your question or prompt here")
=DEEPSEEK("Your question or prompt here")
```

## 📝 Usage Examples

### Basic Usage

```excel
=GEMINI("What is artificial intelligence?")
=OPENAI("Translate 'Hello world' to Spanish")
=CLAUDE("Write a haiku about programming")
=DEEPSEEK("Explain blockchain in simple terms")
```

### With Cell References

```excel
=GEMINI("Translate this to English: " & A1)
=OPENAI("Write a product description for: " & B2)
=CLAUDE("Summarize this text: " & C3)
=DEEPSEEK("Generate keywords for: " & D4)
```

### Practical Applications

```excel
=GEMINI("Generate 5 marketing slogans for " & A2)
=OPENAI("Create an email subject line for " & B2)
=CLAUDE("Analyze the sentiment of: " & C2)
=DEEPSEEK("Write a brief for: " & D2)
```

## ✨ Features

- ✅ **4 AI Providers**: Gemini, OpenAI, Claude, DeepSeek
- ✅ **Simple Setup**: Just one file to install
- ✅ **Direct Cell Usage**: No menus or complex interfaces
- ✅ **Error Handling**: Clear error messages
- ✅ **Flexible**: Works with cell references and dynamic content

## 🛠️ Troubleshooting

**"API Key not found"**
→ Check Script Properties configuration for the provider you're using

**"Function GEMINI is not defined"**
→ Save the Apps Script project and refresh your Google Sheet

**Rate limit errors**
→ Wait a moment between requests or try a different provider

**Slow responses**
→ Gemini and DeepSeek usually offer faster free tiers

## 🔧 Advanced Configuration

### Default Provider

Set `AI_PROVIDER` in Script Properties to change which provider the generic functions use by default.

### Model Selection

The script uses these models by default:

- **Gemini**: `gemini-1.5-flash`
- **OpenAI**: `gpt-3.5-turbo`
- **Claude**: `claude-3-haiku-20240307`
- **DeepSeek**: `deepseek-chat`

## 📋 Requirements

- Google account with access to Google Sheets
- Google Apps Script access
- API key for at least one AI provider

## 📄 License

MIT License - Created by **Rankgnar**

---

**Simple, powerful AI integration for Google Sheets**
