# Outlook AI Add-in

AI-powered Outlook add-in that automatically drafts email replies using ChatGPT.

## Features

- 🤖 **AI-Powered Responses**: Uses ChatGPT to generate professional email replies
- 📧 **Automatic Reply All**: Opens Reply All window with pre-populated AI draft
- 🔒 **Secure API Key Storage**: Your API key is stored securely in Office settings
- 🎯 **Read Mode Only**: Only works when reading emails (not in compose mode)
- ✏️ **Calibri 12pt Font**: Consistent formatting with standard email appearance
- 📝 **Concise Responses**: Generates brief, professional replies by default
- 🚫 **No Signatures**: Avoids adding signatures since Outlook adds them automatically

## Installation

1. **Get the manifest file**: Use the GitHub Pages URL for the manifest:
   ```
   https://mazzproduction.github.io/outlook-ai-addin/manifest.xml
   ```

2. **Install in Outlook**:
   - **Outlook on the Web**: Go to Settings → Add-ins → Custom add-ins → Add from file
   - **Outlook Desktop**: File → Manage Add-ins → My add-ins → Custom add-ins → Add from file

3. **Configure API Key**:
   - First time you use the add-in, it will prompt for your ChatGPT API key
   - Your API key is stored securely and never shared

## Usage

1. **Open an email** in read mode (not compose mode)
2. **Click the add-in button** in the ribbon
3. **Optionally add instructions** in the text box for specific requirements
4. **Click "Draft by AI"**
5. **Reply All window opens** automatically with the AI-generated draft

## Security

- ✅ Your API key is stored locally in Office settings
- ✅ No API keys are stored in the repository or transmitted to third parties
- ✅ All communication is direct between your Outlook and the AI service

## Requirements

- Microsoft Outlook (Web or Desktop)
- ChatGPT API key (from OpenAI)
- Internet connection

## Support

For issues or questions, please create an issue in this repository.

---

**Developed by Jazz Ma**  
*University of Hong Kong*
