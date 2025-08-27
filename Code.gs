// Google Apps Script AI Integration for Sheets and Docs
// Author: Rankgnar
// Support for DeepSeek, OpenAI, Claude, and more

// API Configuration
const API_URLS = {
  gemini: "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent",
  deepseek: "https://api.deepseek.com/chat/completions",
  openai: "https://api.openai.com/v1/chat/completions",
  claude: "https://api.anthropic.com/v1/messages"
};

/**
 * Function executed when opening a document or spreadsheet.
 * Creates custom menu to access AI assistant functions.
 */
function onOpen(e) {
  // This script only works with Google Sheets using functions like =GEMINI(), =OPENAI(), etc.
}

//=================================================================================
// CORE AI API FUNCTIONS
//=================================================================================

/**
 * Get API key for specified provider from script properties
 * @param {string} provider - AI provider name
 * @returns {string} API key
 */
function getApiKey_(provider) {
  const keyProperty = provider.toUpperCase() + '_API_KEY';
  const apiKey = PropertiesService.getScriptProperties().getProperty(keyProperty);
  if (!apiKey) {
    throw new Error(`API Key for ${provider} not found. Please configure it in Script Properties.`);
  }
  return apiKey;
}

/**
 * Get current AI provider from script properties
 * @returns {string} Current AI provider
 */
function getCurrentProvider_() {
  return PropertiesService.getScriptProperties().getProperty('AI_PROVIDER') || 'gemini';
}

/**
 * Make API call to specified AI provider
 * @param {string} prompt - User prompt
 * @param {string} provider - AI provider to use
 * @returns {string} AI response
 */
function callAI(prompt, provider = null) {
  if (!prompt || prompt.toString().trim() === "") {
    return "Error: Prompt cannot be empty.";
  }

  // Convert prompt to string if it's not already
  const promptText = prompt.toString().trim();
  const aiProvider = provider || getCurrentProvider_();
  
  try {
    const apiKey = getApiKey_(aiProvider);
    const apiUrl = API_URLS[aiProvider];
    
    if (!apiUrl) {
      return `Error: Unsupported AI provider: ${aiProvider}`;
    }

    let headers, payload;

    // Configure request based on provider
    switch (aiProvider) {
      case 'gemini':
        headers = {
          'Content-Type': 'application/json'
        };
        payload = {
          "contents": [
            {
              "parts": [
                {"text": promptText}
              ]
            }
          ]
        };
        break;

      case 'deepseek':
      case 'openai':
        headers = {
          'Authorization': 'Bearer ' + apiKey,
          'Content-Type': 'application/json'
        };
        payload = {
          "model": aiProvider === 'deepseek' ? "deepseek-chat" : "gpt-3.5-turbo",
          "messages": [
            {"role": "user", "content": promptText}
          ]
        };
        break;

      case 'claude':
        headers = {
          'x-api-key': apiKey,
          'Content-Type': 'application/json',
          'anthropic-version': '2023-06-01'
        };
        payload = {
          "model": "claude-3-haiku-20240307",
          "max_tokens": 1000,
          "messages": [
            {"role": "user", "content": promptText}
          ]
        };
        break;

      default:
        return `Error: Unsupported provider: ${aiProvider}`;
    }

    const options = {
      'method': 'post',
      'headers': headers,
      'payload': JSON.stringify(payload),
      'muteHttpExceptions': true
    };

    // Add API key to URL for Gemini
    const finalUrl = aiProvider === 'gemini' ? `${apiUrl}?key=${apiKey}` : apiUrl;
    const response = UrlFetchApp.fetch(finalUrl, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      try {
        const jsonResponse = JSON.parse(responseBody);
        
        // Parse response based on provider
        let text;
        switch (aiProvider) {
          case 'gemini':
            text = jsonResponse?.candidates?.[0]?.content?.parts?.[0]?.text;
            if (!text && jsonResponse?.candidates?.[0]?.finishReason) {
              return `Error: ${jsonResponse.candidates[0].finishReason}`;
            }
            break;
          case 'deepseek':
          case 'openai':
            text = jsonResponse?.choices?.[0]?.message?.content;
            break;
          case 'claude':
            text = jsonResponse?.content?.[0]?.text;
            break;
        }
        
        return text ? text.trim() : "Error: No response text found in API response.";
      } catch (parseError) {
        return `Error parsing response: ${parseError.message}`;
      }
    } else {
      try {
        const errorResponse = JSON.parse(responseBody);
        const errorMessage = errorResponse?.error?.message || errorResponse?.message || responseBody;
        return `Error ${responseCode}: ${errorMessage}`;
      } catch (e) {
        return `Error ${responseCode}: ${responseBody}`;
      }
    }

  } catch (e) {
    return "Error: " + e.message;
  }
}

//=================================================================================
// GOOGLE SHEETS FUNCTIONS
//=================================================================================

/**
 * Gemini-specific function
 * @param {string} prompt - Text to send to Gemini
 * @returns Gemini response
 * @customfunction
 */
function GEMINI(prompt) {
  return callAI(prompt, 'gemini');
}

/**
 * DeepSeek-specific function
 * @param {string} prompt - Text to send to DeepSeek
 * @returns DeepSeek response
 * @customfunction
 */
function DEEPSEEK(prompt) {
  return callAI(prompt, 'deepseek');
}

/**
 * OpenAI-specific function
 * @param {string} prompt - Text to send to OpenAI
 * @returns OpenAI response
 * @customfunction
 */
function OPENAI(prompt) {
  return callAI(prompt, 'openai');
}

/**
 * Claude-specific function
 * @param {string} prompt - Text to send to Claude
 * @returns Claude response
 * @customfunction
 */
function CLAUDE(prompt) {
  return callAI(prompt, 'claude');
}


