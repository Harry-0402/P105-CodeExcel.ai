# OpenRouter API Integration Guide

## Overview

CodeExcel AI uses the OpenRouter API to provide access to Google's Gemini 2.0 Flash model, with a free tier available for testing and light use.

## Getting Your API Key

### Step 1: Create OpenRouter Account
1. Visit https://openrouter.ai
2. Click "Sign Up" or "Sign In"
3. Follow the authentication flow

### Step 2: Generate API Key
1. Go to your account settings
2. Navigate to "API Keys" section
3. Click "Create New Key"
4. Copy the generated key

### Step 3: Store API Key Securely
- **Never** hardcode the API key in your source code
- Use the Settings tab in CodeExcel AI to store it securely
- The add-in saves it to browser localStorage
- It persists across sessions but remains local to your browser

## API Model Details

### Default Model
```
google/gemini-2.0-flash-lite-preview-02-05:free
```

**Characteristics:**
- **Free tier**: Yes (limited requests)
- **Latency**: ~2-5 seconds
- **Context Window**: 32,000 tokens
- **Max Output**: 500 tokens (configured in add-in)
- **Temperature**: 0.7 (creative but focused)

### Available Models (Optional)
You can change the model in Settings. Popular alternatives:

- `openai/gpt-3.5-turbo` - Fast, reliable
- `openai/gpt-4` - More capable (paid)
- `anthropic/claude-3-haiku` - Good balance
- `meta-llama/llama-2-70b-chat` - Open source

[View all models at OpenRouter](https://openrouter.ai/models)

## API Endpoint

```
https://openrouter.ai/api/v1/chat/completions
```

## Required Headers

The CodeExcel AI add-in automatically includes:

```javascript
{
    "Authorization": "Bearer YOUR_API_KEY",
    "HTTP-Referer": "https://codeexcel.ai",
    "X-Title": "CodeExcel AI Assistant",
    "Content-Type": "application/json"
}
```

**Important**: 
- `HTTP-Referer` and `X-Title` are required by OpenRouter for rate limiting and identification
- These headers are hardcoded, not configurable
- OpenRouter tracks usage per referer domain

## Request Format

The add-in sends requests in this format:

```json
{
    "model": "google/gemini-2.0-flash-lite-preview-02-05:free",
    "messages": [
        {
            "role": "system",
            "content": "You are an Excel data analyst. Complete the following task or analyze this data concisely."
        },
        {
            "role": "user",
            "content": "The cell value to analyze"
        }
    ],
    "temperature": 0.7,
    "max_tokens": 500,
    "top_p": 1.0,
    "frequency_penalty": 0,
    "presence_penalty": 0
}
```

## Response Format

Success response (HTTP 200):

```json
{
    "id": "...",
    "object": "chat.completion",
    "created": 1234567890,
    "model": "google/gemini-2.0-flash-lite-preview-02-05:free",
    "choices": [
        {
            "finish_reason": "stop",
            "message": {
                "role": "assistant",
                "content": "The AI response text goes here"
            }
        }
    ],
    "usage": {
        "prompt_tokens": 50,
        "completion_tokens": 100,
        "total_tokens": 150
    }
}
```

## Error Responses

### 401 Unauthorized
```json
{
    "error": {
        "message": "Invalid API key",
        "type": "authentication_error"
    }
}
```
**Solution**: Check your API key in Settings tab

### 429 Too Many Requests
```json
{
    "error": {
        "message": "Rate limit exceeded",
        "type": "rate_limit_error"
    }
}
```
**Solution**: Wait before making another request or upgrade your account

### 500 Server Error
```json
{
    "error": {
        "message": "Internal server error",
        "type": "server_error"
    }
}
```
**Solution**: Retry after a few seconds

## Rate Limiting

### Free Tier
- **Requests per minute**: Limited
- **Daily quota**: Varies
- **Concurrent requests**: 1

### Paid Tier
- **Requests per minute**: Higher
- **Daily quota**: Based on subscription
- **Concurrent requests**: Multiple

Check your usage at https://openrouter.ai/account/usage

## Cost (if applicable)

The free model is truly free:
- `google/gemini-2.0-flash-lite-preview-02-05:free` has **$0 cost**
- Other models are pay-per-token

[View pricing at OpenRouter](https://openrouter.ai/#pricing)

## Testing Your API Key

### In Add-in
1. Go to Settings tab
2. Paste your API key
3. Click Save Settings
4. If no error appears, key is valid

### Using cURL (command line)
```bash
curl https://openrouter.ai/api/v1/chat/completions \
  -H "Authorization: Bearer YOUR_API_KEY" \
  -H "HTTP-Referer: https://codeexcel.ai" \
  -H "X-Title: Test" \
  -H "Content-Type: application/json" \
  -d '{
    "model": "google/gemini-2.0-flash-lite-preview-02-05:free",
    "messages": [
      {"role": "user", "content": "Say hello"}
    ]
  }'
```

## Troubleshooting

### Issue: "Error: Invalid API response format"
- API response was malformed
- Check that API key is correct
- Verify model name is valid

### Issue: "Error: Network error: Unable to reach the API"
- CORS issue (browser blocking request)
- No internet connection
- OpenRouter servers are down

**Solution for CORS**: OpenRouter handles CORS if proper headers are sent

### Issue: "Error: API Error: HTTP 401"
- API key is invalid or expired
- Copy key from OpenRouter account settings again
- Re-save in Settings tab

### Issue: Responses are slow
- OpenRouter or Gemini model is overloaded
- Network latency
- Try at off-peak hours

## Advanced Configuration

### Adjusting Temperature

Edit `api.js` to change response creativity:

```javascript
"temperature": 0.7  // 0 = predictable, 1 = creative
```

- **Lower (0.3)**: More focused, analytical (good for data)
- **Higher (0.9)**: More creative, varied

### Changing Max Tokens

Edit `api.js`:

```javascript
"max_tokens": 500  // Increase for longer responses
```

- **Lower**: Faster responses, cheaper
- **Higher**: More detailed responses

### Adjusting Top-P

```javascript
"top_p": 1.0  // 0.9 = more focused, 1.0 = all tokens
```

## Security Notes

1. **Never share your API key** in code, issues, or screenshots
2. **Rotate keys regularly** on OpenRouter account
3. **Monitor usage** to detect unauthorized access
4. **Use HTTPS** in production
5. **Limit requests** to prevent accidental costs

## API Documentation

Full OpenRouter API documentation: https://openrouter.ai/docs

## Support

- OpenRouter Status: https://status.openrouter.io
- OpenRouter Discord: https://discord.gg/openrouter
- Google Gemini Docs: https://ai.google.dev/docs

---

**Last Updated**: January 2026
**API Version**: v1
**SDK**: None (pure fetch API)
