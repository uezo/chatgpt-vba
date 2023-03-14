# chatgpt-vba
ChatGPT API Client for VBA.

# ✨ Features

- Super easy: Just call `ChatGPT.Chat("hello, ChatGPT")` to say hello to ChatGPT👋
- Full API Spec: `ChatGPT.GetCompletion()` takes all ChatGPT parameters except for stream.
- Pure VBA: Run anywhere where VBA runs without any installations.

# 🚀 Quick start

1. Set reference to Microsoft Scripting Runtime.
1. Add ChatGPT.bas and JsonConverter.bas to your VBA Project.
    - https://github.com/uezo/chatgpt-vba/releases
1. Make your script and run.

```vb
Sub Main()
    ChatGPT.ApiKey = "YOUR_API_KEY"
    Debug.Print ChatGPT.Chat("What is the difference between a dolphin and a whale?")
End Sub
```

Or, use `GetCompletion()` to call API with other parameters.

```vb
Sub Main()
    ChatGPT.ApiKey = "YOUR_API_KEY"

    Dim messages() As Dictionary
    messages = ChatGPT.CreateMessages("You are biologist.", "What is the difference between a dolphin and a whale?")
    
    Dim completion As String
    completion = ChatGPT.GetCompletion(messages, maxTokens:=1000, temperature:=0.5)
    
    Debug.Print completion
End Sub
```

# 📚 API Reference

See OpenAI API Reference.

https://platform.openai.com/docs/api-reference/chat/create

# ❤️ Thanks

[VBA-JSON](https://github.com/VBA-tools/VBA-JSON) by Tim Hall, JSON converter for VBA helps me a lot to make HTTP client and this awesome library is included in the release under its license. Thank you!
