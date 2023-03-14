Attribute VB_Name = "ChatGPT"
Option Explicit

Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public ApiKey As String

' Simple interface for GetCompletion
Public Function Chat(userMessage As String, Optional promptMessage As String, Optional maxTokens As Integer = 1000) As String
    Dim messages() As Dictionary

    If promptMessage <> "" Then
        messages = CreateMessages(promptMessage, userMessage)
    Else
        ReDim messages(0)
        Set messages(0) = New Dictionary
        messages(0).Add "role", "user"
        messages(0).Add "content", userMessage
    End If
    
    Chat = GetCompletion(messages, maxTokens:=maxTokens)
End Function

' Call OpenAI Chat.Completion API
Public Function GetCompletion(messages As Variant, _
    Optional model As String = "gpt-3.5-turbo", _
    Optional temperature As Single = 1, _
    Optional topP As Single = 1, _
    Optional n As Integer = 1, _
    Optional stopWords As Variant, _
    Optional maxTokens As Integer = 16, _
    Optional presencePenalty As Single = 0, _
    Optional frequencyPenalty As Single = 0, _
    Optional logitBias As Variant, _
    Optional user As String, _
    Optional asText As Boolean = True _
    ) As Variant
    
    Dim data As New Dictionary
    data.Add "messages", messages
    data.Add "model", model
    data.Add "temperature", temperature
    data.Add "top_p", topP
    data.Add "n", n
    If Not IsMissing(stopWords) Then
        data.Add "stop", stopWords
    End If
    data.Add "max_tokens", maxTokens
    data.Add "presence_penalty", presencePenalty
    data.Add "frequency_penalty", frequencyPenalty
    If Not IsMissing(logitBias) Then
        data.Add "logit_bias", logitBias
    End If
    If user <> "" Then
        data.Add "user", user
    End If
    
    Dim client As Object
    Set client = CreateObject("MSXML2.ServerXMLHTTP")
    client.setTimeouts 30000, 30000, 30000, 60000
    client.Open "POST", "https://api.openai.com/v1/chat/completions"
    client.setRequestHeader "Content-Type", "application/json"
    client.setRequestHeader "Authorization", "Bearer " & ApiKey
    client.send JsonConverter.ConvertToJson(data)

    Do While client.readyState < 4
        Sleep 1
        DoEvents
    Loop

    Dim completion As Dictionary
    Set completion = JsonConverter.ParseJson(client.responseText)
    
    If completion.Exists("error") Then
        Err.Raise 9001, Description:=completion("error")("message")
    End If

    If asText Then
        GetCompletion = completion("choices")(1)("message")("content")
    Else
        Set GetCompletion = completion
    End If
End Function

' Create system and user messages
Public Function CreateMessages(systemContent As String, userContent As String) As Dictionary()
    Dim messages(1) As Dictionary
    
    Set messages(0) = New Dictionary
    messages(0).Add "role", "system"
    messages(0).Add "content", systemContent
    Set messages(1) = New Dictionary
    messages(1).Add "role", "user"
    messages(1).Add "content", userContent
    
    CreateMessages = messages
End Function
