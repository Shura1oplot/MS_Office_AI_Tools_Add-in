Attribute VB_Name = "AITools"
' Version 2024-04-11+1

' References (all):
' - Microsoft Scripting Runtime
' - Microsoft WinHTTP Services, version 5.1
' - Microsoft ActiveX Data Objects 6.1 Library
' - Microsoft Forms 2.0 Object Library

Option Explicit

' ==============================================================================
' Preprocessor Constants
' ==============================================================================

#Const IsPowerPoint = True
#Const IsWord = False
#Const IsExcel = False

#Const DeveloperMode = False

' ==============================================================================
' Imports
' ==============================================================================

#If IsPowerPoint Then

' Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
Public Declare PtrSafe Function WaitMessage Lib "user32" () As Long

#End If

' ==============================================================================
' Global Variables
' ==============================================================================

Private DefaultModel As String

#If IsExcel Then
Private AI_Cache As Dictionary
#End If

' ==============================================================================
' Main LLM Settings
' ==============================================================================

Private Function GetDefaultModel() As String
    GetDefaultModel = GetSetting("AI tools", "Settings", "Model", "")

    If IsEmpty(GetDefaultModel) Or GetDefaultModel = "" Then
        GetDefaultModel = "chatgpt"
    End If
End Function

Private Function GetProvider(Optional model As String) As String
    Dim provider As String

    If IsEmpty(model) Or model = "" Then
        provider = "openrouter"
    ElseIf model = "chatgpt" Or StartsWith(model, "gpt-") Then
        provider = "openai"
    ElseIf model = "claude" Or StartsWith(model, "claude-") Then
        provider = "anthropic"
    ElseIf model = "gemini" Or StartsWith(model, "gemini-") Then
        provider = "google"
    ElseIf model = "mistral" Or StartsWith(model, "mistral-") Then
        provider = "mistralai"
    ElseIf model = "command-r" Or model = "command-r-plus" Then
        provider = "cohere"
    Else
        provider = "openrouter"
    End If

    If provider <> "openrouter" And GetAPIKey(provider, safe:=True) = "" Then
#If DeveloperMode Then
        Debug.Print "No API key for '" & provider & "'; fallback to 'openrouter'"
#End If
        provider = "openrouter"
    End If

    GetProvider = provider
End Function

Private Function GetModelName(model As String, _
                              provider As String) _
                              As String
    ' https://platform.openai.com/docs/models
    If provider = "openai" Then
        If model = "gpt-4" Or model = "gpt-4-turbo" Or model = "chatgpt" Then
            GetModelName = "gpt-4-turbo"
        ElseIf model = "gpt-3" Or model = "gpt-3.5" Then
            GetModelName = "gpt-3.5-turbo-instruct"
        Else
            Err.Raise vbObjectError + 1001, , "Wrong model"
        End If

    ' https://docs.mistral.ai/platform/endpoints/
    ElseIf provider = "mistralai" Then
        If model = "mistral" Then
            GetModelName = "mistral-large-latest"
        ElseIf model = "mistral-medium" Then
            GetModelName = "mistral-medium-latest"
        ElseIf model = "mistral-large" Then
            GetModelName = "mistral-large-latest"
        Else
            Err.Raise vbObjectError + 1001, , "Wrong model"
        End If

    ' https://console.anthropic.com/
    ElseIf provider = "anthropic" Then
        If model = "claude" Or model = "claude-3" Then
            GetModelName = "claude-3-opus-20240229"
        Else
            Err.Raise vbObjectError + 1001, , "Wrong model"
        End If

    ' https://openrouter.ai/docs#models
    ElseIf provider = "openrouter" Then
        ' OpenAI
        If model = "chatgpt" Or StartsWith(model, "gpt-") Then
            GetModelName = "openai/gpt-4-turbo-preview"

        ' Anthropic
        ElseIf model = "claude" Or model = "claude-3" Or model = "claude-3-opus" Then
            GetModelName = "anthropic/claude-3-opus"
        ElseIf model = "claude-2" Then
            GetModelName = "anthropic/claude-2"

        ' Google
        ElseIf model = "gemini" Or model = "gemini-pro" Or model = "gemini-1.0-pro" _
                Or model = "gemini-pro-latest" Then
            GetModelName = "google/gemini-pro"
        ElseIf StartsWith(model, "gemini-") Then
            GetModelName = "google/" & model

        ' Mistral
        ElseIf model = "mistral" Then
            GetModelName = "mistralai/mistral-large"
        ElseIf StartsWith(model, "mistral-") Then
            GetModelName = "mistralai/" & model

        ' Command R/R+
        ElseIf model = "command-r" Or model = "command-r-plus" Then
            GetModelName = "cohere/" & model
        Else
            Err.Raise vbObjectError + 1001, , "Wrong model"
        End If

    ' https://ai.google.dev/models/gemini
    ElseIf provider = "google" Then
        If model = "gemini" Or model = "gemini-pro" Then
            GetModelName = "gemini-1.0-pro"
        ElseIf model = "gemini-pro-latest" Then
            GetModelName = "gemini-1.0-pro-latest"
        Else
            Err.Raise vbObjectError + 1001, , "Wrong model"
        End If

    ' https://docs.cohere.com/docs/models
    ElseIf provider = "cohere" Then
        If model = "command-r" Or model = "command-r-plus" Then
            GetModelName = model
        Else
            Err.Raise vbObjectError + 1001, , "Wrong model"
        End If

    Else
        Err.Raise vbObjectError + 1001, , "Wrong provider"

    End If
End Function

' OpenAI-like API
Private Function GetBaseURL(provider As String) As String
    If provider = "openai" Then
        GetBaseURL = "https://api.openai.com/v1/chat/completions"

    ElseIf provider = "openrouter" Then
        GetBaseURL = "https://openrouter.ai/api/v1/chat/completions"

    ElseIf provider = "mistralai" Then
        GetBaseURL = "https://api.mistral.ai/v1/chat/completions"

    Else
        Err.Raise vbObjectError + 1001, , "Wrong provider"
    End If
End Function

Private Function GetDefaultPreamble() As String
    GetDefaultPreamble = _
        ("You are an AI-driven Microsoft Office add-in, which helps management consultants to prepare business presentations " & _
         "and documents. A user will provide you with a command or ask you a question. Please respond in the most precise way you " & _
         "can without further clarifications of the input. Be concise and to the point. Do not make up facts. Follow the " & _
         "instructions provided. If you are provided with a text to work with, make you response based only on the text provided " & _
         "and without any additions (if the opposit is not clearly stated by the user). Use business language whenever possible " & _
         "unless otherwise stated.")
End Function

' ==============================================================================
' Main Macros
' ==============================================================================

Sub CorrectToStandardEnglish()
    Dim command As String
    command = ("You are a spell checker. Correct the input text delimited by triple quotes ("""""") to standard English. " & _
               "If the input text is in standard English, return it as it is. Avoid decoding any abbreviations, " & _
               "instead preserve them as they are. Try to preserve the length of the input text. " & _
               "Pay close attention to the usage of articles and prepositions. " & _
               "Wrap the result text in triple quotes ("""""") as well." & vbLf & vbLf & _
               "Input text:" & vbLf & _
               """""""{{input}}""""""")
    TransformSelection command:=command, _
                       temperature:=0
End Sub

Sub CorrectToStandardEnglishBusiness()
    Dim command As String
    command = ("You are a professional linguist. Rephrase the input text delimited by triple quotes ("""""") " & _
               "to correct it to standard English in a business style. If the input text is in standard English " & _
               "and follow business style, return it as it is. Avoid decoding any abbreviations, instead " & _
               "preserve them as they are. Try to preserve the length of the input text. The result text " & _
               "should be clear and concise. Pay close attention to the usage of articles and prepositions. " & _
               "Wrap the rephrased text in triple quotes ("""""") as well." & vbLf & vbLf & _
               "Input text:" & vbLf & _
               """""""{{input}}""""""")
    TransformSelection command:=command, _
                       temperature:=0.1
End Sub

Sub ParaphraseShorten()
    Dim command As String
    command = ("You are a professional linguist. Paraphrase the input text delimited by triple quotes ("""""") " & _
               "to reduce its length by quarter or half, but preserve its core meaning and key messages. " & _
               "The output text should be in standard English in a business style, clean and concise. " & _
               "Wrap the paraphrased text in triple quotes ("""""") as well." & vbLf & vbLf & _
               "Input text:" & vbLf & _
               """""""{{input}}""""""")
    TransformSelection command:=command, _
                       temperature:=0.3
End Sub

#If IsPowerPoint Then

Sub RephraseConsultingZeroShot()
    Dim command As String
    command = ("You are a professional linguist. Rephrase the input text delimited by triple quotes " & _
               "("""""") using the style that the leading management consulting company, McKinsey, " & _
               "uses in its articles and presentations. " & _
               "Wrap the rephrased text in triple quotes ("""""") as well." & vbLf & vbLf & _
               "Input text:" & vbLf & _
               """""""{{input}}""""""")
    TransformSelection command:=command, _
                       temperature:=0.5
End Sub

Sub RephraseConsultingMultiShot()
    Dim s As String

    ' Don't call them titles as LLMs understand "titles" in a more common way (non actionable)
    s = ""
    s = s + "You are a professional linguist. "
    s = s + "The following extracts were taken from a business presentation. "
    s = s + "You need to rephrase an extract to change its style to the style usually used by management consultants. "
    s = s + "As guidelines use the style used by top consulting firms, like McKinsey, BCG, and Bain. "
    s = s + "Preserve the meaning and all key messages of the source extract. "
    s = s + "Do not convert the extract to the format of the presentation title. "
    s = s + "The lenght of the rephrased extract should close to the length of the source one. "
    s = s + "The input text is in the ""SOURCE:"" field wrapped in tripple quotes (""""""), "
    s = s + "the rephrased text should be in the ""RESULT:"" field and also wrapped in tripple quotes (""""""), " & vbLf & vbLf
    s = s + "# Examples:" & vbLf & vbLf
    s = s + "SOURCE: """"""We conducted the benchmark exercise in 5 steps to select and study the most digitally advanced and the most relevant to CLIENT companies.""""""" & vbLf
    s = s + "RESULT: """"""We have followed a five-tiered tailored approach to select and benchmark the most digitally advanced and significant companies in the world.""""""" & vbLf
    s = s + "" & vbLf & vbLf
    s = s + "SOURCE: """"""We collected many open indexes and ratings from various sources to select the most relevant entities for the benchmarking.""""""" & vbLf
    s = s + "RESULT: """"""To ensure accuracy in our benchmarking exercise, we made sure to select the entities based on the most significant global sources within the CLIENT's industry.""""""" & vbLf
    s = s + "" & vbLf & vbLf
    s = s + "SOURCE: """"""We studied national, sector-specific, and CLIENT's Corporate strategies to align the benchmarking with Saudi Arabia aspirations.""""""" & vbLf
    s = s + "RESULT: """"""We rigorously aligned our criteria with national, sector-specific, and CLIENT's key performance indicators (KPIs) and strategic aspirations.""""""" & vbLf
    s = s + "" & vbLf & vbLf
    s = s + "SOURCE: """"""To focus on the most relevant benchmarking candidates we screened the long list of 100+ companies using two filters: its specialization and annual operations throughput.""""""" & vbLf
    s = s + "RESULT: """"""We focused our screening criteria on assessing a list of over 100 ports based on two strategic key pillars: specialization and size of business.""""""" & vbLf
    s = s + "" & vbLf & vbLf
    s = s + "SOURCE: """"""Container ports were selected for the benchmarking as this type of cargo is in the focus of the national and sector-specific strategies.""""""" & vbLf
    s = s + "RESULT: """"""�Containers� present promising growth opportunities for CLIENT and is a key focus area for the national and sector-specific strategies.""""""" & vbLf
    s = s + "" & vbLf & vbLf
    s = s + "SOURCE: """"""Based on CLIENT�s and sectoral KPIs we defined selection criteria and developed a scoring model to select 5 target entities for the benchmarking exercise.""""""" & vbLf
    s = s + "RESULT: """"""We developed a scoring model that factored in national, sectorial and organizational KPIs and aspirations to shortlist 5 entities for benchmarking.""""""" & vbLf
    s = s + "" & vbLf & vbLf
    s = s + "SOURCE: """"""Based on 6 selection criteria we defined 10 numeric parameters for our scoring model, which rules are based on current and target CLIENT�s and sectoral KPIs.""""""" & vbLf
    s = s + "RESULT: """"""Our approach involved leveraging those KPIs as the foundation of our model, while implementing a scoring mechanism that encompasses additional factors.""""""" & vbLf
    s = s + "" & vbLf & vbLf
    s = s + "SOURCE: """"""Finally, 5 entities were selected for the benchmarking exercise: 4 with the highest score points, and one additional as the closest competitor of CLIENT in the Middle East.""""""" & vbLf
    s = s + "RESULT: """"""Five ports are strategically selected including the top-performing four companies as per our model, as well as the closest regional competitor.""""""" & vbLf
    s = s + "" & vbLf & vbLf
    s = s + "SOURCE: """"""The benchmarking will focus on analysis of digitalization experience of port authorities and port regulators in alignment with CLIENT�s operating model.""""""" & vbLf
    s = s + "RESULT: """"""To ensure our analysis is aligned with CLIENT�s current and future strategic plans, we will tailor our assessment to focus on key roles in the supply chain.""""""" & vbLf
    s = s + "" & vbLf & vbLf
    s = s + "SOURCE: """"""For a comprehensive benchmarking exercise, six dimensions were defined for the benchmarking to address important questions of the CLIENT�s Digital Strategy.""""""" & vbLf
    s = s + "RESULT: """"""To ensure our benchmarking exercise is consistent with the current state, we have defined six critical dimensions that will significantly impact CLIENT�s digital future.""""""" & vbLf
    s = s + "" & vbLf & vbLf
    s = s + "# Your task:" & vbLf & vbLf
    s = s + "SOURCE: """"""{{input}}""""""" & vbLf & vbLf
    s = s + "Do not repeat the source text, write only the rephrased extract starting with ""RESULT:""."

    TransformSelection command:=s, _
                       temperature:=0.5, _
                       remove_prefix:="RESULT:"
End Sub

#End If

#If IsPowerPoint Or IsWord Then

Sub RephrasePoliteConcise()
    Dim command As String

    command = ("Rewrite the input message delimited by triple quotes ("""""") to be more indirect, " & _
               "polite, delicate, and considerate. Keep it concise and to the point, and make " & _
               "it suitable for communication with executive-level management. " & _
               "Ensure that the cultural sensibilities and professional etiquettes " & _
               "common in Arab and European contexts are considered, " & _
               "maintaining the original message's intent. " & _
               "Wrap the result text in triple quotes ("""""") as well." & vbLf & vbLf & _
               "Input text:" & vbLf & _
               """""""{{input}}""""""")
    TransformSelection command:=command, _
                       temperature:=0.2
End Sub

Sub RephrasePoliteExtra()
    Dim command As String

    command = ("Take the input message delimited by triple quotes ("""""") and rewrite it " & _
               "to be more indirect, polite, delicate, and considerate. Emphasize " & _
               "readiness to collaborate and showing respect for the " & _
               "recipient's time and efforts if suitable. Ensure that " & _
               "the cultural sensibilities and professional etiquettes " & _
               "common in Arab and European contexts are considered, " & _
               "maintaining the original message's intent. " & _
               "Wrap the result text in triple quotes ("""""") as well." & vbLf & vbLf & _
               "Input text:" & vbLf & _
               """""""{{input}}""""""")
    TransformSelection command:=command, _
                       temperature:=0.2, _
                       correct_punctuation:=True
End Sub

#End If

' ==============================================================================
' Run AI for Playground
' ==============================================================================

#If IsPowerPoint Then

Sub RunAI()
    Dim command As String
    Dim source As String
    Dim temperature As Double
    Dim resultShape As Shape

    Dim i As Integer
    Dim slide_tagged As Boolean
    slide_tagged = False

    With ActiveWindow.View.Slide
        For i = 1 To .Tags.Count
            If .Tags.Name(i) = "RUNAIPLAYGROUND" Then
                slide_tagged = True
                Exit For
            End If
        Next i

        If Not slide_tagged Then
            MsgBox ("Cannot run AI on this slide. The command works only on " + _
                    "the playground slide and its copies."), vbCritical, "AI Tools: Error"
            Exit Sub
        End If

        command = GetText(.Shapes("_command_"))
        source = GetText(.Shapes("_source_"))
        temperature = CDbl(GetText(.Shapes("_temperature_"))) / 100
        Set resultShape = .Shapes("_result_")
    End With

    Dim result As String
    result = LLMTextCommand(command:=command, _
                            source:=source, _
                            temperature:=temperature)

    If StartsWith(result, """""""") And EndsWith(result, """""""") Then
        result = RemoveSuffix(RemovePrefix(result, """"""""), """""""")
    End If

    resultShape.TextFrame.TextRange.text = result
End Sub

#End If

' ==============================================================================
' Rephrase Title Macro
' ==============================================================================

#If IsPowerPoint Then

Sub RephraseTitleVariants()
    Dim tr As TextRange
    Dim source As String
    Dim result As String
    Dim token As String
    Dim base_url As String
    Dim url As String
    Dim error_text As String
    Dim request As String
    Dim http As WinHttp.WinHttpRequest
    Dim timeout As Long
    Dim response_str As String
    Dim response_json As Object
    Dim task_id As String
    Dim status As String
    Dim timeout_counter As Integer
    ' Dim ts As Single
    Dim text As String
    Dim i As Long

#If IsWord Then
    Dim d As Scripting.Dictionary
#Else
    Dim d As Dictionary
#End If

    error_text = ("Invalid selection. Please select a text fragment in a " & _
                  "text field, or select one shape with text.")

    With ActiveWindow.Selection
        If .Type = ppSelectionText Then
            Set tr = .TextRange
        ElseIf .Type = ppSelectionShapes Then
            If .ShapeRange.Count > 1 Then
                MsgBox error_text, vbCritical, "AI Tools: Error"
                Exit Sub
            ElseIf .ShapeRange.Count = 0 Then
                MsgBox error_text, vbCritical, "AI Tools: Error"
                Exit Sub
            End If
            With .ShapeRange(1)
                If Not .HasTextFrame Then
                    MsgBox error_text, vbCritical, "AI Tools: Error"
                    Exit Sub
                End If
                If Not .TextFrame.HasText Then
                    MsgBox error_text, vbCritical, "AI Tools: Error"
                    Exit Sub
                End If
                Set tr = .TextFrame.TextRange
            End With
        Else
            MsgBox error_text, vbCritical, "AI Tools: Error"
            Exit Sub
        End If
    End With

    source = tr.text

Begin:  ' ##################

    AIToolsWait.Show
    DoEvents

    token = "502f3434-1d27-4161-a703-b015f5ae787d"
    base_url = "https://ensemble-58c17250eaa2.herokuapp.com"

#If IsWord Then
    Set d = New Scripting.Dictionary
#Else
    Set d = New Dictionary
#End If

    d.Add "text", source

    request = ConvertToJson(d)

    Set http = New WinHttpRequest

    timeout = 10000  ' ms
    http.SetTimeouts timeout, timeout, timeout, timeout

    url = base_url & "/api/v1/ensemble/rephrase_title/create"

#If DeveloperMode Then
    Debug.Print ">>>>>>>>>>>>"
    Debug.Print "URL:", url
    Debug.Print "Request:", request
#End If

    http.Open "POST", url
    http.SetRequestHeader "Authorization", "Bearer " & token
    http.SetRequestHeader "Content-Type", "application/json"

#If Not DeveloperMode Then
    On Error GoTo HTTPErrorHandler
#End If
    http.Send request
    On Error GoTo 0

    response_str = http.ResponseText
    response_str = DecodeText(response_str, "ISO-8859-1", "UTF-8")

#If DeveloperMode Then
    Debug.Print "Response:", response_str
    Debug.Print "<<<<<<<<<<<<"
#End If

    Set response_json = ParseJson(response_str)

    task_id = response_json("task_id")

    url = base_url & "/api/v1/ensemble/rephrase_title/retrieve?task_id=" & task_id

    status = ""

    timeout_counter = 300

    Do While status <> "completed"
        ' Variant 1
        ' Freezes the application
        'Sleep 1000  ' 1 sec

        ' Variant 2
        ' Loads CPU heavily
        'ts = Timer()
        'Do While Timer() <= ts + 1
        '    DoEvents
        'Loop

        ' Variant 3
        ' Not available in PowerPoint
        'DoEvents
        'Application.Wait (Now + TimeValue("0:00:01"))

        ' Variant 4
        ' Best options so far
        UnblockingWait 1

#If DeveloperMode Then
        Debug.Print ">>>>>>>>>>>>"
        Debug.Print "URL:", url
#End If

        http.Open "GET", url
        http.SetRequestHeader "Authorization", "Bearer " & token
        http.SetRequestHeader "Content-Type", "application/json"

#If Not DeveloperMode Then
        On Error GoTo HTTPErrorHandler
#End If
        http.Send
        On Error GoTo 0

        response_str = http.ResponseText
        response_str = DecodeText(response_str, "ISO-8859-1", "UTF-8")

#If DeveloperMode Then
        Debug.Print "Response:", response_str
        Debug.Print "<<<<<<<<<<<<"
#End If

        Set response_json = ParseJson(response_str)

        status = response_json("status")

        If status = "error" Then
            Err.Raise vbObjectError + 1001, , _
                "API Error: " & response_json("error") & ". Press 'End' and try again."
        End If

        If status <> "created" And status <> "in_progress" _
                And status <> "completed" Then
            Err.Raise vbObjectError + 1001, , _
                "API Error: Unknown status '" & status & "'. Press 'End' and try again."
        End If

        timeout_counter = timeout_counter - 1

        If timeout_counter = 0 Then
            Err.Raise vbObjectError + 1001, , _
                "Error: Timeout. Press 'End' and try again."
        End If
    Loop

    Do While AIToolsWait.Visible
        AIToolsWait.Hide
        DoEvents
    Loop

    text = "Orig: " + source + vbCr + vbCr

    For i = 1 To response_json("result")("output").Count
        text = text + response_json("result")("output")(i)
        text = text + vbCr
        text = text + "By: " + response_json("result")("model")(i)
        text = text + vbCr + vbCr + vbCr
    Next i

    text = Trim_(text)

    With AIToolsOutput
        .Tag = ""
        .TextBoxOutput.text = text
        .Show  ' Blocking
        result = ""
        On Error Resume Next
        result = .Tag
        On Error GoTo 0
    End With

    If Not IsEmpty(result) And result <> "" Then
        tr.text = result
    End If

    Exit Sub

HTTPErrorHandler:
    error_text = Err.Description

    On Error GoTo 0

#If DeveloperMode Then
    Debug.Print "HTTP Error:", error_text
    Debug.Print "<<<<<<<<<<<<"
#End If

    Do While AIToolsWait.Visible
        AIToolsWait.Hide
    Loop

    If IsEmpty(error_text) Or error_text = "" Then
        error_text = "Unknown"
    End If

    Err.Raise vbObjectError + 1001, , _
        "HTTP Error: " & error_text & ". Press 'End' and try again."
End Sub

#End If

' ==============================================================================
' Excel UDF
' ==============================================================================

#If IsExcel Then

#If DeveloperMode Then

Sub ZTestAI()
    AI 1, ActiveWindow.Selection
End Sub

#End If

Function AI(mode As Integer, _
            input_data As Range) As Variant
    ' Mode = 1 -> Add new rows
    ' Mode = 2 -> Fill missing values in a column

    ' Static AI_Cache As Dictionary

    Dim out_arr() As Variant
    Dim cclb As Long, ccub As Long
    Dim guidance As String
    Dim command As String
    Dim result As String
    Dim result_rows() As String
    Dim cached As Boolean

    Dim i As Long, j As Long
    Dim s As String
    Dim a() As String

    If AI_Cache Is Nothing Then
        Set AI_Cache = New Dictionary
    End If

    guidance = _
        ("You are a VBA User Defined Function (UDF) in Excel. " & _
         "A user will provide you an input table with a request, a command or a question. " & _
         "Please respond in the most precise way you can without further clarifications of the input. " & _
         "Be concise and to the point. Do not make up facts. Follow the instructions provided. " & _
         "Your response should be formated as a table: each line is a table row, and columns " & _
         "should be split by a vertical bar (|). Do not repeat user's input. " & _
         "Use business language whenever possible unless otherwise stated." & vbLf & vbLf)

    If mode = 1 Then
        guidance = guidance & _
                  ("You should add new rows to the table. Please follow the pattern provided by the user " & _
                   "(an instruction, a table headers, and sample rows). User's input formated as a Markdown " & _
                   "table. Your should output only new rows of this table." & vbLf & vbLf)

        guidance = guidance & _
                  ("Example 1" & vbLf & _
                   "---------" & vbLf & vbLf & _
                   "User:" & vbLf & _
                   "|List the last 5 Olympic champions in figure skating|||" & vbLf & _
                   "|Name|Year|Country|" & vbLf & vbLf & _
                   "Assistant:" & vbLf & _
                   "|Yuzuru Hanyu|2014|Japan|" & vbLf & _
                   "|Evan Lysacek|2010|United States|" & vbLf & _
                   "|Evgeni Plushenko|2006|Russia|" & vbLf & _
                   "|Alexei Yagudin|2002|Russia|")

        guidance = guidance & vbLf & vbLf & vbLf

        guidance = guidance & _
                  ("Example 2" & vbLf & _
                   "---------" & vbLf & vbLf & _
                   "User:" & vbLf & _
                   "|List the last 5 US presidents|" & vbLf & vbLf & _
                   "Assistant:" & vbLf & _
                   "|Joe Biden|" & vbLf & _
                   "|Donald Trump|" & vbLf & _
                   "|Barack Obama|" & vbLf & _
                   "|George W. Bush|" & vbLf & _
                   "|Bill Clinton|")

        command = RangeToText(input_data)

    ElseIf mode = 2 Then
        guidance = guidance & _
                  ("You should fill missing values (marked as '?' in the input table). " & _
                   "Replace '?' with appropriate values and write updated table. " & _
                   "User's input formated as a Markdown table. " & _
                   "Do not repeat the instructions provided, output only an updated table with the " & _
                   "header row if available." & vbLf & vbLf)

        guidance = guidance & _
                  ("Example 1" & vbLf & _
                   "---------" & vbLf & vbLf & _
                   "User:" & vbLf & _
                   "Capitals of countries and currencies (codes)" & vbLf & _
                   "|Country|Capital|Currency|" & vbLf & _
                   "|Italy|?|?|" & vbLf & _
                   "|UAE|?|?|" & vbLf & _
                   "|USA|?|?|" & vbLf & vbLf & _
                   "Assistant:" & vbLf & _
                   "|Country|Capital|Currency|" & _
                   "|Italy|Rome|EUR|" & vbLf & _
                   "|UAE|Abu Dhabi|AED|" & vbLf & _
                   "|USA|Washington, D.C.|USD|")

        guidance = guidance & _
                  ("Example 2" & vbLf & _
                   "---------" & vbLf & vbLf & _
                   "User:" & vbLf & _
                   "|Full name|Gender|" & vbLf & _
                   "|Emily Johnson|Female|" & vbLf & _
                   "|David Martinez|Male|" & vbLf & vbLf & _
                   "|Aisha Patel|?|" & vbLf & _
                   "|Thomas Brown|?|" & vbLf & _
                   "|Yuki Tanaka|?|" & vbLf & vbLf & _
                   "Assistant:" & vbLf & _
                   "|Full name|Gender|" & vbLf & _
                   "|Emily Johnson|Female|" & vbLf & _
                   "|David Martinez|Male|" & vbLf & vbLf & _
                   "|Aisha Patel|Female|" & vbLf & _
                   "|Thomas Brown|Male|" & vbLf & _
                   "|Yuki Tanaka|Female|")

        guidance = guidance & vbLf & vbLf & _
                   ("Due to the limitation of the macro, do not output Markdown header delimiter (e.g., |---|---|).")

        command = RangeToText(input_data)

    Else
        Err.Raise vbObjectError + 1001, , "Invalid mode: " & CStr(mode) & ". Should be 1 or 2."

    End If

    cached = False

    If AI_Cache.Exists(command) Then
        cached = True
    End If

    If Not cached Then
        result = TransformText(source:="", _
                               command:=command, _
                               preamble:=guidance, _
                               temperature:=0, _
                               correct_punctuation:=False, _
                               extract_tripple_quotes:=False)

        result = Trim_(result)
        result = Replace(result, vbCr, "")

        AI_Cache.Add command, result

    Else
        result = AI_Cache(command)

    End If

    result_rows = Split(result, vbLf)

    cclb = 0
    ccub = 0

    For i = LBound(result_rows) To UBound(result_rows)
        s = result_rows(i)

        If Left(s, 1) = "|" Then
            s = Right(s, Len(s) - 1)
        End If

        If Right(s, 1) = "|" Then
            s = Left(s, Len(s) - 1)
        End If

        s = Trim_(s)

        a = Split(s, "|")

        If LBound(a) < cclb Then
            cclb = LBound(a)
        End If

        If UBound(a) > ccub Then
            ccub = UBound(a)
        End If
    Next i

    ReDim out_arr(LBound(result_rows) To UBound(result_rows), _
                  cclb To ccub)

    For i = LBound(result_rows) To UBound(result_rows)
        s = result_rows(i)

        If Left(s, 1) = "|" Then
            s = Right(s, Len(s) - 1)
        End If

        If Right(s, 1) = "|" Then
            s = Left(s, Len(s) - 1)
        End If

        s = Trim_(s)

        a = Split(s, "|")

        For j = LBound(a) To UBound(a)
            out_arr(i, j) = Trim_(a(j))
        Next j
    Next i

    AI = out_arr
End Function

Sub AIClearCache()
    If AI_Cache Is Nothing Then
        Set AI_Cache = New Dictionary
    End If

    AI_Cache.RemoveAll
End Sub

#End If

' ==============================================================================
' Forms
' ==============================================================================

Sub OpenSettings()
    With AIToolsSettings
        .OpenAIAPIKeyTextBox.text = _
            GetSetting("AI tools", "API Keys", "openai", "")
        .GoogleAIAPIKeyTextBox.text = _
            GetSetting("AI tools", "API Keys", "google", "")
        .AnthropicAPIKeyTextBox.text = _
            GetSetting("AI tools", "API Keys", "anthropic", "")
        .OpenRouterAPIKeyTextBox.text = _
            GetSetting("AI tools", "API Keys", "openrouter", "")
        .MistralAIAPIKeyTextBox.text = _
            GetSetting("AI tools", "API Keys", "mistralai", "")
        .CohereAPIKeyTextBox.text = _
            GetSetting("AI tools", "API Keys", "cohere", "")
        .Show
    End With
End Sub

' ==============================================================================
' CustomUI Callbacks
' ==============================================================================

Sub AIDefaultModelDropdownCallback(control As IRibbonControl, _
                                   id As String, _
                                   index As Integer)
    SaveSetting "AI tools", "Settings", "Model", id
End Sub

Sub AIDefaultModelGetSelectedItemID(control As IRibbonControl, _
                                    ByRef returnedVal As Variant)
    returnedVal = GetDefaultModel()
End Sub

Sub CorrectToStandardEnglishButtonCallback(control As IRibbonControl)
    CorrectToStandardEnglish
End Sub

Sub CorrectToStandardEnglishBusinessButtonCallback(control As IRibbonControl)
    CorrectToStandardEnglishBusiness
End Sub

Sub ParaphraseShortenButtonCallback(control As IRibbonControl)
    ParaphraseShorten
End Sub

#If IsPowerPoint Then

Sub RephraseConsultingZeroShotButtonCallback(control As IRibbonControl)
    RephraseConsultingZeroShot
End Sub

Sub RephraseConsultingMultiShotButtonCallback(control As IRibbonControl)
    RephraseConsultingMultiShot
End Sub

#End If

#If IsPowerPoint Or IsWord Then

Sub RephrasePoliteConciseButtonCallback(control As IRibbonControl)
    RephrasePoliteConcise
End Sub

Sub RephrasePoliteExtraButtonCallback(control As IRibbonControl)
    RephrasePoliteExtra
End Sub

#End If

Sub SettingsButtonCallback(control As IRibbonControl)
    OpenSettings
End Sub

Sub EnforceRnQComplianceCheckboxOnActionCallback(control As IRibbonControl, _
                                                 pressed As Boolean)
    ' TODO
End Sub

'Callback for EnforceRnQComplianceCheckbox getPressed
Sub EnforceRnQComplianceCheckboxGetPressedCallback(control As IRibbonControl, _
                                                   ByRef returnedVal)
    returnedVal = False
    ' TODO
End Sub

#If IsPowerPoint Then

Sub RunAIButtonCallback(control As IRibbonControl)
    RunAI
End Sub

Sub RephraseTitleVariantsButtonCallback(control As IRibbonControl)
    RephraseTitleVariants
End Sub

#End If

' ==============================================================================
' Developer functions
' ==============================================================================

#If DeveloperMode Then
#If IsPowerPoint Then

Sub TagSlideAsAIPlayground()
    ActiveWindow.View.Slide.Tags.Add "RUNAIPLAYGROUND", "true"
End Sub

Sub PrintTags()
    Dim i As Integer

    With ActiveWindow.View.Slide.Tags
        For i = 1 To .Count
            Debug.Print "Name = '" & CStr(.Name(i)) & "', Value = '" & CStr(.Value(i)) & "'"
        Next i
    End With
End Sub

Sub RemoveAllTags()
    Dim i As Integer

    With ActiveWindow.View.Slide.Tags
        For i = .Count To 1 Step -1
            .Delete .Name(i)
        Next i
    End With
End Sub

#End If
#End If

' ==============================================================================
' Service functions
' ==============================================================================

#If IsPowerPoint Then

Private Sub TransformSelection(command As String, _
                               Optional preamble As String, _
                               Optional temperature As Double = 0, _
                               Optional correct_punctuation As Boolean = True, _
                               Optional extract_tripple_quotes As Boolean = True, _
                               Optional anonymize_client As String, _
                               Optional stop_word As String, _
                               Optional model As String, _
                               Optional remove_prefix As String)
    Dim tr As TextRange
    Dim source As String
    Dim result As String

    Dim error_text As String
    error_text = ("Invalid selection. Please select a text fragment in a " & _
                  "text field, or select one shape with text.")

    With ActiveWindow.Selection
        If .Type = ppSelectionText Then
            Set tr = .TextRange
        ElseIf .Type = ppSelectionShapes Then
            If .ShapeRange.Count > 1 Then
                MsgBox error_text, vbCritical, "AI Tools: Error"
                Exit Sub
            ElseIf .ShapeRange.Count = 0 Then
                MsgBox error_text, vbCritical, "AI Tools: Error"
                Exit Sub
            End If
            With .ShapeRange(1)
                If Not .HasTextFrame Then
                    MsgBox error_text, vbCritical, "AI Tools: Error"
                    Exit Sub
                End If
                If Not .TextFrame.HasText Then
                    MsgBox error_text, vbCritical, "AI Tools: Error"
                    Exit Sub
                End If
                Set tr = .TextFrame.TextRange
            End With
        Else
            MsgBox error_text, vbCritical, "AI Tools: Error"
            Exit Sub
        End If
    End With

    With tr
        source = .text

        result = TransformText(source:=source, _
                               command:=command, _
                               preamble:=preamble, _
                               temperature:=temperature, _
                               correct_punctuation:=correct_punctuation, _
                               extract_tripple_quotes:=extract_tripple_quotes, _
                               anonymize_client:=anonymize_client, _
                               stop_word:=stop_word, _
                               model:=model, _
                               remove_prefix:=remove_prefix)

        If Not IsEmpty(result) And result <> "" Then
            .text = result
        End If
    End With
End Sub

#End If

#If IsWord Then

Private Sub TransformSelection(command As String, _
                               Optional preamble As String, _
                               Optional temperature As Double = 0, _
                               Optional correct_punctuation As Boolean = True, _
                               Optional extract_tripple_quotes As Boolean = True, _
                               Optional anonymize_client As String, _
                               Optional stop_word As String, _
                               Optional model As String, _
                               Optional remove_prefix As String)
    Dim source As String
    Dim result As String

    Dim deselect_chars As Integer
    Dim i As Integer
    Dim c As String

    Dim error_text As String
    error_text = "Invalid selection. Please select a text fragment."

    With ActiveWindow.Selection
        If .Type <> wdSelectionNormal Then
            MsgBox error_text, vbCritical, "AI Tools: Error"
            Exit Sub
        End If

        source = .text

        deselect_chars = 0

        For i = Len(source) To 1 Step -1
            c = Mid(source, i, 1)

            If c = Chr(10) Or c = Chr(13) Or c = " " Then
                deselect_chars = deselect_chars + 1
            Else
                Exit For
            End If
        Next i

        result = TransformText(source:=source, _
                               command:=command, _
                               preamble:=preamble, _
                               temperature:=temperature, _
                               correct_punctuation:=correct_punctuation, _
                               extract_tripple_quotes:=extract_tripple_quotes, _
                               anonymize_client:=anonymize_client, _
                               stop_word:=stop_word, _
                               model:=model, _
                               remove_prefix:=remove_prefix)

        If Not IsEmpty(result) And Trim_(result) <> Trim_(source) Then
            result = RTrim_(result)

            If result <> "" Then
                .MoveEnd Unit:=wdCharacter, Count:=-deselect_chars
                .text = result
                .MoveEnd Unit:=wdCharacter, Count:=deselect_chars
            End If
        End If
    End With
End Sub

#End If

#If IsExcel Then

Private Sub TransformSelection(command As String, _
                               Optional preamble As String, _
                               Optional temperature As Double = 0, _
                               Optional correct_punctuation As Boolean = True, _
                               Optional extract_tripple_quotes As Boolean = True, _
                               Optional anonymize_client As String, _
                               Optional stop_word As String, _
                               Optional model As String, _
                               Optional remove_prefix As String)
    Dim source As String
    Dim result As String
    Dim addr As String

    Dim error_text As String
    error_text = "Invalid selection. Please select a single cell."

    With ActiveWindow.Selection
        If .Cells.Count = 0 Then
            MsgBox error_text, vbCritical, "AI Tools: Error"
            Exit Sub
        ElseIf .Cells.Count > 1 Then
            If .MergeCells Then
                If .Cells(1).MergeArea.Count <> .Cells.Count Then
                    MsgBox error_text, vbCritical, "AI Tools: Error"
                    Exit Sub
                End If
            Else
                MsgBox error_text, vbCritical, "AI Tools: Error"
                Exit Sub
            End If
        End If

        addr = .Address(External:=True)

        source = .Cells(1).text

        result = TransformText(source:=source, _
                               command:=command, _
                               preamble:=preamble, _
                               temperature:=temperature, _
                               correct_punctuation:=correct_punctuation, _
                               extract_tripple_quotes:=extract_tripple_quotes, _
                               anonymize_client:=anonymize_client, _
                               stop_word:=stop_word, _
                               model:=model, _
                               remove_prefix:=remove_prefix)
    End With

    If IsEmpty(result) Or result = "" Then
        Exit Sub
    End If

    If source = result Then
        Exit Sub
    End If

    With AIToolsExcel
        .Tag = addr
        .TextBoxInput.text = source
        .TextBoxOutput.text = result
        .Show
    End With
End Sub

#End If

' ##################################################################################################

Private Function TransformText(ByVal source As String, _
                               ByVal command As String, _
                               Optional preamble As String, _
                               Optional temperature As Double = 0, _
                               Optional correct_punctuation As Boolean = True, _
                               Optional extract_tripple_quotes As Boolean = True, _
                               Optional anonymize_client As String, _
                               Optional stop_word As String, _
                               Optional model As String, _
                               Optional remove_prefix As String) _
                               As String
    Dim result As String
    Dim c As String
    Dim full_stop_added As Boolean
    Dim a, b As Long

    source = RTrim_(source)

    If correct_punctuation Then
        c = Right(source, 1)
        full_stop_added = False

        If c <> "." And _
           c <> ":" And _
           c <> ";" And _
           c <> "!" And _
           c <> "?" And _
           c <> "," And _
           c <> "%" And _
           1 = 1 Then
            source = source & "."
            full_stop_added = True
        End If
    End If

    If Not IsEmpty(anonymize_client) And anonymize_client <> "" Then
        command = Replace(command, anonymize_client, "[Client]")
        source = Replace(source, anonymize_client, "[Client]")
    End If

    result = LLMTextCommand(command:=command, _
                            source:=source, _
                            preamble:=preamble, _
                            model:=model, _
                            temperature:=temperature, _
                            stop_word:=stop_word)

    If Not IsEmpty(result) Then
        result = Trim_(result)
    End If

    If Not IsEmpty(result) And result <> "" Then
        If extract_tripple_quotes Then
            a = InStr(result, """""""")
            b = InStrRev(result, """""""")

            If a > 0 And b > 0 And a <> b Then
                result = Mid(result, a + 3, Len(result) - a - 2)
                result = Left(result, InStr(result, """""""") - 1)
                result = Trim_(result)
            End If
        End If

        If correct_punctuation And full_stop_added Then
            If Right(result, 1) = "." Then
                result = Left(result, Len(result) - 1)
            End If
        Else
            If Right(result, 1) = "." And Right(source, 1) <> "." Then
                result = Left(result, Len(result) - 1)
            End If
        End If
    End If

    If Not IsEmpty(anonymize_client) And anonymize_client <> "" Then
        result = Replace(result, "[Client]", anonymize_client)
    End If

    If Not IsEmpty(remove_prefix) And remove_prefix <> "" Then
        If Len(result) > Len(remove_prefix) And Left(result, Len(remove_prefix)) = remove_prefix Then
            result = Right(result, Len(result) - Len(remove_prefix))
        End If
    End If

    If source = Trim(source) And result <> Trim(result) Then
        result = Trim(result)
    End If

    If StartsWith(result, """""""") And EndsWith(result, """""""") Then
        result = RemoveSuffix(RemovePrefix(result, """"""""), """""""")
    End If

    ' Chr(34) -> "
    If Len(result) >= 2 And Right(result, 1) = Chr(34) And Left(result, 1) = Chr(34) Then
        If Len(source) >= 2 And Not (Right(source, 1) = Chr(34) And Left(source, 1) = Chr(34)) Then
            result = Mid(result, 2, Len(result) - 2)
        End If
    End If

    TransformText = result
End Function

Private Function LLMTextCommand(ByVal command As String, _
                                ByVal source As String, _
                                Optional preamble As String, _
                                Optional temperature As Double = 0, _
                                Optional model As String, _
                                Optional placeholder As String = "{{input}}", _
                                Optional stop_word As String, _
                                Optional normalize_newline As Boolean = True) _
                                As String
    Dim result As String
    Dim prompt As String

    command = Trim_(command)
    source = Trim_(source)

    If command <> "" And source <> "" Then
        If Not IsEmpty(placeholder) And placeholder <> "" And InStr(command, placeholder) > 0 Then
            prompt = Replace(command, placeholder, source)
        Else
            prompt = ("You are given instructions and input text delimited by triple quotes (""""""). " & _
                      "Apply the instructions to the input text and write the result." & vbLf & vbLf & vbLf & _
                      "# Instructions" & vbLf & _
                      command & vbLf & vbLf & vbLf & _
                      "# Input text" & vbLf & _
                      """""""" & source & """""""")
        End If
    End If

    prompt = Trim_(prompt)

    If prompt = "" Then
        LLMTextCommand = ""
        Exit Function
    End If

    result = LLMChat(prompt:=prompt, _
                     preamble:=preamble, _
                     temperature:=temperature, _
                     stop_word:=stop_word, _
                     model:=model)

    result = Trim_(result)

    If normalize_newline Then
        result = Replace(result, vbNewLine, vbLf)
        result = Replace(result, Chr(13), "")
    End If

    LLMTextCommand = result
End Function

Private Function GetAPIKey(provider As String, _
                           Optional safe As Boolean = False) As String
    GetAPIKey = GetSetting("AI tools", "API Keys", provider, "")

    If safe Then
        Exit Function
    End If

    If IsEmpty(GetAPIKey) Or GetAPIKey = "" Then
        Err.Raise vbObjectError + 1001, , _
            ("API Key for " & provider & " is not set. " & _
             "Press 'End', go to AI Tools Settings, set " & _
             "the API key for " & provider & ", and try again.")
    End If
End Function

Private Function LLMChat(prompt As String, _
                         Optional ByVal preamble As String, _
                         Optional temperature As Double = 0, _
                         Optional stop_word As String, _
                         Optional ByVal model As String) _
                         As String
    Dim provider As String
    Dim base_url As String

    If IsEmpty(model) Or model = "" Then
        model = GetDefaultModel()
    End If

    provider = GetProvider(model)

    If IsEmpty(preamble) Or preamble = "" Then
        preamble = GetDefaultPreamble()
    End If

    If provider = "openai" _
            Or provider = "openrouter" _
            Or provider = "mistralai" _
            Then
        LLMChat = LLMChatOpenAI(provider:=provider, _
                                model:=model, _
                                prompt:=prompt, _
                                preamble:=preamble, _
                                temperature:=temperature, _
                                stop_word:=stop_word)

    ElseIf provider = "anthropic" Then
        LLMChat = LLMChatAnthropic(model:=model, _
                                   prompt:=prompt, _
                                   preamble:=preamble, _
                                   temperature:=temperature, _
                                   stop_word:=stop_word)

    ElseIf provider = "google" Then
        LLMChat = LLMChatGoogleAI(model:=model, _
                                  prompt:=prompt, _
                                  preamble:=preamble, _
                                  temperature:=temperature, _
                                  stop_word:=stop_word)

    ElseIf provider = "cohere" Then
        LLMChat = LLMChatCohere(model:=model, _
                                prompt:=prompt, _
                                preamble:=preamble, _
                                temperature:=temperature, _
                                stop_word:=stop_word)

    Else
        Err.Raise vbObjectError + 1001, , "Wrong provider"
    End If
End Function

Private Function LLMChatOpenAI(provider As String, _
                               model As String, _
                               prompt As String, _
                               preamble As String, _
                               Optional temperature As Double = 0, _
                               Optional stop_word As String) _
                               As String
    Dim base_url As String

    base_url = GetBaseURL(provider)

#If IsWord Then
    Dim payload As Scripting.Dictionary
    Set payload = New Scripting.Dictionary
#Else
    Dim payload As Dictionary
    Set payload = New Dictionary
#End If

    payload.Add "model", GetModelName(model, provider)
    payload.Add "temperature", temperature
    payload.Add "max_tokens", 2000

    Dim messages As Collection
    Set messages = New Collection

#If IsWord Then
    Dim d As Scripting.Dictionary
    Set d = New Scripting.Dictionary
#Else
    Dim d As Dictionary
    Set d = New Dictionary
#End If

    d.Add "role", "system"
    d.Add "content", preamble

    messages.Add d

#If IsWord Then
    Set d = New Scripting.Dictionary
#Else
    Set d = New Dictionary
#End If

    d.Add "role", "user"
    d.Add "content", prompt

    messages.Add d

    payload.Add "messages", messages

    If Not IsEmpty(stop_word) And stop_word <> "" Then
        payload.Add "stop", stop_word
    End If

    Dim request As String
    request = ConvertToJson(payload)

#If DeveloperMode Then
    Debug.Print ">>>>>>>>>>>>"
    Debug.Print "Provider:", provider
    Debug.Print "Base URL:", base_url
    Debug.Print "Request:", request
#End If

    Dim http As WinHttp.WinHttpRequest
    Set http = New WinHttpRequest

    Dim timeout As Long
    timeout = 60000  ' ms
    http.SetTimeouts timeout, timeout, timeout, timeout

    http.Open "POST", base_url
    http.SetRequestHeader "Content-Type", "application/json"
    http.SetRequestHeader "Authorization", "Bearer " & GetAPIKey(provider)

    On Error GoTo ErrorHandler
    http.Send request
    On Error GoTo 0
    GoTo NoErrors

ErrorHandler:
    On Error GoTo 0
#If DeveloperMode Then
    Debug.Print "HTTP Error:", Err.Description
    Debug.Print "<<<<<<<<<<<<"
#End If
    Err.Raise vbObjectError + 1001, , _
        "HTTP Error: " & Err.Description & ". Press 'End' and try again."

NoErrors:
    Dim response_str As String
    response_str = http.ResponseText
    response_str = DecodeText(response_str, "ISO-8859-1", "UTF-8")

#If DeveloperMode Then
    Debug.Print "Response:", response_str
    Debug.Print "<<<<<<<<<<<<"
#End If

    Dim response_json As Object
    Set response_json = ParseJson(response_str)

    Dim i As Integer

    For i = 0 To response_json.Count - 1
        If response_json.keys()(i) = "error" Then
            Err.Raise vbObjectError + 1001, , _
                "LLM service provider returned the error: " & _
                response_json("error")("message") & ". Press 'End' and try again."
        End If
    Next i

    LLMChatOpenAI = response_json("choices")(1)("message")("content")
End Function

Private Function LLMChatAnthropic(model As String, _
                                  prompt As String, _
                                  preamble As String, _
                                  Optional temperature As Double = 0, _
                                  Optional stop_word As String) _
                                  As String
    Dim base_url As String
    base_url = "https://api.anthropic.com/v1/messages"

#If IsWord Then
    Dim payload As Scripting.Dictionary
    Set payload = New Scripting.Dictionary
#Else
    Dim payload As Dictionary
    Set payload = New Dictionary
#End If

    payload.Add "model", GetModelName(model, "anthropic")
    payload.Add "temperature", temperature
    payload.Add "max_tokens", 4096  ' required

    payload.Add "system", preamble

    Dim messages As Collection
    Set messages = New Collection

#If IsWord Then
    Dim d As Scripting.Dictionary
    Set d = New Scripting.Dictionary
#Else
    Dim d As Dictionary
    Set d = New Dictionary
#End If

    d.Add "role", "user"
    d.Add "content", prompt

    messages.Add d

    payload.Add "messages", messages

    Dim stop_sequences As Collection

    If Not IsEmpty(stop_word) And stop_word <> "" Then
        Set stop_sequences = New Collection
        stop_sequences.Add stop_word
        payload.Add "stop_sequences", stop_sequences
    End If

    Dim request As String
    request = ConvertToJson(payload)

#If DeveloperMode Then
    Debug.Print ">>>>>>>>>>>>"
    Debug.Print "Provider:", "antchropic"
    Debug.Print "Base URL:", base_url
    Debug.Print "Request:", request
#End If

    Dim http As WinHttp.WinHttpRequest
    Set http = New WinHttpRequest

    Dim timeout As Long
    timeout = 60000  ' ms
    http.SetTimeouts timeout, timeout, timeout, timeout

    http.Open "POST", base_url
    http.SetRequestHeader "Content-Type", "application/json"
    http.SetRequestHeader "X-API-Key", GetAPIKey("anthropic")
    http.SetRequestHeader "Anthropic-Version", "2023-06-01"

    On Error GoTo ErrorHandler
    http.Send request
    On Error GoTo 0
    GoTo NoErrors

ErrorHandler:
    On Error GoTo 0
#If DeveloperMode Then
    Debug.Print "HTTP Error:", Err.Description
    Debug.Print "<<<<<<<<<<<<"
#End If
    Err.Raise vbObjectError + 1001, , _
        "HTTP Error: " & Err.Description & ". Press 'End' and try again."

NoErrors:
    Dim response_str As String
    response_str = http.ResponseText
    response_str = DecodeText(response_str, "ISO-8859-1", "UTF-8")

#If DeveloperMode Then
    Debug.Print "Response:", response_str
    Debug.Print "<<<<<<<<<<<<"
#End If

    Dim response_json As Object
    Set response_json = ParseJson(response_str)

    Dim i As Integer

    For i = 0 To response_json.Count - 1
        If response_json.keys()(i) = "error" Then
            Err.Raise vbObjectError + 1001, , _
                "LLM service provider returned the error: " & _
                response_json("error")("message") & ". Press 'End' and try again."
        End If
    Next i

    LLMChatAnthropic = response_json("content")(1)("text")
End Function

Private Function LLMChatGoogleAI(model As String, _
                                 prompt As String, _
                                 preamble As String, _
                                 Optional temperature As Double = 0, _
                                 Optional stop_word As String) _
                                 As String
    Dim url As String

    url = "https://generativelanguage.googleapis.com/v1beta/models/" _
        & GetModelName(model, "google") & ":generateContent"

#If IsWord Then
    Dim payload As Scripting.Dictionary
    Set payload = New Scripting.Dictionary
#Else
    Dim payload As Dictionary
    Set payload = New Dictionary
#End If

#If IsWord Then
    Dim gc As Scripting.Dictionary
    Set gc = New Scripting.Dictionary
#Else
    Dim gc As Dictionary
    Set gc = New Dictionary
#End If

    If Not IsEmpty(stop_word) And stop_word <> "" Then
        Dim ss As Collection
        ss.Add stop_word
        gc.Add "stopSequences", ss
    End If

    gc.Add "temperature", temperature
    gc.Add "maxOutputTokens", 2000

    payload.Add "generationConfig", gc

    Dim ct As Collection
    Set ct = New Collection

#If IsWord Then
    Dim p As Scripting.Dictionary
    Set p = New Scripting.Dictionary
#Else
    Dim p As Dictionary
    Set p = New Dictionary
#End If

    p.Add "role", "user"

#If IsWord Then
    Dim t As Scripting.Dictionary
    Set t = New Scripting.Dictionary
#Else
    Dim t As Dictionary
    Set t = New Dictionary
#End If

    t.Add "text", preamble

    Dim ts As Collection
    Set ts = New Collection

    ts.Add t

    p.Add "parts", ts

    ct.Add p

#If IsWord Then
    Set p = New Scripting.Dictionary
#Else
    Set p = New Dictionary
#End If

    p.Add "role", "model"

#If IsWord Then
    Set t = New Scripting.Dictionary
#Else
    Set t = New Dictionary
#End If

    t.Add "text", "Sure, I can help you with that."

    Set ts = New Collection

    ts.Add t

    p.Add "parts", ts

    ct.Add p

#If IsWord Then
    Set p = New Scripting.Dictionary
#Else
    Set p = New Dictionary
#End If

    p.Add "role", "user"

#If IsWord Then
    Set t = New Scripting.Dictionary
#Else
    Set t = New Dictionary
#End If

    t.Add "text", prompt

    p.Add "parts", t

    ct.Add p

    payload.Add "contents", ct

    Dim request As String
    request = ConvertToJson(payload)

#If DeveloperMode Then
    Debug.Print ">>>>>>>>>>>>"
    Debug.Print "Provider:", "google"
    Debug.Print "URL:", url
    Debug.Print "Request:", request
#End If

    Dim http As WinHttp.WinHttpRequest
    Set http = New WinHttpRequest

    Dim timeout As Long
    timeout = 60000  ' ms
    http.SetTimeouts timeout, timeout, timeout, timeout

    http.Open "POST", url & "?key=" & GetAPIKey("google")
    http.SetRequestHeader "Content-Type", "application/json"

    On Error GoTo ErrorHandler
    http.Send request
    On Error GoTo 0
    GoTo NoErrors

ErrorHandler:
    On Error GoTo 0
#If DeveloperMode Then
    Debug.Print "HTTP Error:", Err.Description
    Debug.Print "<<<<<<<<<<<<"
#End If
    Err.Raise vbObjectError + 1001, , _
        "HTTP Error: " & Err.Description & ". Press 'End' and try again."

NoErrors:
    Dim response_str As String
    response_str = http.ResponseText
    response_str = DecodeText(response_str, "ISO-8859-1", "UTF-8")

#If DeveloperMode Then
    Debug.Print "Response:", response_str
    Debug.Print "<<<<<<<<<<<<"
#End If

    Dim response_json As Object
    Set response_json = ParseJson(response_str)

    Dim i As Integer

    For i = 0 To response_json.Count - 1
        If response_json.keys()(i) = "error" Then
            Err.Raise vbObjectError + 1001, , _
                "LLM service provider returned the error: " & _
                response_json("error")("message") & " Press 'End' and try again."
        End If
    Next i

    LLMChatGoogleAI = response_json("candidates")(1)("content")("parts")(1)("text")
End Function

Private Function LLMChatCohere(model As String, _
                               prompt As String, _
                               preamble As String, _
                               Optional temperature As Double = 0, _
                               Optional stop_word As String) _
                               As String
    Dim base_url As String
    base_url = "https://api.cohere.ai/v1/chat"

#If IsWord Then
    Dim payload As Scripting.Dictionary
    Set payload = New Scripting.Dictionary
#Else
    Dim payload As Dictionary
    Set payload = New Dictionary
#End If

    payload.Add "model", GetModelName(model, "cohere")
    payload.Add "temperature", temperature
    payload.Add "preamble", preamble
    payload.Add "message", prompt

    Dim stop_sequences As Collection

    If Not IsEmpty(stop_word) And stop_word <> "" Then
        Set stop_sequences = New Collection
        stop_sequences.Add stop_word
        payload.Add "stop_sequences", stop_sequences
    End If

    Dim request As String
    request = ConvertToJson(payload)

#If DeveloperMode Then
    Debug.Print ">>>>>>>>>>>>"
    Debug.Print "Provider:", "cohere"
    Debug.Print "Base URL:", base_url
    Debug.Print "Request:", request
#End If

    Dim http As WinHttp.WinHttpRequest
    Set http = New WinHttpRequest

    Dim timeout As Long
    timeout = 60000  ' ms
    http.SetTimeouts timeout, timeout, timeout, timeout

    http.Open "POST", base_url
    http.SetRequestHeader "Accept", "application/json"
    http.SetRequestHeader "Content-Type", "application/json"
    http.SetRequestHeader "Authorization", "Bearer " & GetAPIKey("cohere")

    On Error GoTo ErrorHandler
    http.Send request
    On Error GoTo 0
    GoTo NoErrors

ErrorHandler:
    On Error GoTo 0
#If DeveloperMode Then
    Debug.Print "HTTP Error:", Err.Description
    Debug.Print "<<<<<<<<<<<<"
#End If
    Err.Raise vbObjectError + 1001, , _
        "HTTP Error: " & Err.Description & ". Press 'End' and try again."

NoErrors:
    Dim response_str As String
    response_str = http.ResponseText
    response_str = DecodeText(response_str, "ISO-8859-1", "UTF-8")

#If DeveloperMode Then
    Debug.Print "Response:", response_str
    Debug.Print "<<<<<<<<<<<<"
#End If

    Dim response_json As Object
    Set response_json = ParseJson(response_str)

    If http.status >= 300 Then
        Dim i As Integer

        For i = 0 To response_json.Count - 1
            If response_json.keys()(i) = "message" Then
                Err.Raise vbObjectError + 1001, , _
                    "LLM service provider returned the error: " & _
                    response_json("message") & ". Press 'End' and try again."
            End If
            If response_json.keys()(i) = "data" Then
                Err.Raise vbObjectError + 1001, , _
                    "LLM service provider returned the error: " & _
                    response_json("data") & ". Press 'End' and try again."
            End If
        Next i
    End If

    LLMChatCohere = response_json("text")
End Function

' #############################################################################

Private Function DecodeText(ByVal text As String, _
                            ByVal fromCharset As String, _
                            ByVal toCharset As String) _
                            As String
    With New ADODB.Stream
        .Type = 2
        .mode = 3
        .Charset = fromCharset
        .Open
        .WriteText text
        .Position = 0
        .Charset = toCharset
        DecodeText = .ReadText(-1)
        .Close
    End With
End Function

Private Function LTrim_(s As String) As String
    s = LTrim(s)

    Dim c As String

    Do While Len(s) > 0
        c = Left(s, 1)
        
        If c = Chr(10) Or c = Chr(13) Or c = " " Or c = Chr(9) Then
            s = Right(s, Len(s) - 1)
        Else
            Exit Do
        End If
    Loop
    
    LTrim_ = s
End Function

Private Function RTrim_(s As String) As String
    s = RTrim(s)

    Dim c As String

    Do While Len(s) > 0
        c = Right(s, 1)
        
        If c = Chr(10) Or c = Chr(13) Or c = " " Or c = Chr(9) Then
            s = Left(s, Len(s) - 1)
        Else
            Exit Do
        End If
    Loop
    
    RTrim_ = s
End Function

Private Function Trim_(s As String) As String
    Trim_ = LTrim_(RTrim_(s))
End Function

Private Function StartsWith(str As String, prefix As String) As Boolean
    If Len(prefix) > Len(str) Then
        StartsWith = False
        Exit Function
    End If
    
    StartsWith = Left(str, Len(prefix)) = prefix
End Function

Private Function EndsWith(str As String, suffix As String) As Boolean
    If Len(suffix) > Len(str) Then
        EndsWith = False
        Exit Function
    End If
    
    EndsWith = Right(str, Len(suffix)) = suffix
End Function

Private Function RemovePrefix(str As String, prefix As String) As String
    If Not StartsWith(str, prefix) Then
        RemovePrefix = str
        Exit Function
    End If
    
    RemovePrefix = Right(str, Len(str) - Len(prefix))
End Function

Private Function RemoveSuffix(str As String, suffix As String) As String
    If Not EndsWith(str, suffix) Then
        RemoveSuffix = str
        Exit Function
    End If
    
    RemoveSuffix = Left(str, Len(str) - Len(suffix))
End Function

#If IsPowerPoint Then

Private Function GetText(s As Shape) As String
    GetText = s.TextFrame.TextRange.text
End Function

#End If

#If IsExcel Then

Private Function RangeToText(r As Range) As String
    Dim a As Variant
    Dim rc As Long, cc As Long
    Dim i As Long, j As Long
    Dim result As String
    Dim s As String
    Dim x As String

    a = r.Value
    
    If VarType(a) < vbArray Then
        RangeToText = "|" & Sanitize(a) & "|" & vbLf
        Exit Function
    End If

    rc = UBound(a, 1)
    cc = UBound(a, 2)
    
    result = ""
    
    For i = LBound(a, 1) To UBound(a, 1)
        s = ""
        
        For j = LBound(a, 2) To UBound(a, 2)
            s = s & "|" & Sanitize(a(i, j))
        Next j
        
        result = result & s & "|" & vbLf
    Next i
    
    RangeToText = result
End Function

Private Function Sanitize(x As Variant) As String
    Dim s As String
    s = CStr(x)
    s = Trim_(s)
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, " ")
    s = Replace(s, "|", "\|")
    Sanitize = s
End Function

#End If

#If IsPowerPoint Then

Private Sub UnblockingWait(seconds As Double)
    Dim endtime As Double
    endtime = DateTime.Timer + seconds
    
    Do
        WaitMessage
        DoEvents
    Loop While DateTime.Timer < endtime
End Sub

#End If