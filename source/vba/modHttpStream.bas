Attribute VB_Name = "modHttpStream"
Option Explicit

'=======================================================================
' modHttpStream - streaming chat-completion client for PatentTools
'
' This module talks to an OpenAI-compatible endpoint over WinINet, which is
' the only in-process API in VBA that allows reading a response
' incrementally. Server-sent events are decoded chunk by chunk, so the Word
' status bar can show:
'   - prompt (prefill) progress, when the server supports llama.cpp's
'     "return_progress" and emits a "prompt_progress" object
'   - generated-token progress and elapsed time afterwards
'
' The only public entry point is StreamChatCompletion.
'=======================================================================

#If VBA7 Then
    Private Declare PtrSafe Function InternetOpenA Lib "wininet.dll" ( _
        ByVal lpszAgent As String, ByVal dwAccessType As Long, _
        ByVal lpszProxy As String, ByVal lpszProxyBypass As String, _
        ByVal dwFlags As Long) As LongPtr

    Private Declare PtrSafe Function InternetConnectA Lib "wininet.dll" ( _
        ByVal hInternet As LongPtr, ByVal lpszServerName As String, _
        ByVal nServerPort As Long, ByVal lpszUserName As String, _
        ByVal lpszPassword As String, ByVal dwService As Long, _
        ByVal dwFlags As Long, ByVal dwContext As LongPtr) As LongPtr

    Private Declare PtrSafe Function HttpOpenRequestA Lib "wininet.dll" ( _
        ByVal hConnect As LongPtr, ByVal lpszVerb As String, _
        ByVal lpszObjectName As String, ByVal lpszVersion As String, _
        ByVal lpszReferrer As String, ByVal lplpszAcceptTypes As LongPtr, _
        ByVal dwFlags As Long, ByVal dwContext As LongPtr) As LongPtr

    Private Declare PtrSafe Function HttpAddRequestHeadersA Lib "wininet.dll" ( _
        ByVal hRequest As LongPtr, ByVal lpszHeaders As String, _
        ByVal dwHeadersLength As Long, ByVal dwModifiers As Long) As Long

    Private Declare PtrSafe Function HttpSendRequestA Lib "wininet.dll" ( _
        ByVal hRequest As LongPtr, ByVal lpszHeaders As String, _
        ByVal dwHeadersLength As Long, ByRef lpOptional As Any, _
        ByVal dwOptionalLength As Long) As Long

    Private Declare PtrSafe Function HttpQueryInfoA Lib "wininet.dll" ( _
        ByVal hRequest As LongPtr, ByVal dwInfoLevel As Long, _
        ByRef lpvBuffer As Any, ByRef lpdwBufferLength As Long, _
        ByRef lpdwIndex As Long) As Long

    Private Declare PtrSafe Function InternetQueryDataAvailable Lib "wininet.dll" ( _
        ByVal hFile As LongPtr, ByRef lpdwNumberOfBytesAvailable As Long, _
        ByVal dwFlags As Long, ByVal dwContext As LongPtr) As Long

    Private Declare PtrSafe Function InternetReadFile Lib "wininet.dll" ( _
        ByVal hFile As LongPtr, ByRef lpBuffer As Any, _
        ByVal dwNumberOfBytesToRead As Long, _
        ByRef lpdwNumberOfBytesRead As Long) As Long

    Private Declare PtrSafe Function InternetSetOptionA Lib "wininet.dll" ( _
        ByVal hInternet As LongPtr, ByVal dwOption As Long, _
        ByRef lpBuffer As Any, ByVal dwBufferLength As Long) As Long

    Private Declare PtrSafe Function InternetCloseHandle Lib "wininet.dll" ( _
        ByVal hInet As LongPtr) As Long
#Else
    Private Declare Function InternetOpenA Lib "wininet.dll" ( _
        ByVal lpszAgent As String, ByVal dwAccessType As Long, _
        ByVal lpszProxy As String, ByVal lpszProxyBypass As String, _
        ByVal dwFlags As Long) As Long

    Private Declare Function InternetConnectA Lib "wininet.dll" ( _
        ByVal hInternet As Long, ByVal lpszServerName As String, _
        ByVal nServerPort As Long, ByVal lpszUserName As String, _
        ByVal lpszPassword As String, ByVal dwService As Long, _
        ByVal dwFlags As Long, ByVal dwContext As Long) As Long

    Private Declare Function HttpOpenRequestA Lib "wininet.dll" ( _
        ByVal hConnect As Long, ByVal lpszVerb As String, _
        ByVal lpszObjectName As String, ByVal lpszVersion As String, _
        ByVal lpszReferrer As String, ByVal lplpszAcceptTypes As Long, _
        ByVal dwFlags As Long, ByVal dwContext As Long) As Long

    Private Declare Function HttpAddRequestHeadersA Lib "wininet.dll" ( _
        ByVal hRequest As Long, ByVal lpszHeaders As String, _
        ByVal dwHeadersLength As Long, ByVal dwModifiers As Long) As Long

    Private Declare Function HttpSendRequestA Lib "wininet.dll" ( _
        ByVal hRequest As Long, ByVal lpszHeaders As String, _
        ByVal dwHeadersLength As Long, ByRef lpOptional As Any, _
        ByVal dwOptionalLength As Long) As Long

    Private Declare Function HttpQueryInfoA Lib "wininet.dll" ( _
        ByVal hRequest As Long, ByVal dwInfoLevel As Long, _
        ByRef lpvBuffer As Any, ByRef lpdwBufferLength As Long, _
        ByRef lpdwIndex As Long) As Long

    Private Declare Function InternetQueryDataAvailable Lib "wininet.dll" ( _
        ByVal hFile As Long, ByRef lpdwNumberOfBytesAvailable As Long, _
        ByVal dwFlags As Long, ByVal dwContext As Long) As Long

    Private Declare Function InternetReadFile Lib "wininet.dll" ( _
        ByVal hFile As Long, ByRef lpBuffer As Any, _
        ByVal dwNumberOfBytesToRead As Long, _
        ByRef lpdwNumberOfBytesRead As Long) As Long

    Private Declare Function InternetSetOptionA Lib "wininet.dll" ( _
        ByVal hInternet As Long, ByVal dwOption As Long, _
        ByRef lpBuffer As Any, ByVal dwBufferLength As Long) As Long

    Private Declare Function InternetCloseHandle Lib "wininet.dll" ( _
        ByVal hInet As Long) As Long
#End If

Private Const INTERNET_OPEN_TYPE_PRECONFIG As Long = 0
Private Const INTERNET_SERVICE_HTTP        As Long = 3

Private Const INTERNET_FLAG_RELOAD          As Long = &H80000000
Private Const INTERNET_FLAG_NO_CACHE_WRITE  As Long = &H4000000
Private Const INTERNET_FLAG_KEEP_CONNECTION As Long = &H400000
Private Const INTERNET_FLAG_SECURE          As Long = &H800000
Private Const INTERNET_FLAG_PRAGMA_NOCACHE  As Long = &H100

Private Const INTERNET_OPTION_CONNECT_TIMEOUT As Long = 2
Private Const INTERNET_OPTION_SEND_TIMEOUT    As Long = 5
Private Const INTERNET_OPTION_RECEIVE_TIMEOUT As Long = 6

Private Const HTTP_QUERY_STATUS_CODE   As Long = 19
Private Const HTTP_QUERY_FLAG_NUMBER   As Long = &H20000000

Private Const HTTP_ADDREQ_FLAG_ADD_IF_NEW As Long = &H10000000

Private Const READ_BUFFER_SIZE As Long = 16384

' Status bar is refreshed at most this often (seconds), so streaming stays cheap.
Private Const STATUS_MIN_INTERVAL As Double = 0.2

' WinINet blocks inside InternetQueryDataAvailable until data arrives or the
' receive timeout expires. The receive timeout is therefore set to a short slice
' and the timeout is treated as "still waiting", so the loop regains control
' several times per second and can repaint the status bar. The overall patience
' is governed by timeoutSec, measured from the last byte received.
Private Const RECEIVE_SLICE_MS As Long = 700
Private Const ERROR_INTERNET_TIMEOUT As Long = 12002

' Absolute ceiling, as a multiple of timeoutSec. The idle deadline alone is not
' enough: a server that keeps sending SSE pings but never produces a token would
' otherwise keep Word busy indefinitely.
Private Const MAX_TOTAL_FACTOR As Long = 30

' Registry keys used to remember the measured prompt-processing speed, so a
' percentage can be estimated for servers that report no progress themselves.
Private Const CALIB_APP     As String = "PatentTools"
Private Const CALIB_SECTION As String = "Performance"
Private Const CALIB_KEY     As String = "PrefillTokensPerSec"

' Rough characters-per-token ratio for mixed German/English legal prose.
Private Const CHARS_PER_TOKEN As Double = 3.7

' Spinner state. The Word status bar is plain text, so a rotating character is
' the only available "the add-in is alive" signal. It advances on every status
' refresh, which is why the refresh must happen even when there is nothing new
' to report.
Private m_spinIndex As Long

' Result of the one-time probe for the llama.cpp native API, cached per base URL
' for the lifetime of the Word session. Not persisted: the user may stop Ollama
' and start llama-server on the same port between sessions.
Private m_probeBase As String
Private m_probeDone As Boolean
Private m_probeNative As Boolean

' Everything the read loop accumulates about one response. Passed ByRef to the
' line handler and the progress display: nineteen positional arguments would be
' an invitation to silent argument-order bugs.
Private Type HS_State
    ' --- transport / parse results ---
    sawSseData      As Boolean   ' anything stream-shaped was seen
    rawBody         As String    ' raw text, for diagnostics and the fallback
    contentSoFar    As String    ' assembled assistant text
    finishReason    As String
    deltaCount      As Long      ' chunks of *answer* text (reasoning excluded)

    ' --- server-reported prefill progress (llama.cpp /completion) ---
    prefillTotal    As Double
    prefillCache    As Double
    prefillDone     As Double

    ' --- reasoning phase ---
    reasoningCount  As Long      ' chunks of reasoning text
    reasoningChars  As Long      ' characters of reasoning received
    inThinkBlock    As Boolean   ' inside an inline <think> ... </think> block
    firstReasoningAt As Double

    ' --- phase timing ---
    streamStarted   As Boolean   ' first real chunk arrived => prefill is over
    prefillSec      As Double    ' send -> first chunk, i.e. prompt processing
    thinkSec        As Double    ' first reasoning -> first answer token
    sentAt          As Double
    startedAt       As Double

    ' --- estimation inputs ---
    estPromptTokens As Double
    calibRate       As Double
End Type

'=======================================================================
' PUBLIC ENTRY POINT
'=======================================================================

' Sends requestJson to an OpenAI-compatible /v1/chat/completions endpoint with
' streaming enabled, updates the Word status bar while the response arrives and
' returns the assembled assistant text.
'
' requestJson must already contain "stream":true (see
' modRefSigns.BuildChatCompletionJson_JSONMode).
'
' Returns True on success. On failure, outError carries a readable message.
Public Function StreamChatCompletion( _
    ByVal endpoint As String, _
    ByVal requestJson As String, _
    ByRef outContent As String, _
    ByRef outFinishReason As String, _
    ByRef outError As String, _
    Optional ByVal timeoutSec As Long = 120, _
    Optional ByVal apiKey As String = "") As Boolean

#If VBA7 Then
    Dim hOpen As LongPtr
    Dim hConn As LongPtr
    Dim hReq As LongPtr
#Else
    Dim hOpen As Long
    Dim hConn As Long
    Dim hReq As Long
#End If
    Dim hostName As String
    Dim hostPort As Long
    Dim resPath As String
    Dim useTls As Boolean
    Dim reqFlags As Long
    Dim headers As String
    Dim bodyBytes() As Byte
    Dim bodyLen As Long
    Dim statusCode As Long
    Dim msTimeout As Long
    Dim optValue As Long
    Dim atEnd As Boolean
    Dim retryWithoutStream As Boolean
    Dim lastErr As Long
    Dim st As HS_State
    Dim buf(0 To READ_BUFFER_SIZE - 1) As Byte
    Dim avail As Long
    Dim toRead As Long
    Dim gotBytes As Long

    Dim lineBuf() As Byte
    Dim lineLen As Long
    Dim lastStatus As Double
    Dim lastDataAt As Double
    Dim i As Long
    Dim b As Long
    Dim oneLine As String

    outContent = ""
    outFinishReason = ""
    outError = ""

    If timeoutSec < 5 Then timeoutSec = 5
    msTimeout = timeoutSec * 1000&

    If Not HS_ParseUrl(endpoint, hostName, hostPort, resPath, useTls) Then
        outError = "Could not parse the endpoint URL: " & endpoint
        Exit Function
    End If

    On Error GoTo Fail

    hOpen = InternetOpenA("PatentTools", INTERNET_OPEN_TYPE_PRECONFIG, _
                          vbNullString, vbNullString, 0)
    If hOpen = 0 Then
        outError = "WinINet could not be initialised."
        Exit Function
    End If

    optValue = 15000
    InternetSetOptionA hOpen, INTERNET_OPTION_CONNECT_TIMEOUT, optValue, 4
    optValue = msTimeout
    InternetSetOptionA hOpen, INTERNET_OPTION_SEND_TIMEOUT, optValue, 4

    ' Full timeout for now: the receive timeout also governs how long
    ' HttpSendRequest waits for the response headers, and cutting that short
    ' would abort the request during prompt processing. It is reduced to a short
    ' slice on the request handle only after the headers have arrived.
    InternetSetOptionA hOpen, INTERNET_OPTION_RECEIVE_TIMEOUT, optValue, 4

    hConn = InternetConnectA(hOpen, hostName, hostPort, vbNullString, _
                            vbNullString, INTERNET_SERVICE_HTTP, 0, 0)
    If hConn = 0 Then
        outError = "Could not connect to " & hostName & ":" & CStr(hostPort) & "."
        GoTo CleanUp
    End If

    reqFlags = INTERNET_FLAG_RELOAD Or INTERNET_FLAG_NO_CACHE_WRITE _
               Or INTERNET_FLAG_KEEP_CONNECTION Or INTERNET_FLAG_PRAGMA_NOCACHE
    If useTls Then reqFlags = reqFlags Or INTERNET_FLAG_SECURE

    hReq = HttpOpenRequestA(hConn, "POST", resPath, "HTTP/1.1", _
                            vbNullString, 0, reqFlags, 0)
    If hReq = 0 Then
        outError = "Could not create the HTTP request for " & resPath & "."
        GoTo CleanUp
    End If

    bodyBytes = HS_StringToUtf8(requestJson)
    bodyLen = UBound(bodyBytes) - LBound(bodyBytes) + 1

    headers = "Content-Type: application/json" & vbCrLf & _
              "Accept: text/event-stream" & vbCrLf & _
              "Cache-Control: no-cache" & vbCrLf
    If Len(Trim$(apiKey)) > 0 Then
        headers = headers & "Authorization: Bearer " & Trim$(apiKey) & vbCrLf
    End If

    HttpAddRequestHeadersA hReq, headers, Len(headers), HTTP_ADDREQ_FLAG_ADD_IF_NEW

    ' Estimated prompt size, used for the progress estimate when the server
    ' reports none. The JSON overhead is negligible next to the prompt itself.
    st.estPromptTokens = Int(Len(requestJson) / CHARS_PER_TOKEN)
    st.calibRate = HS_GetCalibratedRate()

    Application.StatusBar = HS_Spin() & " Sending " & _
                            Format$(st.estPromptTokens, "#,##0") & _
                            " estimated tokens to " & hostName & " ..."
    DoEvents

    st.sentAt = Timer

    If HttpSendRequestA(hReq, vbNullString, 0, bodyBytes(LBound(bodyBytes)), bodyLen) = 0 Then
        outError = "The request to " & hostName & " failed (connection refused, " & _
                   "TLS problem or timeout)."
        GoTo CleanUp
    End If

    statusCode = HS_GetStatusCode(hReq)

    ' Headers are in. From here on, read in short slices so that a silent server
    ' cannot block Word: a timeout is now an expected event, not an error.
    optValue = RECEIVE_SLICE_MS
    InternetSetOptionA hReq, INTERNET_OPTION_RECEIVE_TIMEOUT, optValue, 4

    st.startedAt = Timer
    lastDataAt = st.startedAt
    lastStatus = 0
    ReDim lineBuf(0 To 4095)
    lineLen = 0

    Application.StatusBar = "Waiting for the model to start ..."

    Do
        If HS_Elapsed(st.startedAt) > CDbl(timeoutSec) * MAX_TOTAL_FACTOR Then
            If Len(st.contentSoFar) = 0 Then
                outError = "Giving up after " & _
                           Format$(HS_Elapsed(st.startedAt) / 60#, "0") & _
                           " minutes without a usable answer."
                GoTo CleanUp
            End If
            Exit Do
        End If

        avail = 0
        If InternetQueryDataAvailable(hReq, avail, 0, 0) = 0 Then
            lastErr = Err.LastDllError

            If lastErr = ERROR_INTERNET_TIMEOUT Then
                ' The server is simply silent - almost always prompt processing.
                ' Keep Word alive and show an estimate until the idle deadline.
                If HS_Elapsed(lastDataAt) > timeoutSec Then
                    If Len(st.contentSoFar) = 0 And Len(st.rawBody) = 0 Then
                        outError = "No data received within " & CStr(timeoutSec) & _
                                   " s (measured from the last byte received)."
                        GoTo CleanUp
                    End If
                    Exit Do
                End If

                If HS_Elapsed(lastStatus) >= STATUS_MIN_INTERVAL Then
                    lastStatus = Timer
                    HS_ShowProgress st
                    DoEvents
                End If

                GoTo ContinueOuter
            End If

            ' Genuine transport error. If content already arrived, treat the
            ' stream as ended; otherwise report it.
            If Len(st.contentSoFar) = 0 And Len(st.rawBody) = 0 Then
                outError = "The connection failed while waiting for the response " & _
                           "(WinINet error " & CStr(lastErr) & ")."
                GoTo CleanUp
            End If
            Exit Do
        End If

        If avail = 0 Then Exit Do   ' end of response

        Do While avail > 0
            toRead = avail
            If toRead > READ_BUFFER_SIZE Then toRead = READ_BUFFER_SIZE

            gotBytes = 0
            If InternetReadFile(hReq, buf(0), toRead, gotBytes) = 0 Then
                atEnd = True
                Exit Do
            End If
            If gotBytes = 0 Then
                atEnd = True
                Exit Do
            End If

            lastDataAt = Timer

            For i = 0 To gotBytes - 1
                b = buf(i)

                If b = 10 Then                       ' LF terminates an SSE line
                    oneLine = HS_Utf8ToString(lineBuf, lineLen)
                    lineLen = 0

                    HS_HandleLine oneLine, st

                ElseIf b <> 13 Then                  ' ignore CR
                    If lineLen > UBound(lineBuf) Then
                        ReDim Preserve lineBuf(0 To (UBound(lineBuf) + 1) * 2 - 1)
                    End If
                    lineBuf(lineLen) = CByte(b)
                    lineLen = lineLen + 1
                End If
            Next i

            avail = avail - gotBytes
        Loop

        If atEnd Then Exit Do

        If HS_Elapsed(lastStatus) >= STATUS_MIN_INTERVAL Then
            lastStatus = Timer
            HS_ShowProgress st
            DoEvents
        End If

ContinueOuter:
    Loop

    ' Flush a last line that was not terminated by LF.
    If lineLen > 0 Then
        oneLine = HS_Utf8ToString(lineBuf, lineLen)
        HS_HandleLine oneLine, st
    End If

    If statusCode >= 400 Then
        ' Backwards compatibility: some OpenAI-compatible gateways reject the
        ' unknown "return_progress" field, or "stream" itself, with a 4xx instead
        ' of ignoring it. In that case retry once as a plain, non-streamed call.
        If HS_CanStripLayer(requestJson) Then
            Select Case statusCode
                Case 400, 404, 415, 422, 501
                    retryWithoutStream = True
                    GoTo CleanUp
            End Select
        End If

        outError = "Server returned HTTP " & CStr(statusCode) & "." & vbCrLf & _
                   Left$(HS_FirstNonEmpty(st.rawBody, st.contentSoFar), 600)
        GoTo CleanUp
    End If

    ' Fallback: the server ignored "stream" and answered with a single
    ' non-streamed completion object.
    If Not st.sawSseData Then
        If Len(Trim$(st.rawBody)) > 0 Then
            st.contentSoFar = HS_ExtractMessageContent(st.rawBody)

            If Len(st.contentSoFar) = 0 Then
                ' Native /completion answers non-streamed with a top-level
                ' "content" member and no "message" wrapper.
                st.contentSoFar = HS_ExtractJsonStringMember(st.rawBody, "content")
            End If

            If Len(st.finishReason) = 0 Then
                st.finishReason = HS_ExtractFinishReason(st.rawBody)
            End If
        End If
    End If

    outContent = st.contentSoFar
    outFinishReason = st.finishReason

    If Len(Trim$(outContent)) = 0 Then
        outError = "The server sent no assistant content." & vbCrLf & _
                   Left$(st.rawBody, 600)
        GoTo CleanUp
    End If

    ' Learn the prompt-processing speed for the next run.
    HS_StoreCalibration st

    StreamChatCompletion = True

CleanUp:
    If hReq <> 0 Then InternetCloseHandle hReq
    If hConn <> 0 Then InternetCloseHandle hConn
    If hOpen <> 0 Then InternetCloseHandle hOpen

    If retryWithoutStream Then
        ' Each retry removes one layer, so the recursion is finite: first the
        ' vendor extras, then streaming itself.
        Application.StatusBar = "Server rejected a request option - retrying " & _
                                "with a reduced request ..."
        DoEvents
        outError = ""
        StreamChatCompletion = StreamChatCompletion( _
            endpoint, HS_StripLayer(requestJson), _
            outContent, outFinishReason, outError, timeoutSec, apiKey)
    End If

    Exit Function

Fail:
    outError = "Streaming failed: " & Err.Description & " (error " & CStr(Err.Number) & ")"
    Resume CleanUp
End Function

'=======================================================================
' SSE HANDLING
'=======================================================================

Private Sub HS_HandleLine(ByVal rawLine As String, ByRef st As HS_State)
    Dim s As String
    Dim payload As String
    Dim delta As String
    Dim reasoning As String
    Dim fr As String
    Dim tail As String

    s = Trim$(rawLine)
    If Len(s) = 0 Then Exit Sub

    If LCase$(Left$(s, 5)) = "data:" Then
        payload = Trim$(Mid$(s, 6))
    ElseIf Left$(s, 1) = "{" Then
        ' Some servers stream bare JSON lines without the SSE "data:" prefix.
        payload = s
    Else
        ' Neither an SSE field nor the start of a JSON object. As long as nothing
        ' streaming-shaped has been seen, this may be a continuation line of a
        ' pretty-printed non-streamed response body, so keep it for the fallback.
        ' Newlines between JSON tokens are legal whitespace and cannot occur
        ' inside a string literal, so the body reassembles correctly.
        If Not st.sawSseData And Not HS_IsSseMetaLine(s) Then
            If Len(st.rawBody) < 200000 Then st.rawBody = st.rawBody & s & vbLf
        End If
        Exit Sub
    End If

    If payload = "[DONE]" Then
        st.sawSseData = True
        Exit Sub
    End If
    If Len(payload) = 0 Then Exit Sub

    ' Keep the raw text for diagnostics and for the non-streaming fallback,
    ' but do not grow it without bound.
    If Len(st.rawBody) < 200000 Then st.rawBody = st.rawBody & payload & vbLf

    ' llama.cpp prefill progress (requires "return_progress":true). Progress
    ' chunks arrive *during* prefill, so they must not count as the stream
    ' having started - that flag means "prompt processing is over".
    If InStr(1, payload, """prompt_progress""", vbBinaryCompare) > 0 Then
        st.sawSseData = True
        HS_ParsePromptProgress payload, st.prefillTotal, st.prefillCache, st.prefillDone
    End If

    ' Reasoning delivered on its own channel: llama.cpp and several gateways put
    ' it in delta.reasoning_content, deliberately outside delta.content so it
    ' never pollutes the answer. It is displayed but not accumulated.
    reasoning = HS_ExtractDeltaMember(payload, "reasoning_content")
    If Len(reasoning) > 0 Then
        HS_MarkStreamStarted st
        If st.firstReasoningAt = 0 Then st.firstReasoningAt = Timer
        st.reasoningCount = st.reasoningCount + 1
        st.reasoningChars = st.reasoningChars + Len(reasoning)
    End If

    delta = HS_ExtractDeltaContent(payload)
    If Len(delta) > 0 Then
        HS_MarkStreamStarted st
        st.contentSoFar = st.contentSoFar & delta

        ' Other templates stream reasoning inline as <think> ... </think> inside
        ' the normal content channel. Such chunks are kept (CleanupModelOutput
        ' strips the block later) but counted as thinking, not as answer, so the
        ' status bar reports the phase the user is actually waiting through.
        If st.inThinkBlock Or InStr(1, delta, "<think", vbTextCompare) > 0 Then
            st.inThinkBlock = True
            If st.firstReasoningAt = 0 Then st.firstReasoningAt = Timer
            st.reasoningCount = st.reasoningCount + 1
            st.reasoningChars = st.reasoningChars + Len(delta)

            ' The closing tag may be split across two chunks, so test the tail of
            ' the accumulated text rather than this chunk alone. Bounded length
            ' keeps this O(1) instead of rescanning the whole answer every chunk.
            tail = Right$(st.contentSoFar, Len(delta) + 16)
            If InStr(1, tail, "</think", vbTextCompare) > 0 Then st.inThinkBlock = False
        Else
            st.deltaCount = st.deltaCount + 1
            If st.thinkSec = 0 And st.firstReasoningAt > 0 Then
                st.thinkSec = HS_Elapsed(st.firstReasoningAt)
            End If
        End If
    ElseIf InStr(1, payload, """delta""", vbBinaryCompare) > 0 Then
        ' A chunk carrying only role or an empty delta: still proof that prompt
        ' processing has finished.
        HS_MarkStreamStarted st
    End If

    fr = HS_ExtractFinishReason(payload)
    If Len(fr) > 0 Then
        st.sawSseData = True
        st.finishReason = fr
    End If
End Sub

' First real chunk of the response: the prompt has been processed. The duration
' is frozen here because it is both the headline number for the user and the
' basis of the speed calibration - and it must not include thinking time.
Private Sub HS_MarkStreamStarted(ByRef st As HS_State)
    st.sawSseData = True

    If Not st.streamStarted Then
        st.streamStarted = True
        st.prefillSec = HS_Elapsed(st.sentAt)
    End If
End Sub

' Value of a string member inside the "delta" object of a streamed chunk.
Private Function HS_ExtractDeltaMember(ByVal payload As String, _
                                      ByVal key As String) As String
    Dim pDelta As Long
    Dim pKey As Long
    Dim pQuote As Long
    Dim pEnd As Long

    pDelta = InStr(1, payload, """delta""", vbBinaryCompare)
    If pDelta = 0 Then pDelta = 1

    pKey = InStr(pDelta, payload, """" & key & """", vbBinaryCompare)
    If pKey = 0 Then Exit Function

    pQuote = HS_SkipToStringStart(payload, pKey + Len(key) + 2)
    If pQuote = 0 Then Exit Function          ' e.g. "reasoning_content":null

    pEnd = HS_FindJsonStringEnd(payload, pQuote + 1)
    If pEnd = 0 Then Exit Function

    HS_ExtractDeltaMember = HS_JsonUnescape( _
        Mid$(payload, pQuote + 1, pEnd - pQuote - 1))
End Function

' True for SSE bookkeeping lines that must never be mistaken for body text.
Private Function HS_IsSseMetaLine(ByVal s As String) As Boolean
    Dim t As String

    t = LCase$(s)
    HS_IsSseMetaLine = _
        (Left$(t, 1) = ":") _
        Or (Left$(t, 6) = "event:") _
        Or (Left$(t, 3) = "id:") _
        Or (Left$(t, 6) = "retry:")
End Function

' Length of the region that may be searched for top-level members: everything
' before "messages". Claim text may legitimately contain the word "stream", and
' it must never be mistaken for a request flag.
Private Function HS_TopLevelLimit(ByVal json As String) As Long
    Dim p As Long
    Dim q As Long

    ' "messages" for the chat API, "prompt" for the native /completion API.
    ' Whichever comes first ends the region that may be edited: everything from
    ' there on is user text and must never be touched.
    p = InStr(1, json, """messages""", vbBinaryCompare)
    q = InStr(1, json, """prompt""", vbBinaryCompare)

    If p = 0 Then p = q
    If q > 0 And q < p Then p = q

    If p = 0 Then
        HS_TopLevelLimit = Len(json)
    Else
        HS_TopLevelLimit = p - 1
    End If
End Function

' True if the request body still asks for streaming.
Private Function HS_HasStreamFlags(ByVal json As String) As Boolean
    Dim head As String

    head = Left$(json, HS_TopLevelLimit(json))
    HS_HasStreamFlags = _
        (InStr(1, head, """stream""", vbBinaryCompare) > 0) _
        Or (InStr(1, head, """return_progress""", vbBinaryCompare) > 0)
End Function

' True if a vendor-specific option is present that a strict gateway may reject
' while still supporting plain streaming.
Private Function HS_HasVendorFlags(ByVal json As String) As Boolean
    Dim head As String

    head = Left$(json, HS_TopLevelLimit(json))
    HS_HasVendorFlags = _
        (InStr(1, head, """return_progress""", vbBinaryCompare) > 0) _
        Or (InStr(1, head, """sse_ping_interval""", vbBinaryCompare) > 0) _
        Or (InStr(1, head, """chat_template_kwargs""", vbBinaryCompare) > 0)
End Function

' True if anything can still be given up before declaring failure.
Private Function HS_CanStripLayer(ByVal json As String) As Boolean
    HS_CanStripLayer = HS_HasVendorFlags(json) Or HS_HasStreamFlags(json)
End Function

' Removes exactly one layer of optional request members: the vendor extras
' first, because losing them only costs progress detail, and streaming itself
' only if the reduced request is still rejected.
Private Function HS_StripLayer(ByVal json As String) As String
    If HS_HasVendorFlags(json) Then
        json = HS_RemoveJsonBoolMember(json, "return_progress")
        json = HS_RemoveJsonBoolMember(json, "sse_ping_interval")
        json = HS_RemoveJsonBoolMember(json, "chat_template_kwargs")
    Else
        json = HS_RemoveJsonBoolMember(json, "stream_options")
        json = HS_RemoveJsonBoolMember(json, "stream")
    End If

    HS_StripLayer = json
End Function

' Removes the top-level member "<key>": <value> together with one adjacent
' comma. The search is restricted to the region before "messages", so message
' content is never touched.
Private Function HS_RemoveJsonBoolMember(ByVal json As String, ByVal key As String) As String
    Dim pKey As Long
    Dim pEnd As Long
    Dim depth As Long
    Dim ch As String
    Dim head As String
    Dim p As Long

    HS_RemoveJsonBoolMember = json

    head = Left$(json, HS_TopLevelLimit(json))
    pKey = InStr(1, head, """" & key & """", vbBinaryCompare)
    If pKey = 0 Then Exit Function

    ' Walk forward to the end of the member value.
    pEnd = pKey + Len(key) + 2
    depth = 0

    Do While pEnd <= Len(json)
        ch = Mid$(json, pEnd, 1)

        If ch = "{" Or ch = "[" Then
            depth = depth + 1
        ElseIf ch = "}" Or ch = "]" Then
            If depth = 0 Then Exit Do        ' end of the enclosing object
            depth = depth - 1
        ElseIf ch = "," And depth = 0 Then
            pEnd = pEnd + 1                  ' swallow the separating comma
            Exit Do
        End If

        pEnd = pEnd + 1
    Loop

    json = Left$(json, pKey - 1) & Mid$(json, pEnd)

    ' The member may have been the last one in its object, leaving "...,}".
    ' Repair that at the splice point only - never with a global Replace.
    p = pKey
    Do While p <= Len(json)
        If Mid$(json, p, 1) <> " " Then Exit Do
        p = p + 1
    Loop

    If p <= Len(json) Then
        If Mid$(json, p, 1) = "}" Then
            p = pKey - 1
            Do While p >= 1
                If Mid$(json, p, 1) <> " " Then Exit Do
                p = p - 1
            Loop

            If p >= 1 Then
                If Mid$(json, p, 1) = "," Then
                    json = Left$(json, p - 1) & Mid$(json, p + 1)
                End If
            End If
        End If
    End If

    HS_RemoveJsonBoolMember = json
End Function

Private Sub HS_ParsePromptProgress( _
    ByVal payload As String, _
    ByRef prefillTotal As Double, _
    ByRef prefillCache As Double, _
    ByRef prefillDone As Double)

    Dim pObj As Long
    Dim scope As String
    Dim v As Double

    pObj = InStr(1, payload, """prompt_progress""", vbBinaryCompare)
    If pObj = 0 Then Exit Sub

    ' Limit the key search to the prompt_progress object itself.
    scope = Mid$(payload, pObj)
    pObj = InStr(1, scope, "}")
    If pObj > 0 Then scope = Left$(scope, pObj)

    v = HS_JsonNumber(scope, "total")
    If v >= 0 Then prefillTotal = v

    v = HS_JsonNumber(scope, "cache")
    If v >= 0 Then prefillCache = v

    v = HS_JsonNumber(scope, "processed")
    If v >= 0 Then prefillDone = v
End Sub

' Writes the current phase to the status bar. The phases are ordered from most
' to least specific, so the most informative available description wins:
'
'   1. answering   - answer tokens are arriving
'   2. thinking    - reasoning tokens are arriving, no answer yet
'   3. prefill     - the server reports real prompt_progress
'   4. starting    - the stream opened but nothing usable has arrived yet
'   5. estimate    - no progress channel; extrapolate from a measured rate
'   6. alive       - nothing is known; spinner and seconds only
Private Sub HS_ShowProgress(ByRef st As HS_State)
    Dim pct As Double
    Dim effTotal As Double
    Dim doneTokens As Double
    Dim remainSec As Double
    Dim elapsedSec As Double
    Dim waitedSec As Double
    Dim suffix As String

    elapsedSec = HS_Elapsed(st.startedAt)
    waitedSec = HS_Elapsed(st.sentAt)

    ' --- 1. answering ---
    If st.deltaCount > 0 Then
        If st.reasoningCount > 0 Then
            suffix = " (thought for " & Format$(st.thinkSec, "0") & " s)"
        End If

        Application.StatusBar = HS_Spin() & " Receiving model answer: " & _
            CStr(st.deltaCount) & " chunks, " & _
            Format$(Len(st.contentSoFar), "#,##0") & " characters in " & _
            Format$(elapsedSec, "0") & " s" & suffix
        Exit Sub
    End If

    ' --- 2. thinking ---
    If st.reasoningCount > 0 Then
        Application.StatusBar = HS_Spin() & " Model is thinking: " & _
            Format$(st.reasoningChars, "#,##0") & " characters of reasoning in " & _
            Format$(HS_Elapsed(st.firstReasoningAt), "0") & " s" & _
            " (prompt processed in " & Format$(st.prefillSec, "0") & " s)"
        Exit Sub
    End If

    ' --- 3. real prefill progress ---
    If st.prefillTotal > 0 Then
        ' Report timed progress, i.e. excluding tokens served from cache.
        effTotal = st.prefillTotal - st.prefillCache

        If effTotal > 0 Then
            pct = (st.prefillDone - st.prefillCache) / effTotal * 100#
        Else
            pct = st.prefillDone / st.prefillTotal * 100#
        End If

        If pct < 0 Then pct = 0
        If pct > 100 Then pct = 100

        Application.StatusBar = HS_Spin() & " Processing prompt: " & _
            Format$(pct, "0.0") & "% (" & Format$(st.prefillDone, "0") & "/" & _
            Format$(st.prefillTotal, "0") & " tokens, " & _
            Format$(st.prefillCache, "0") & " cached) - " & _
            Format$(elapsedSec, "0") & " s"
        Exit Sub
    End If

    ' --- 4. the stream is open but has produced nothing usable yet ---
    If st.streamStarted Then
        Application.StatusBar = HS_Spin() & " Prompt processed in " & _
            Format$(st.prefillSec, "0") & " s - model is starting to answer ..."
        Exit Sub
    End If

    ' --- 5. no progress channel: extrapolate from the measured rate ---
    If st.calibRate > 0 And st.estPromptTokens > 0 Then
        doneTokens = st.calibRate * waitedSec
        pct = doneTokens / st.estPromptTokens * 100#

        If pct > 99# Then pct = 99#     ' never claim completion we cannot see

        remainSec = (st.estPromptTokens - doneTokens) / st.calibRate
        If remainSec < 0 Then remainSec = 0

        Application.StatusBar = HS_Spin() & " Processing prompt (estimated): ~" & _
            Format$(pct, "0") & "% of ~" & Format$(st.estPromptTokens, "#,##0") & _
            " tokens at ~" & Format$(st.calibRate, "#,##0") & " tok/s - " & _
            Format$(waitedSec, "0") & " s elapsed, ~" & _
            Format$(remainSec, "0") & " s left"
        Exit Sub
    End If

    If st.estPromptTokens > 0 Then
        Application.StatusBar = HS_Spin() & " Processing prompt: ~" & _
            Format$(st.estPromptTokens, "#,##0") & " tokens sent, " & _
            Format$(waitedSec, "0") & " s elapsed " & _
            "(measuring speed for the next run)"
        Exit Sub
    End If

    ' --- 6. last resort: prove the add-in is alive ---
    Application.StatusBar = HS_Spin() & " Waiting for the model ... " & _
        Format$(elapsedSec, "0") & " s"
End Sub

' Next spinner frame. Kept deliberately to plain ASCII: the status bar font is
' whatever Word uses, and box-drawing or Braille frames render as boxes in some
' locales.
Private Function HS_Spin() As String
    Const FRAMES As String = "|/-\"

    m_spinIndex = m_spinIndex + 1
    If m_spinIndex > 4 Then m_spinIndex = 1

    HS_Spin = Mid$(FRAMES, m_spinIndex, 1)
End Function

' Reads the prompt-processing speed measured during earlier runs.
' Returns 0 when nothing has been measured yet.
Private Function HS_GetCalibratedRate() As Double
    Dim s As String

    On Error Resume Next
    s = GetSetting(CALIB_APP, CALIB_SECTION, CALIB_KEY, "")
    On Error GoTo 0

    If Len(s) = 0 Then Exit Function
    If Not IsNumeric(s) Then Exit Function

    HS_GetCalibratedRate = Val(s)
End Function

' Measures how fast the prompt was processed and stores a smoothed value.
' Server-reported token counts are preferred; otherwise the estimate is used.
' Measures how fast the prompt was processed and stores a smoothed value.
' Server-reported token counts are preferred; otherwise the estimate is used.
'
' The measured interval is send -> first streamed chunk, which is prompt
' processing only. Timing to the first *answer* token instead would fold the
' whole thinking phase into the rate and make every later estimate far too slow.
Private Sub HS_StoreCalibration(ByRef st As HS_State)
    Dim tokens As Double
    Dim rate As Double
    Dim previous As Double

    If Not st.streamStarted Then Exit Sub      ' never got a chunk to time
    If st.prefillSec < 0.5 Then Exit Sub       ' too short to measure anything

    If st.prefillTotal > 0 Then
        tokens = st.prefillTotal - st.prefillCache   ' tokens actually computed
    Else
        tokens = st.estPromptTokens
    End If

    If tokens <= 0 Then Exit Sub

    rate = tokens / st.prefillSec
    If rate <= 0 Then Exit Sub

    ' Exponential smoothing, so one cache hit or one cold start cannot
    ' invalidate the estimate for every following run.
    previous = HS_GetCalibratedRate()
    If previous > 0 Then rate = 0.7 * previous + 0.3 * rate

    ' Stored as a plain integer: Format$ would use the locale decimal separator
    ' (a comma in German), which Val would then truncate on the way back.
    On Error Resume Next
    SaveSetting CALIB_APP, CALIB_SECTION, CALIB_KEY, CStr(CLng(rate))
    On Error GoTo 0
End Sub

Private Function HS_Elapsed(ByVal sinceTimer As Double) As Double
    Dim d As Double

    d = Timer - sinceTimer
    If d < 0 Then d = d + 86400#   ' Timer wrapped around midnight

    HS_Elapsed = d
End Function

Private Function HS_FirstNonEmpty(ByVal a As String, ByVal b As String) As String
    If Len(Trim$(a)) > 0 Then
        HS_FirstNonEmpty = a
    Else
        HS_FirstNonEmpty = b
    End If
End Function

'=======================================================================
' MINIMAL JSON READERS (streaming chunks are small and well-formed)
'=======================================================================

' Returns the numeric value of "key": <number> or -1 if absent/not numeric.
Private Function HS_JsonNumber(ByVal s As String, ByVal key As String) As Double
    Dim k As Long
    Dim i As Long
    Dim ch As String
    Dim numTxt As String

    HS_JsonNumber = -1

    k = InStr(1, s, """" & key & """", vbBinaryCompare)
    If k = 0 Then Exit Function

    ' First character behind the closing quote of the key.
    i = k + Len(key) + 2

    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        If ch = ":" Then
            i = i + 1
            Exit Do
        ElseIf ch = " " Or ch = vbTab Then
            i = i + 1
        Else
            Exit Function            ' not a "key": value pair
        End If
    Loop

    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        If ch = " " Or ch = vbTab Then
            i = i + 1
        Else
            Exit Do
        End If
    Loop

    numTxt = ""
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "0" And ch <= "9") Or ch = "-" Or ch = "+" Or ch = "." _
           Or ch = "e" Or ch = "E" Then
            numTxt = numTxt & ch
            i = i + 1
        Else
            Exit Do
        End If
    Loop

    If Len(numTxt) = 0 Then Exit Function
    If Not IsNumeric(numTxt) Then Exit Function

    ' CDbl is locale-aware; JSON always uses a dot.
    HS_JsonNumber = Val(numTxt)
End Function

' choices[0].delta.content of a streaming chunk.
Private Function HS_ExtractDeltaContent(ByVal payload As String) As String
    Dim pDelta As Long
    Dim pContent As Long
    Dim pQuote As Long
    Dim pEnd As Long

    pDelta = InStr(1, payload, """delta""", vbBinaryCompare)

    If pDelta = 0 Then
        ' Native llama.cpp /completion stream: the generated text sits in a
        ' top-level "content" member, with no "choices" wrapper at all.
        ' Documented chunk shape: {"content":"...","tokens":[...],"stop":false}.
        If InStr(1, payload, """choices""", vbBinaryCompare) > 0 Then Exit Function
        pDelta = 1
    End If

    ' Note: """content""" cannot match """reasoning_content""" because the
    ' opening quote must sit directly in front of "content".
    pContent = InStr(pDelta, payload, """content""", vbBinaryCompare)
    If pContent = 0 Then Exit Function

    pQuote = HS_SkipToStringStart(payload, pContent + 9)
    If pQuote = 0 Then Exit Function          ' e.g. "content":null

    pEnd = HS_FindJsonStringEnd(payload, pQuote + 1)
    If pEnd = 0 Then Exit Function

    HS_ExtractDeltaContent = HS_JsonUnescape(Mid$(payload, pQuote + 1, pEnd - pQuote - 1))
End Function

' choices[0].message.content of a non-streamed completion object.
Private Function HS_ExtractMessageContent(ByVal jsonText As String) As String
    Dim pMsg As Long
    Dim pContent As Long
    Dim pQuote As Long
    Dim pEnd As Long

    pMsg = InStr(1, jsonText, """message""", vbBinaryCompare)
    If pMsg = 0 Then Exit Function

    pContent = InStr(pMsg, jsonText, """content""", vbBinaryCompare)
    If pContent = 0 Then Exit Function

    pQuote = HS_SkipToStringStart(jsonText, pContent + 9)
    If pQuote = 0 Then Exit Function

    pEnd = HS_FindJsonStringEnd(jsonText, pQuote + 1)
    If pEnd = 0 Then Exit Function

    HS_ExtractMessageContent = HS_JsonUnescape(Mid$(jsonText, pQuote + 1, pEnd - pQuote - 1))
End Function

' Returns the last non-null finish_reason found in the text, or "".
Private Function HS_ExtractFinishReason(ByVal jsonText As String) As String
    Dim p As Long
    Dim pQuote As Long
    Dim pEnd As Long
    Dim result As String

    p = 1
    Do
        p = InStr(p, jsonText, """finish_reason""", vbBinaryCompare)
        If p = 0 Then Exit Do

        pQuote = HS_SkipToStringStart(jsonText, p + 15)
        If pQuote > 0 Then
            pEnd = HS_FindJsonStringEnd(jsonText, pQuote + 1)
            If pEnd > 0 Then
                result = Mid$(jsonText, pQuote + 1, pEnd - pQuote - 1)
            End If
        End If

        p = p + 15
    Loop

    If Len(result) = 0 Then
        ' Native /completion reports "stop_type" instead. Documented values:
        ' "none" (still generating), "eos", "limit" (n_predict reached) and
        ' "word". Mapped onto the OpenAI vocabulary the callers already handle,
        ' because "length" is what triggers the truncation warning.
        result = HS_ExtractJsonStringMember(jsonText, "stop_type")

        Select Case LCase$(result)
            Case "limit":            result = "length"
            Case "eos", "word":      result = "stop"
            Case "none":             result = ""
        End Select
    End If

    HS_ExtractFinishReason = result
End Function

' Value of a top-level JSON string member, or "" if absent or not a string.
Private Function HS_ExtractJsonStringMember(ByVal jsonText As String, _
                                            ByVal key As String) As String
    Dim p As Long
    Dim pQuote As Long
    Dim pEnd As Long

    p = InStr(1, jsonText, """" & key & """", vbBinaryCompare)
    If p = 0 Then Exit Function

    pQuote = HS_SkipToStringStart(jsonText, p + Len(key) + 2)
    If pQuote = 0 Then Exit Function

    pEnd = HS_FindJsonStringEnd(jsonText, pQuote + 1)
    If pEnd = 0 Then Exit Function

    HS_ExtractJsonStringMember = HS_JsonUnescape( _
        Mid$(jsonText, pQuote + 1, pEnd - pQuote - 1))
End Function

' From startPos, skips ": " and returns the position of the opening quote of a
' JSON string value. Returns 0 if the value is not a string (null, number, ...).
Private Function HS_SkipToStringStart(ByVal s As String, ByVal startPos As Long) As Long
    Dim i As Long
    Dim ch As String

    i = startPos
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)

        Select Case ch
            Case """"
                HS_SkipToStringStart = i
                Exit Function
            Case ":", " ", vbTab
                i = i + 1
            Case Else
                Exit Function
        End Select
    Loop
End Function

Private Function HS_FindJsonStringEnd(ByVal s As String, ByVal startPos As Long) As Long
    Dim i As Long
    Dim ch As String
    Dim escaped As Boolean

    For i = startPos To Len(s)
        ch = Mid$(s, i, 1)

        If escaped Then
            escaped = False
        ElseIf ch = "\" Then
            escaped = True
        ElseIf ch = """" Then
            HS_FindJsonStringEnd = i
            Exit Function
        End If
    Next i
End Function

Private Function HS_JsonUnescape(ByVal s As String) As String
    Dim i As Long
    Dim ch As String
    Dim hex4 As String
    Dim out As String
    Dim pos As Long
    Dim j As Long

    out = Space$(Len(s) + 8)
    pos = 0
    i = 1

    Do While i <= Len(s)
        ch = Mid$(s, i, 1)

        If ch = "\" And i < Len(s) Then
            i = i + 1
            ch = Mid$(s, i, 1)

            Select Case ch
                Case "n": ch = vbLf
                Case "r": ch = vbCr
                Case "t": ch = vbTab
                Case "b": ch = Chr$(8)
                Case "f": ch = Chr$(12)
                Case "/": ch = "/"
                Case """": ch = """"
                Case "\": ch = "\"
                Case "u"
                    hex4 = Mid$(s, i + 1, 4)
                    If Len(hex4) = 4 And HS_IsHex4(hex4) Then
                        ch = HS_ChrWSafe(CLng("&H" & hex4))
                        i = i + 4
                    Else
                        ch = "\u"
                    End If
                Case Else
                    ch = "\" & ch
            End Select
        End If

        For j = 1 To Len(ch)
            pos = pos + 1
            If pos > Len(out) Then out = out & Space$(64)
            Mid$(out, pos, 1) = Mid$(ch, j, 1)
        Next j

        i = i + 1
    Loop

    HS_JsonUnescape = Left$(out, pos)
End Function

Private Function HS_IsHex4(ByVal s As String) As Boolean
    Dim i As Long

    For i = 1 To 4
        If InStr(1, "0123456789abcdefABCDEF", Mid$(s, i, 1), vbBinaryCompare) = 0 Then
            Exit Function
        End If
    Next i

    HS_IsHex4 = True
End Function

' ChrW$ takes a signed Integer; code points above &H7FFF must be passed as
' their negative equivalent.
Private Function HS_ChrWSafe(ByVal cp As Long) As String
    If cp > &H7FFF& Then
        HS_ChrWSafe = ChrW$(cp - &H10000&)
    Else
        HS_ChrWSafe = ChrW$(cp)
    End If
End Function

'=======================================================================
' UTF-8 <-> VBA STRING
'=======================================================================

' The old proxy used StrConv(data, vbUnicode), which interprets the bytes in
' the ANSI code page and therefore destroyed every non-ASCII character
' (umlauts, en dashes, typographic quotes). These two helpers do real UTF-8.
Public Function HS_StringToUtf8(ByVal s As String) As Byte()
    Dim out() As Byte
    Dim i As Long
    Dim n As Long
    Dim cp As Long
    Dim lo As Long

    ReDim out(0 To 4 * Len(s) + 1)
    n = 0
    i = 1

    Do While i <= Len(s)
        cp = AscW(Mid$(s, i, 1))
        If cp < 0 Then cp = cp + 65536

        ' Combine a surrogate pair into one code point.
        If cp >= &HD800& And cp <= &HDBFF& And i < Len(s) Then
            lo = AscW(Mid$(s, i + 1, 1))
            If lo < 0 Then lo = lo + 65536

            If lo >= &HDC00& And lo <= &HDFFF& Then
                cp = &H10000& + (cp - &HD800&) * &H400& + (lo - &HDC00&)
                i = i + 1
            End If
        End If

        If cp < &H80& Then
            out(n) = CByte(cp): n = n + 1
        ElseIf cp < &H800& Then
            out(n) = CByte(&HC0& Or (cp \ &H40&)): n = n + 1
            out(n) = CByte(&H80& Or (cp And &H3F&)): n = n + 1
        ElseIf cp < &H10000& Then
            out(n) = CByte(&HE0& Or (cp \ &H1000&)): n = n + 1
            out(n) = CByte(&H80& Or ((cp \ &H40&) And &H3F&)): n = n + 1
            out(n) = CByte(&H80& Or (cp And &H3F&)): n = n + 1
        Else
            out(n) = CByte(&HF0& Or (cp \ &H40000&)): n = n + 1
            out(n) = CByte(&H80& Or ((cp \ &H1000&) And &H3F&)): n = n + 1
            out(n) = CByte(&H80& Or ((cp \ &H40&) And &H3F&)): n = n + 1
            out(n) = CByte(&H80& Or (cp And &H3F&)): n = n + 1
        End If

        i = i + 1
    Loop

    If n = 0 Then
        ReDim out(0 To 0)
        out(0) = 0
        HS_StringToUtf8 = out
        Exit Function
    End If

    ReDim Preserve out(0 To n - 1)
    HS_StringToUtf8 = out
End Function

' Decodes the first count bytes of b as UTF-8. A truncated trailing sequence
' cannot occur here because lines are split on LF (0x0A), which never appears
' inside a multi-byte UTF-8 sequence.
Public Function HS_Utf8ToString(ByRef b() As Byte, ByVal count As Long) As String
    Dim out As String
    Dim pos As Long
    Dim i As Long
    Dim k As Long
    Dim c As Long
    Dim cp As Long
    Dim extra As Long
    Dim piece As String

    If count <= 0 Then Exit Function

    out = Space$(count)
    pos = 0
    i = 0

    Do While i < count
        c = b(i)

        If c < &H80& Then
            cp = c
            extra = 0
        ElseIf (c And &HE0&) = &HC0& Then
            cp = c And &H1F&
            extra = 1
        ElseIf (c And &HF0&) = &HE0& Then
            cp = c And &HF&
            extra = 2
        ElseIf (c And &HF8&) = &HF0& Then
            cp = c And &H7&
            extra = 3
        Else
            cp = &HFFFD&
            extra = 0
        End If

        If i + extra >= count Then
            cp = &HFFFD&
            extra = count - i - 1
        End If

        For k = 1 To extra
            cp = cp * &H40& + (b(i + k) And &H3F&)
        Next k

        i = i + extra + 1

        If cp > &HFFFF& Then
            cp = cp - &H10000&
            piece = HS_ChrWSafe(&HD800& + (cp \ &H400&)) & _
                    HS_ChrWSafe(&HDC00& + (cp Mod &H400&))
            pos = pos + 1
            Mid$(out, pos, 1) = Left$(piece, 1)
            pos = pos + 1
            Mid$(out, pos, 1) = Right$(piece, 1)
        Else
            pos = pos + 1
            Mid$(out, pos, 1) = HS_ChrWSafe(cp)
        End If
    Loop

    HS_Utf8ToString = Left$(out, pos)
End Function

'=======================================================================
' URL / HTTP HELPERS
'=======================================================================

Private Function HS_ParseUrl(ByVal url As String, ByRef hostName As String, _
                             ByRef hostPort As Long, ByRef resPath As String, _
                             ByRef useTls As Boolean) As Boolean
    Dim rest As String
    Dim p As Long
    Dim hostPart As String

    url = Trim$(url)

    If LCase$(Left$(url, 8)) = "https://" Then
        useTls = True
        hostPort = 443
        rest = Mid$(url, 9)
    ElseIf LCase$(Left$(url, 7)) = "http://" Then
        useTls = False
        hostPort = 80
        rest = Mid$(url, 8)
    Else
        Exit Function
    End If

    If Len(rest) = 0 Then Exit Function

    p = InStr(1, rest, "/")
    If p = 0 Then
        hostPart = rest
        resPath = "/"
    Else
        hostPart = Left$(rest, p - 1)
        resPath = Mid$(rest, p)
    End If

    ' Strip user:password@ if present.
    p = InStrRev(hostPart, "@")
    If p > 0 Then hostPart = Mid$(hostPart, p + 1)

    p = InStrRev(hostPart, ":")
    If p > 0 And InStr(1, hostPart, "]") = 0 Then
        If IsNumeric(Mid$(hostPart, p + 1)) Then
            hostPort = CLng(Mid$(hostPart, p + 1))
            hostPart = Left$(hostPart, p - 1)
        End If
    End If

    If Len(hostPart) = 0 Then Exit Function

    hostName = hostPart
    HS_ParseUrl = True
End Function

#If VBA7 Then
Private Function HS_GetStatusCode(ByVal hReq As LongPtr) As Long
#Else
Private Function HS_GetStatusCode(ByVal hReq As Long) As Long
#End If
    Dim code As Long
    Dim bufLen As Long
    Dim idx As Long

    code = 0
    bufLen = 4
    idx = 0

    If HttpQueryInfoA(hReq, HTTP_QUERY_STATUS_CODE Or HTTP_QUERY_FLAG_NUMBER, _
                      code, bufLen, idx) <> 0 Then
        HS_GetStatusCode = code
    Else
        HS_GetStatusCode = 0
    End If
End Function

'=======================================================================
' BACKEND DETECTION AND THE NATIVE llama.cpp ROUTE
'
' Two backends are supported deliberately:
'
'   * Ollama (the configured default, port 11434) speaks only the
'     OpenAI-compatible /v1/chat/completions API and reports no prompt
'     progress. It gets streaming plus the calibrated estimate.
'   * llama-server (llama.cpp) additionally offers /apply-template and
'     /completion, and /completion is the only documented endpoint that emits
'     "prompt_progress". It gets real percentages.
'
' Detection is a single cheap GET /props, whose result is cached for the Word
' session. Everything below fails soft: if any step does not work, the caller
' silently uses the OpenAI-compatible route.
'=======================================================================

' True if the base URL is a llama.cpp server exposing the native API.
Public Function PT_HasNativeProgressApi(ByVal baseUrl As String) As Boolean
    Dim body As String
    Dim errText As String

    If m_probeDone And StrComp(m_probeBase, baseUrl, vbTextCompare) = 0 Then
        PT_HasNativeProgressApi = m_probeNative
        Exit Function
    End If

    m_probeBase = baseUrl
    m_probeDone = True
    m_probeNative = False

    Application.StatusBar = "Checking what the server at " & baseUrl & " supports ..."
    DoEvents

    ' Short timeout: this must never be the reason a call feels slow.
    If PT_HttpJson("GET", baseUrl & "/props", "", "", 5, body, errText) Then
        ' Ollama has no /props at all. llama.cpp always reports build_info,
        ' and chat_template confirms a chat-capable model is loaded.
        m_probeNative = (InStr(1, body, """build_info""", vbBinaryCompare) > 0) _
                    And (InStr(1, body, """chat_template""", vbBinaryCompare) > 0)
    End If

    PT_HasNativeProgressApi = m_probeNative
End Function

' Renders chat messages into a prompt string using the model's own chat
' template, via POST /apply-template. Returns False if the endpoint is absent
' or the answer is unusable; the caller then falls back to the chat API.
Public Function PT_ApplyTemplate(ByVal baseUrl As String, _
                                 ByVal messagesArrayJson As String, _
                                 ByVal apiKey As String, _
                                 ByVal timeoutSec As Long, _
                                 ByRef outPrompt As String, _
                                 ByRef outError As String, _
                                 Optional ByVal enableThinking As Boolean = False) As Boolean

    Dim body As String
    Dim requestBody As String

    Application.StatusBar = "Applying the chat template on the server ..."
    DoEvents

    ' Pass enable_thinking via chat_template_kwargs to control Gemma 4 thinking mode
    ' CRITICAL: The template is rendered HERE, not in /completion
    If enableThinking Then
        requestBody = "{""messages"":" & messagesArrayJson & ",""chat_template_kwargs"":{""enable_thinking"":true}}"
    Else
        requestBody = "{""messages"":" & messagesArrayJson & ",""chat_template_kwargs"":{""enable_thinking"":false}}"
    End If

    If Not PT_HttpJson("POST", baseUrl & "/apply-template", requestBody, _
                       apiKey, timeoutSec, body, outError) Then
        Exit Function
    End If

    outPrompt = HS_ExtractJsonStringMember(body, "prompt")

    If Len(outPrompt) = 0 Then
        outError = "/apply-template returned no usable prompt."
        Exit Function
    End If

    PT_ApplyTemplate = True
End Function

' Minimal blocking JSON request for the two short auxiliary calls (/props and
' /apply-template). Neither runs inference, so blocking is acceptable here -
' unlike the completion call itself, which must stream.
'
' WinHttp is used rather than WinINet because no incremental reading is needed.
' The body is passed as a byte array and the response is read as ResponseBody,
' so UTF-8 is handled by the same codecs as the streaming path instead of being
' left to an implicit ANSI conversion.
Public Function PT_HttpJson(ByVal method As String, _
                            ByVal url As String, _
                            ByVal body As String, _
                            ByVal apiKey As String, _
                            ByVal timeoutSec As Long, _
                            ByRef outBody As String, _
                            ByRef outError As String) As Boolean

    Dim http As Object
    Dim ms As Long
    Dim bodyBytes() As Byte
    Dim respBytes() As Byte
    Dim respLen As Long

    On Error GoTo Failed

    ms = timeoutSec * 1000
    If ms <= 0 Then ms = 30000

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts ms, ms, ms, ms
    http.Open method, url, False
    http.SetRequestHeader "Accept", "application/json"

    If Len(apiKey) > 0 Then
        http.SetRequestHeader "Authorization", "Bearer " & apiKey
    End If

    If Len(body) > 0 Then
        http.SetRequestHeader "Content-Type", "application/json; charset=utf-8"
        bodyBytes = HS_StringToUtf8(body)
        http.Send bodyBytes
    Else
        http.Send
    End If

    If http.Status < 200 Or http.Status >= 300 Then
        outError = "HTTP " & CStr(http.Status) & " from " & url
        GoTo Cleanup
    End If

    ' ResponseBody is a byte array, so multi-byte characters survive. It cannot
    ' be passed straight into HS_Utf8ToString: that expects a typed Byte() plus
    ' a length, and an empty body leaves the array unallocated, which makes
    ' UBound raise error 9.
    respBytes = http.ResponseBody

    respLen = 0
    On Error Resume Next
    respLen = UBound(respBytes) - LBound(respBytes) + 1
    On Error GoTo Failed

    If respLen > 0 Then outBody = HS_Utf8ToString(respBytes, respLen)
    PT_HttpJson = True

Cleanup:
    Set http = Nothing
    Exit Function

Failed:
    outError = "Request to " & url & " failed: " & Err.Description
    Set http = Nothing
End Function
