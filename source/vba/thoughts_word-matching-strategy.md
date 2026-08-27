# Word Matching Fallback Strategies for InsertReferenceSignsOnly

## Issue Discovered (Date: 2025-01-xx)

When model outputs contain **literal Unicode escape sequences** as text (e.g., `u2003cpropelled` instead of actual EM SPACE character), the matching algorithm fails even though semantically it should match.

### Debug Trace from Real Failure
```
Original token: [selfpropelled]
Model token 1: [self]
Model token 2: [u2003cpropelled]

Forward Skip: failed (looking for selfpropelled in later model tokens)
Merge: combined "self" + "u2003cpropelled" → "selfu2003cpropelled" ≠ "selfpropelled" ✗
Bidirectional Lookahead: no matching sequence found

Result: Hard failure, error shown to user
```

## Current Recovery Mechanisms (All Failed)

1. **Forward Skip (lines 907-916):** Tolerates model omitting 1-2 original tokens
2. **Model-Token Merge (lines 934-957):** Combines up to 4 model tokens to match 1 original token  
3. **Bidirectional Lookahead (lines 978-990):** Skips 1-2 tokens on BOTH sides before re-attempting match

## Why They All Failed Here

The core issue is that the merge logic produces `"selfu2003cpropelled"` which doesn't canonical-match `"selfpropelled"`. The literal `\u` escape sequence corrupts the comparison.

---

## Option A: Clean Unicode Escapes in NormalizeAnalysisText (Primary Solution)

Add detection and removal of literal escape sequences before tokenization:

```vba
' In modJsonHelper.bas - NormalizeAnalysisText function

' Remove/fix literal Unicode escape sequences that some LLMs output as text
' Patterns like \u2003, U+2003, etc. can appear as literal ASCII in model output
Do While InStr(s, "\u") > 0 Or InStr(s, "U+") > 0
    s = RemoveUnicodeEscapeAndTrailingChar(s)
Loop

Private Function RemoveUnicodeEscapeAndTrailingChar(ByVal s As String) As String
    Dim pos As Long
    Dim hexPart As String
    
    ' Check for \uXXXX pattern (lowercase u with 4 hex digits)
    pos = InStr(s, "\u")
    If pos > 0 And pos + 5 <= Len(s) Then
        hexPart = Mid$(s, pos + 2, 4)
        If IsHexDigits(hexPart) Then
            ' Remove \uXXXX (6 chars) and any single trailing letter like 'c'
            s = Left$(s, pos - 1) & Mid$(s, pos + 7)
            RemoveUnicodeEscapeAndTrailingChar = s
            Exit Function
        End If
    End If
    
    ' Check for U+XXXX pattern (uppercase with plus)  
    pos = InStr(s, "U+")
    If pos > 0 And pos + 5 <= Len(s) Then
        hexPart = Mid$(s, pos + 2, 4)
        If IsHexDigits(hexPart) Then
            ' Remove U+XXXX (5 chars) and any single trailing letter like 'c'
            s = Left$(s, pos - 1) & Mid$(s, pos + 6)
            RemoveUnicodeEscapeAndTrailingChar = s
            Exit Function
        End If
    End If
    
    RemoveUnicodeEscapeAndTrailingChar = s
End Function

Private Function IsHexDigits(ByVal s As String) As Boolean
    Dim i As Long, ch As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If Not ((ch >= "0" And ch <= "9") Or (ch >= "A" And ch <= "F") Or (ch >= "a" And ch <= "f")) Then
            IsHexDigits = False
            Exit Function
        End If
    Next i
    IsHexDigits = True
End Function
```

**Pros:** Clean, handles root cause, works before tokenization
**Cons:** Requires modifying normalization logic in modJsonHelper.bas

---

## Option B: Lenient Substring Fallback (Alternative/Supplement)

Add a final fallback check when all other strategies fail - allow skipping tokens if the text contains the other as substring:

```vba
' After bidirectional lookahead fails, before hard failure:

Dim originalCanonical As String, modelCanonical As String
originalCanonical = CStr(oWordsCmp(iO)(0))
modelCanonical = CStr(mWordsCmp(iM)(0))

' Lenient fallback: check if one is substring of other
' Handles cases like "selfpropelled" vs "u2003cpropelled" 
If InStr(originalCanonical, modelCanonical) > 0 Or InStr(modelCanonical, originalCanonical) > 0 Then
    ' Text contains the other - likely corruption/escape sequence issue
    ' Skip this token on whichever side is corrupted and try to continue
    iO = iO + 1
    iM = iM + 1
    GoTo NextLoop
End If
```

**Pros:** Simple addition, handles corruption cases without full normalization change
**Cons:** May be too aggressive (false positives), doesn't fix root cause

---

## Option C: Character Similarity Threshold (More Complex)

Add Levenshtein distance check as final fallback - if edit distance < 30% of token length, treat as match:

```vba
' Calculate edit distance between original and model token
Dim dist As Long, maxLen As Long, threshold As Double
dist = LevenshteinDistance(originalCanonical, modelCanonical)
maxLen = IIf(Len(originalCanonical) > Len(modelCanonical), Len(originalCanonical), Len(modelCanonical))
threshold = 0.3 ' 30% tolerance

If dist / maxLen < threshold Then
    ' Near-match: accept and continue
    If Len(refText) > 0 Then
        ' Insert reference sign at original position
        insertAt = CLng(oWordsPos(iO)(2)) + delta
        Set insRng = targetRng.Duplicate
        insRng.SetRange targetRng.Start + insertAt, targetRng.Start + insertAt
        insRng.InsertAfter " " & refText
        delta = delta + Len(refText) + 1
    End If
    iO = iO + 1
    iM = iM + 1
    GoTo NextLoop
End If
```

**Pros:** Most robust for various corruption types (typos, escape sequences, etc.)
**Cons:** Requires implementing Levenshtein algorithm, more CPU overhead

---

## Recommendation

**Primary:** Implement Option A (Unicode escape cleanup in normalization) - fixes the root cause cleanly

**Backup/Complementary:** Consider Option B (substring fallback) as a final safety net before hard failure - catches edge cases without much complexity

**Option C** could be useful if similar issues arise with other types of corruption, but adds more code complexity.

---

## Additional Notes

- The current merge logic ALREADY handles "2 words vs 1 word" scenario
- The problem is specifically that the merge produces corrupted text due to literal escape sequences
- Without fixing Unicode escapes, making skip logic more lenient may not help much for THIS case
- However, substring/character similarity fallbacks could help for OTHER corruption types (misspellings, etc.)
