Option Explicit

Private pScriptControl As Object
Private pAppendMode As Boolean
Private pCode As String
Private pCallback As String

Public Property Let code(ByVal c As String)
    On Error GoTo Errhandler
    
    If pAppendMode Then
        pCode = pCode & c
    Else
        pCode = c
    End If
    
    Dim mycode As String
    
    ' to convert the VBA-array to JS-array
    mycode = "function getArray(arrayIn) {" & _
                "return new VBArray(arrayIn).toArray();}" & _
            "function setArray(ja) {" & _
                "var dict = new ActiveXObject('Scripting.Dictionary');" & _
                "for (var i=0;i < ja.length; i++ )dict.add(i,ja[i]);" & _
                "return dict.items();}" & pCode & pCallback
    
    
    ' if pScriptControl already exists
    pScriptControl.addcode mycode
    
Errhandler:
    If Err.Number <> 0 Then
     Dim sc As Object
     Set sc = CreateObject("msscriptcontrol.scriptcontrol")
     
     With sc
         .Language = "jscript"
         .addcode mycode
     End With
     
     Set pScriptControl = sc
    End If
    
End Property


Public Property Get js() As Object
    Set js = pScriptControl
End Property

Public Property Let appendMode(ByVal c As Boolean)
    pAppendMode = c
End Property

Public Property Let callback(ByVal c As String)
    pCallback = c
    Me.code = pCode
End Property


' load js library from local path relative the Workbook
Public Sub loadLib(ByVal libPath As String)
    
    Dim fso As Object
    Set fso = CreateObject("scripting.filesystemobject")
    
    Dim targPath As String
    targPath = ThisWorkbook.Path & "\" & libPath
    
    Dim ts As Object
    Set ts = fso.opentextfile(targPath)
    
    pCode = pCode & ts.readall()
    
End Sub
