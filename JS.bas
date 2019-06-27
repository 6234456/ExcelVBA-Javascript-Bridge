Option Explicit

Private pScriptControl As Object
Private pAppendMode As Boolean
Private pCode As String
Private pCallback As String
Private pDefaultCode As String

Private Sub Class_Initialize()
     Set pScriptControl = CreateObjectx86("msscriptcontrol.scriptcontrol")
    
    ' to convert the VBA-array to JS-array and parser to the VBA-Object
    pDefaultCode = "function getArray(arrayIn) {" & _
                "return new VBArray(arrayIn).toArray();}" & _
            "function setArray(ja) {" & _
                "var dict = new ActiveXObject('Scripting.Dictionary');" & _
                "for (var i=0;i < ja.length; i++ )dict.add(i,ja[i]);" & _
                "return dict.items();}" & _
            "function parserJSON(s){ return eval('(' + s + ')');}" & _
            "function getType(o) {return {}.toString.call(o).slice(8, -1)}" & _
            "function parser(obj){" & _
                "var type = getType(obj); var res = new ActiveXObject('Scripting.Dictionary');" & _
                "if(type == 'Array'){ for(var i = 0 ; i < obj.length ; i++) res.add(i, parser(obj[i])); return res.items();}" & _
                "else if(type == 'Object'){ for(var i in obj) res.add(i, obj[i]); return res;}" & _
                "else return obj; }"
     
    pScriptControl.Language = "jscript"
End Sub

Private Sub Class_Terminate()
   Set pScriptControl = Nothing
   
   #If Win64 Then
   CreateObjectx86 , True
    #End If
End Sub

Public Property Let code(ByVal c As String)
    
    If pAppendMode Then
        pCode = pCode & c
    Else
        pCode = c
    End If
    
    Dim mycode As String
    
    ' to convert the VBA-array to JS-array and parser to the VBA-Object
    mycode = pDefaultCode & _
            pCode & pCallback
    
    ' if pScriptControl already exists
    pScriptControl.addcode mycode
    
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


' social.msdn.microsoft.com/Forums/en-US/c6e9c23f-a455-4138-ad86-954d95420739/excel-vba-compatibility-issues-with-microsoft-scriptcontrol-10%3Fforum%3Disvvba
Function CreateObjectx86(Optional sProgID, Optional bClose = False)

    #If Win64 Then
        Static oWnd As Object
       Dim bRunning As Boolean
       
        bRunning = InStr(TypeName(oWnd), "HTMLWindow") > 0
        If bClose Then
            If bRunning Then oWnd.Close
            Exit Function
        End If
        
        If Not bRunning Then
            Set oWnd = CreateWindow()
            oWnd.execScript "Function CreateObjectx86(sProgID): Set CreateObjectx86 = CreateObject(sProgID): End Function", "VBScript"
        End If
        Set CreateObjectx86 = oWnd.CreateObjectx86(sProgID)
    #Else
        Set CreateObjectx86 = CreateObject(sProgID)
    #End If

End Function

Function CreateWindow()

    ' source http://forum.script-coding.com/viewtopic.php?pid=75356#p75356
    Dim sSignature, oShellWnd, oProc

    On Error Resume Next
    sSignature = Left(CreateObject("Scriptlet.TypeLib").GUID, 38)
    CreateObject("WScript.Shell").Run "%systemroot%\syswow64\mshta.exe about:""about:<head><script>moveTo(-32000,-32000);document.title='x86Host'</script><hta:application showintaskbar=no /><object id='shell' classid='clsid:8856F961-340A-11D0-A96B-00C04FD705A2'><param name=RegisterAsBrowser value=1></object><script>shell.putproperty('" & sSignature & "',document.parentWindow);</script></head>""", 0, False
    Do
        For Each oShellWnd In CreateObject("Shell.Application").Windows
            Set CreateWindow = oShellWnd.GetProperty(sSignature)
            If Err.Number = 0 Then Exit Function
            Err.clear
        Next
    Loop

End Function
