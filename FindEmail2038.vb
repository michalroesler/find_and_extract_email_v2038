Option Explicit

Sub FindEmail2038()           '[0-9;A-z;,._-]{1;}\@[0-9;A-z;._-]{1;}
                              '[0-9;A-z;,._-]{1;}\@[0-9;A-z;._-]{1;}
                              'this code works fine - it loops through all email addresses found in the contract
                              
Dim wordApp As Word.Application
Dim wordDoc As Word.Document        'and it's concatenating them into the destination string variable
Dim excelApp As Excel.Application   'sometimes this code grabs the next line after the last email address, which is the first line of the next patagraph "2. Wszelka korespondencja związana z wykonaniem niniejszej umowy, w tym "
Dim rng As Word.Range
Dim srchRng As Word.Range

Dim emailAdr As String
Dim ws As Excel.Worksheet
Dim iCount As Integer
Dim eAddresses As Object
Dim addressNum As Integer
Dim wykr As Range
Dim flag As Boolean
                              
Set wordApp = GetObject(, "Word.Application")
Set excelApp = GetObject(, "Excel.Application")
Set wordDoc = wordApp.ActiveDocument
Set eAddresses = CreateObject("Scripting.Dictionary")

Set rng = wordApp.ActiveDocument.Content            ' shorter alternative code is ```Set rng = ActiveDocument.Content```
Set ws = excelApp.ActiveSheet
Set wykr = ActiveSheet.Range("A10:J50").Find("wykreślenie", , xlValues)
flag = False

excelApp.Application.Visible = True
wordApp.Application.Visible = True
Debug.Print wykr.Address
addressNum = 1

    With rng.Find
        .Text = "@"           'we only look for the @ character,
        .Wrap = wdFindStop
        .Forward = True
        .MatchWildcards = False
        .Execute
        'Debug.Print rng.text
        
        Do While .Found       'therefore we need to build whole email addres around this @ character;
            Set srchRng = rng.Duplicate
            srchRng.MoveStartUntil Cset:=" ", Count:=wdBackward      'therefore we need to build whole email addres around this @ character;
            srchRng.MoveEndUntil Cset:=","            'ask the question how to MoveEndUntil "," or ";" or ".";
            Debug.Print srchRng.Text
            
            If flag = True Then         'this if statement omits the first email address found in the contract;
                emailAdr = emailAdr & ", " & srchRng.Text           'assign value to emailAdr variable
                Debug.Print emailAdr
            End If
            
            If Not eAddresses.Exists(srchRng.Text) Then
                eAddresses.Add srchRng.Text, addressNum
                addressNum = addressNum + 1
            End If
            .Execute
            flag = True
        Loop
    End With
    emailAdr = Right(emailAdr, Len(emailAdr) - 2)
    Debug.Print emailAdr
    
    'emailAdr input into the cell
    'ws.Range("C32").Value = emailAdr
    wykr.Offset(0, 2) = emailAdr
    
End Sub
