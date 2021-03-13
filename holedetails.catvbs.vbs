'Written by Surya Selvam 03-01-2021

Private Sub CommandButton1_Click()

Dim ColeHoles As Collection
Set ColeHoles = New Collection
  
'Set Document = CATIA.ActiveDocument

Set filesys = CATIA.FileSystem
crlf = Chr(10)
filename = "c:\temp\Holes.txt"
If filesys.FileExists(filename) Then
  filesys.DeleteFile (filename)
End If

Set File = filesys.CreateFile(filename, True)
Set stream = File.OpenAsTextStream("ForWriting")
'**************************************************************
Dim curDoc  As Document

For s = 1 To CATIA.Documents.Count

Dim curShapes  As Shapes
Set curDoc = CATIA.Documents.Item(s)

If InStr(curDoc.Name, ".CATPart") > 0 Then

Set oPartDoc = curDoc 'CATIA.ActiveDocument
oPartDoc.Activate
Set oBody = oPartDoc.Part.Bodies.Item("PartBody")
Set curShapes = oBody.Shapes

For m = 1 To curShapes.Count
    Dim curShape  As Shape
    Set curShape = curShapes.Item(m)
         If curShape.Name Like "Hole.*" Then
         
        'variant Object returning array
        Dim sHole As Variant
        Set sHole = curShape
        
        'get the origin
        Dim Origin(2)
        
        sHole.GetOrigin Origin
            stream.Write (sHole.Name & "," & sHole.Diameter.Value & "," & sHole.Type & ", x=" & Origin(0) & ";y=" & Origin(1) & ";z=" & Origin(2) & crlf)
        
        Dim CYSdist(3) As String
        CYSdist(0) = sHole.Name
        CYSdist(1) = Origin(0)
        CYSdist(2) = Origin(1)
        CYSdist(3) = Origin(2)
        
        'ColeHoles.Item (CYSdist)
        ColeHoles.Add Item:=CYSdist
        End If
      Next
End If
Next
' Valdiate Hole is in Co axial
Dim matchHole As Collection
Set matchHole = New Collection


For sm = 1 To ColeHoles.Count

Dim X As String
Dim y As String
Dim Z As String

 X = ColeHoles(sm)(1)
 y = ColeHoles(sm)(2)
 Z = ColeHoles(sm)(3)
 
        For ms = 1 To ColeHoles.Count
           Dim CoAxial(1) As String
           If ms <> sm Then
                  If (X = ColeHoles(ms)(1) And y = ColeHoles(ms)(2)) Or (y = ColeHoles(ms)(2) And Z = ColeHoles(ms)(3)) Or (X = ColeHoles(ms)(1) And Z = ColeHoles(ms)(3)) Then
                  
                  CoAxial(0) = ColeHoles(sm)(0)
                  CoAxial(1) = ColeHoles(ms)(0)
                  matchHole.Add Item:=CoAxial
                  
                  End If
                   
           End If
        Next
 'holeinforloop
 Next
 'stream.Write ("Co - axial Hole fetaures in this product")
 For k = 1 To matchHole.Count
   stream.Write (matchHole(k)(0) & " and " & matchHole(k)(1) & " is Co axial" & crlf)
Next

MsgBox "Completed"

End Sub

