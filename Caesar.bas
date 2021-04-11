Attribute VB_Name = "Caesar"
'Caesar cipher encoding routines
Public sen As Long
Public logg As String

Public Function caesarD(ByVal tx As String) As String
Dim rev As String
Dim tp As String
rev = reverse(tx)
For k = 1 To Len(rev)
  s = Asc(Mid(rev, k, 1)) - 3
  tp = tp & Chr(s)
Next
caesarD = tp
End Function

Public Function caesarE(ByVal tx As String) As String
Dim rev As String
Dim tp As String
rev = reverse(tx)
For k = 1 To Len(rev)
  s = Asc(Mid(rev, k, 1)) + 3
  tp = tp & Chr(s)
Next
caesarE = tp
End Function

Private Function reverse(ByVal da As String) As String
Dim temp As String
temp = ""
For p = Len(da) To 1 Step -1
  f = Mid(da, p, 1)
  temp = temp & f
Next
reverse = temp
End Function
