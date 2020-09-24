VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "XML Note Creator"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   8535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cSave 
      Caption         =   "Save XML Document"
      Height          =   375
      Left            =   2040
      TabIndex        =   11
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox XMLSource 
      Height          =   4215
      Left            =   4080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   360
      Width           =   4335
   End
   Begin VB.CommandButton cCreate 
      Caption         =   "Create XML Document"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox tNote 
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1680
      Width           =   3735
   End
   Begin VB.TextBox tObject 
      Height          =   285
      Left            =   600
      TabIndex        =   3
      Top             =   840
      Width           =   3255
   End
   Begin VB.TextBox tTo 
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
   Begin VB.TextBox tFrom 
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label Label5 
      Caption         =   "XML Source:"
      Height          =   255
      Left            =   4080
      TabIndex        =   10
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lNote 
      Caption         =   "Note:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label lObj 
      Caption         =   "Object:"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lTo 
      Alignment       =   1  'Right Justify
      Caption         =   "To:"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lFrom 
      Caption         =   "From:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cCreate_Click()

'The main goal of this application is to show how to make a standard document from user input info.

Dim xmltext As String

'These are the main parameters for an XML document.
'The first declare the version and the other the related
'style sheet.

xmltext = "<?xml version=""1.0"" encoding=""ISO-8859-1""?>" & vbCrLf
xmltext = xmltext & "<?xml-stylesheet href=""note.css"" type=""text/css""?>" & vbCrLf & vbCrLf

'Now starting the main <NOTE>
xmltext = xmltext & "<NOTE>" & vbCrLf
'Putting the title
xmltext = xmltext & "<TITLE>XMLNote Creator by Vinchenzo (www.planet-source-code.com)</TITLE>" & vbCrLf & vbCrLf
'Starting the header
xmltext = xmltext & "<HEADER>" & vbCrLf & vbCrLf
'The destinator
xmltext = xmltext & "<TO>" & vbCrLf & "To: <NAME>" & tTo.Text & "</NAME>" & vbCrLf & "</TO>" & vbCrLf & vbCrLf
'The Sender
xmltext = xmltext & "<FROM>" & vbCrLf & "From: <NAME>" & tFrom.Text & "</NAME>" & vbCrLf & "</FROM>" & vbCrLf & vbCrLf
'The date
xmltext = xmltext & "<TIMEDATE>" & vbCrLf & "Date: <DATE>" & Date & "</DATE> - <TIME>" & Time & "</TIME>" & vbCrLf & "</TIMEDATE>" & vbCrLf & vbCrLf
'The subject
xmltext = xmltext & "<SUBJECT>" & vbCrLf & "Subject: <DESC>" & tObject & "</DESC>" & vbCrLf & "</SUBJECT>" & vbCrLf & vbCrLf & "</HEADER>" & vbCrLf & vbCrLf & vbCrLf
'The main message
xmltext = xmltext & "<BODY>" & vbCrLf & "Message:" & vbCrLf & "<MESSAGE>" & vbCrLf & tNote.Text & vbCrLf & "</MESSAGE>" & vbCrLf & "</BODY>" & vbCrLf & vbCrLf & vbCrLf
'The end of the XML text
xmltext = xmltext & "</NOTE>"

XMLSource.Text = xmltext

End Sub

Private Sub cSave_Click()
Dim tmpi As Integer
Dim tmpi2 As Integer
Dim tmpi3 As Integer

Dim filename As String
If FileExists(App.Path & "\note.xml") = True Then
    tmpi = Int((Rnd(1) * 10) + 1)
    tmpi2 = Int((Rnd(1) * 10) + 1)
    tmpi3 = Int((Rnd(1) * 10) + 1)
    If FileExists(App.Path & "\note" & tmpi & tmpi2 & tmpi3 & ".xml") = True Then
        MsgBox "Could not save xml file", , "Error"
        Exit Sub
    Else
        filename = App.Path & "\note" & tmpi & tmpi2 & tmpi3 & ".xml"
    End If
Else
    filename = App.Path & "\note.xml"
End If


Open filename For Output As 1
Print #1, XMLSource.Text
Close 1

MsgBox "File " & filename & " was successfully saved!", , "Saved"
End Sub


Public Function FileExists(filename As String) As Boolean

On Error GoTo nonexist
Open filename For Input As 1
    FileExists = True
Close 1

Exit Function
nonexist:
FileExists = False

End Function
