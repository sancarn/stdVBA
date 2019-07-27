VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    'init
    Dim sb As clsStringBuilder
    Set sb = New clsStringBuilder
    
    'concatenation
    sb.Append "Very "
    sb.Append "many "
    sb.Append "pieces "
    sb.Append "of data"
    
    'result
    Debug.Print "All data: "; sb.ToString
    Debug.Print "Len of data: "; sb.length
    
    'removing 5 characters from position # 1
    sb.Remove 1, 5
    Debug.Print "New data: "; sb.ToString
    
    'inserting string in position # 1
    sb.Insert 1, "Not "
    Debug.Print "New data: "; sb.ToString
    
    'overwrite part of the text from position # 1 (same like MID(str,x) = "...")
    sb.Overwrite 1, "How"
    Debug.Print "New data: "; sb.ToString
    
    'getting 2 first characters
    Debug.Print "2 left chars: "; sb.ToStringLeft(2)
    
    'getting 2 characters from the end
    Debug.Print "2 end chars: "; sb.ToStringRight(2)
    
    'getting 2 characters from the middle, beginning from position 6
    Debug.Print "2 middle chars: "; sb.ToStringMid(6, 2)
    
    'getting a pointer to a NUL terminated string to use somehow (e.g. write on disk by ptr, WriteFile, e.t.c.)
    'this method is much faster than .ToString()
    'warning: you should use this pointer before calling next any method of StringBuilder, that may cause changing its data
    Debug.Print "ptr to string: "; sb.ToStringPtr
    
    'replacing the data (same as Clear + Append)
    sb.StringData = "Anew"
    Debug.Print "New data: "; sb.ToString
    
    'go to the next line (append CrLf)
    sb.AppendLine ""
    'append second line with CrLf at the end
    sb.AppendLine "Second line"
    'append third line without CrLf
    sb.Append "Third Line"
    
    Debug.Print sb.ToString
    
    'clear all data
    sb.Clear
    Debug.Print "Len of data (after clear): "; sb.length
    
    'Search samples (by default search is case sensitive)
    'Set new data
    sb.StringData = "|textile|Some|text|to|search"
    Debug.Print "New data: "; sb.ToString
    
    'Simple search ('text' will be found inside 'textile' word)
    Debug.Print "Position of 'text': " & sb.Find(1, "text")
    
    'Simple search (start search from position 3)
    Debug.Print "Position of 'text': " & sb.Find(3, "text")
    
    Debug.Print "'some' (case sensitive): " & sb.Find(1, "some")
    Debug.Print "'some' (case insensitive): " & sb.Find(1, "some", , vbTextCompare)
    
    'Search by delimiter
    Debug.Print "searching for |text|: " & sb.Find(1, "text", "|")
    
    'Search for empty string, saved with delimiter |
    Debug.Print "empty string (delim = '|'): " & sb.Find(1, "", "|")
    
    'Undo operations
    sb.StringData = "Some data "
    Debug.Print "Orig. string: " & sb.ToString
    
    sb.Append "remove"
    Debug.Print "After Append: " & sb.ToString
    'or you can use .UndoAppend
    sb.Undo
    Debug.Print "After Undo:   " & sb.ToString
    
    sb.Insert 6, "bad "
    Debug.Print "After Insert: " & sb.ToString
    'or you can use .UndoInsert
    sb.Undo
    Debug.Print "After Undo:   " & sb.ToString
    
    sb.Overwrite 1, "Here"
    Debug.Print "After Overwrite: " & sb.ToString
    'or you can use .UndoOverwrite
    sb.Undo
    Debug.Print "After Undo:      " & sb.ToString
    
    sb.Remove 2, 5
    Debug.Print "After Remove: " & sb.ToString
    'or you can use .UndoRemove
    sb.Undo
    Debug.Print "After Undo:   " & sb.ToString
    
    'when you finished work with the class
    Set sb = Nothing
    
    Unload Me
End Sub
