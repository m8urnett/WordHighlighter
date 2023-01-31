VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "modMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'Callback for btnYellow onAction
Sub btnYellowClick(control As IRibbonControl)
   If Len(Selection.Text) Then
      With Selection.Shading
         .Texture = wdTextureNone
         .ForegroundPatternColor = wdColorAutomatic
         .BackgroundPatternColor = 8451836
      End With
   End If
End Sub

'Callback for btnBlue onAction
Sub btnBlueClick(control As IRibbonControl)
   If Len(Selection.Text) Then
      With Selection.Shading
         .Texture = wdTextureNone
         .ForegroundPatternColor = wdColorAutomatic
         .BackgroundPatternColor = 15389890
      End With
   End If
End Sub

'Callback for btnGreen onAction
Sub btnGreenClick(control As IRibbonControl)
   If Len(Selection.Text) Then
      With Selection.Shading
         .Texture = wdTextureNone
         .ForegroundPatternColor = wdColorAutomatic
         .BackgroundPatternColor = 13106133
      End With
   End If
End Sub

'Callback for btnOrange onAction
Sub btnOrangeClick(control As IRibbonControl)
   If Len(Selection.Text) Then
      With Selection.Shading
         .Texture = wdTextureNone
         .ForegroundPatternColor = wdColorAutomatic
         .BackgroundPatternColor = 9422335
      End With
   End If
End Sub

'Callback for btnPink onAction
Sub btnPinkClick(control As IRibbonControl)
   If Len(Selection.Text) Then
      With Selection.Shading
         .Texture = wdTextureNone
         .ForegroundPatternColor = wdColorAutomatic
         .BackgroundPatternColor = 11448049
      End With
   End If
End Sub

'Callback for btnPurple onAction
Sub btnPurpleClick(control As IRibbonControl)
   If Len(Selection.Text) Then
      With Selection.Shading
         .Texture = wdTextureNone
         .ForegroundPatternColor = wdColorAutomatic
         .BackgroundPatternColor = 15845336
      End With
   End If
End Sub

'Callback for btnClear onAction
Sub btnClearHighlightsClick(control As IRibbonControl)
   If Len(Selection.Text) Then
      With Selection.Shading
         .Texture = wdTextureNone
         .ForegroundPatternColor = wdColorAutomatic
         .BackgroundPatternColor = wdColorAutomatic
      End With
   End If
End Sub
