VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UIClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private oCUIAutomation As New CUIAutomation

Private Enum Condition
   ui_NamePropertyId
   ui_AutomationIdPropertyId
   ui_ClassNamePropertyId
   ui_LocalizedControlTypePropertyId
   [_First] = ui_NamePropertyId
   [_Last] = ui_LocalizedControlTypePropertyId
End Enum

Private element_type As Integer

Public CurrentElementMother As String

Public CurrentElement As UIAutomationClient.IUIAutomationElement

'Set to -1 to get desktop or search a name of element on screen
Public Function Get_Element(e As String)
On Error GoTo KeepTrying
    Dim tps As Condition
    For tps = ui_NamePropertyId To ui_LocalizedControlTypePropertyId
    
        If e = CStr(-1) Then
            Set CurrentElement = oCUIAutomation.GetRootElement
        Else
            Set CurrentElement = oCUIAutomation.GetRootElement.FindFirst( _
            TreeScope_Children, _
            Get_PropertyCondition(tps, e))
        End If
    
    Next tps

    

KeepTrying:
    If CurrentElement Is Nothing Then
        Set CurrentElement = oCUIAutomation.GetRootElement
        Dim ElementArray As UIAutomationClient.IUIAutomationElementArray
        Dim Count As Integer
        Set ElementArray = CurrentElement.FindAll(TreeScope_Children, _
        oCUIAutomation.CreateTrueCondition)

        For Count = 1 To ElementArray.Length
            'if CurrentElement is null the App Root don't found the CurrentElement
            'Debug.Print "app-root: " & CurrentElement.CurrentClassName
            If CurrentElement Is Nothing Then
                Set CurrentElement = oCUIAutomation.GetRootElement
            End If
            For tps2 = ui_NamePropertyId To ui_LocalizedControlTypePropertyId
              'If StrComp(e, CurrentElement.CurrentName, vbTextCompare) <> 0 Then
              '    'getting childs from desktop
              '    Set CurrentElement = ElementArray.GetElement(Count)
                  'Debug.Print "app-root: " & CurrentElement.CurrentName & CurrentElement.CurrentClassName
             '     CurrentElementMother = CurrentElement.CurrentName

             '     Set CurrentElement = CurrentElement.FindFirst(TreeScope_Subtree, _
             '     Get_PropertyCondition(tps2, e))

              'Else
               '   Get_Element = CurrentElement

                '  Exit For
              'End If
            Next tps2
            
            'Set CurrentElement = oCUIAutomation.GetRootElement.FindFirst(TreeScope_Children, Get_PropertyCondition(ui_NamePropertyId, e))
        Next Count
    End If
End Function
















Private Function Get_PropertyCondition(Pp As Condition, Rq As String) As UIAutomationClient.IUIAutomationCondition
On Error GoTo Hell
    Select Case Pp
        Case 0
            Set Get_PropertyCondition = _
            oCUIAutomation.CreatePropertyCondition(UIAutomationClient.UIA_NamePropertyId, Rq)
        Case 1
            Set Get_PropertyCondition = _
            oCUIAutomation.CreatePropertyCondition(UIAutomationClient.UIA_AutomationIdPropertyId, Rq)
        Case 2
            Set Get_PropertyCondition = _
            oCUIAutomation.CreatePropertyCondition(UIAutomationClient.UIA_ClassNamePropertyId, Rq)
        Case 3
            Set Get_PropertyCondition = _
            oCUIAutomation.CreatePropertyCondition(UIAutomationClient.UIA_LocalizedControlTypePropertyId, Rq)
    End Select
Hell:
    ' "o inferno  aqui"
End Function



















' Busca Element(Object), Enum(1), "TargetStringName"
'Function Busca(obj As UIAutomationClient.IUIAutomationElement, _
'    typed As Condition, _
'    strFinalElemSearch As String, _
'    ByRef elem As UIAutomationClient.IUIAutomationElement) _
'    As UIAutomationClient.IUIAutomationElement
'        On Error Resume Next
'
'        Dim ended As Boolean
'        ended = False
'        Dim walker As UIAutomationClient.IUIAutomationTreeWalker
'        Dim element1 As UIAutomationClient.IUIAutomationElementArray
'        Dim element2 As UIAutomationClient.IUIAutomationElement
'
'        Set walker = oAutomation.ControlViewWalker
'        Dim condition1 As UIAutomationClient.IUIAutomationCondition
'        Set condition1 = oAutomation.CreateTrueCondition
'        Set element1 = obj.FindAll(TreeScope_Children, condition1)
'
'        'aguarda execução para q o pc possa fazer outras tarefas
'        DoEvents
'        If element1.Length <> 0 Then
'                Set element2 = obj.FindFirst(TreeScope_Children, condition1)
'        End If
'        Do While Not element2 Is Nothing
'            Select Case typed
'                Case eUIA_NamePropertyId
'                    If StrComp(element2.CurrentName, strFinalElemSearch) = 0 Then
'                        ended = True
'
'                        Set Busca = element2
'                        Exit Function
'                    End If
'
'                Case eUIA_AutomationIdPropertyId
'                    If StrComp(element2.CurrentAutomationId, strFinalElemSearch) = 0 Then
'                        ended = True
'
'                        Set Busca = element2
'                        Exit Function
'                    End If
'
'                Case eUIA_ClassNamePropertyId
'                    If StrComp(element2.CurrentClassName, strFinalElemSearch) = 0 Then
'                        ended = True
'
'                        Set Busca = element2
'                        Exit Function
'                    End If
'
'                Case eUIA_LocalizedControlTypePropertyId
'                    If StrComp(element2.CurrentLocalizedControlType, strFinalElemSearch) = 0 Then
'                        ended = True
'
'                        Set Busca = element2
'                        Exit Function
'                    End If
'            End Select
'
'            Search element2, typed, strFinalElemSearch
'
'            If Not element2 Is Nothing Then Exit Function
'
'            Set element2 = walker.GetNextSiblingElement(element2)
'
'        Loop
'        If ended Then Exit Function
'End Function


