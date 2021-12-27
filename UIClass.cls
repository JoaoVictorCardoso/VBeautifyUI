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

Public InvokePattern As UIAutomationClient.IUIAutomationInvokePattern
Public LegacyAccPattern As UIAutomationClient.IUIAutomationLegacyIAccessiblePattern
'Set to -1 to get desktop or search a name of element on screen
Public Function Get_Element(e As String)
On Error GoTo KeepTrying
    Dim tps As Condition
    For tps = ui_NamePropertyId To ui_LocalizedControlTypePropertyId

        If e = CStr(-1) Then
            Set CurrentElement = oCUIAutomation.GetRootElement
        ElseIf Not CurrentElement Is Nothing Then
        'de uma mae busca um elemento
            Dim Element As UIAutomationClient.IUIAutomationElement
            
            'pegar todos os filhos
            Set Element = CurrentElement.FindFirst(TreeScope_Subtree, Get_PropertyCondition(tps, e))

            
            If Not Element Is Nothing Then
                'Debug.Print CurrentElement.CurrentClassName & " -> " & Element.CurrentName
                
                Debug.Print CurrentElement.CurrentAutomationId & "-(" & Element.CurrentLocalizedControlType & ")"
                
                Set CurrentElement = Element
                Get_Element = Element
                GoTo finaliza
                Exit For
            End If
        Else
        'se elemento vazio busca da raiz busca um elemento
            Set CurrentElement = oCUIAutomation.GetRootElement.FindFirst( _
            TreeScope_Children, _
            Get_PropertyCondition(tps, e))
            '
            
            If Not CurrentElement Is Nothing Then
                CurrentElementMother = CurrentElement.CurrentName
                'Debug.Print CurrentElement.CurrentName
                'Exit For
                Get_Element = CurrentElement
                GoTo finaliza
                Exit For
                
            End If
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
                'Set CurrentElement = oCUIAutomation.GetRootElement
                GoTo jumP
            End If
            For tps = ui_NamePropertyId To ui_LocalizedControlTypePropertyId
            
                Set CurrentElement = ElementArray.GetElement(Count)
                
                CurrentElementMother = CurrentElement.CurrentName
                
                Set CurrentElement = CurrentElement.FindFirst(TreeScope_Subtree, _
                Get_PropertyCondition(tps, e))
                
                If Not CurrentElement Is Nothing Then
                    
                    GoTo finaliza
                    Exit For
                End If
            Next tps


jumP:
           ' Else
            '    Get_Element = CurrentElement
            '
            '    Exit For
            'End If

            'Set CurrentElement = oCUIAutomation.GetRootElement.FindFirst(TreeScope_Children, Get_PropertyCondition(ui_NamePropertyId, e))
        Next Count
    Else
        'se tem mãe, existe.
        Dim Elemn As UIAutomationClient.IUIAutomationElement
        Dim tipos As Condition
        
        For tipos = ui_NamePropertyId To ui_LocalizedControlTypePropertyId
            'pegar todos os filhos
            Set Elemn = oCUIAutomation.GetRootElement.FindFirst(TreeScope_Subtree, Get_PropertyCondition(tipos, e))
            
            If Not Elemn Is Nothing Then
                'Debug.Print CurrentElement.CurrentClassName & Elemn.CurrentName
                Set CurrentElement = Elemn
                Get_Element = Elemn
                Exit For
            End If
        
        Next tipos
    
    End If

finaliza:
    'Debug.Print ""
End Function

Public Function Fill(txt As Variant, Optional ByVal elemt_ As Variant = Nothing)
On Error GoTo n_clic
    If elemt_ Is Nothing Then
        Set LegacyAccPattern = elemt_.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
        
        LegacyAccPattern.SetValue (txt)

    Else
        Set LegacyAccPattern = CurrentElement.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
        
        LegacyAccPattern.SetValue (txt)
    End If
n_clic:
    If CurrentElement Is Nothing Then
    Else
        
        Set LegacyAccPattern = CurrentElement.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
        LegacyAccPattern.CurrentValue
        
    End If

End Function


Public Function Clique(Optional ByVal element_ As Variant = Nothing)
On Error GoTo nao_clic
    If element_ Is Nothing Then
        Set InvokePattern = CurrentElement.GetCurrentPattern(UIAutomationClient.UIA_InvokePatternId)
        If CurrentElement.CurrentIsEnabled Then
            InvokePattern.Invoke
        End If
    Else
        Set InvokePattern = element_.GetCurrentPattern(UIAutomationClient.UIA_InvokePatternId)
        If element_.CurrentIsEnabled Then
            InvokePattern.Invoke
        End If
    End If
    
nao_clic:
    'Debug.Print "Error nao clickou"
    If Not element_ Is Nothing Then
        Set InvokePattern = element_.GetCurrentPattern(UIAutomationClient.UIA_InvokePatternId)
        If element_.CurrentIsEnabled Then
            InvokePattern.Invoke
        End If
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
    ' "o inferno rola aqui"
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




