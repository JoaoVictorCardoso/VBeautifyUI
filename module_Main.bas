Attribute VB_Name = "module_Main"
Option Explicit
Sub Main()
    Dim uie As New UIClass
   

    uie.Get_Element ("Formatar")
    
    Debug.Print uie.CurrentElementMother
    
    Debug.Print uie.CurrentElement.CurrentLocalizedControlType
    
    
End Sub

