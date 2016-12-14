'<script language="vbscript">

' This is used by a group of controls: 2 Radio - 1 Edit
' When a Radio is selected the Edit control is disabled or...
' or is Focused
Sub TwoRadioOneEditDisable(RadioDisableControl, EditControl)
 If RadioDisableControl.Checked then
   EditControl.Disabled = true
 Else
   EditControl.Disabled = false
   EditControl.Focus
 End If  
End Sub


' Reads a simple date control and returs the result as a Date.
' If erros occur nothing is returned
Function GetDateFromSimpleControl(ctrlname)
 Dim Data
 On Error Resume Next
 Data = DateSerial(document.all(ctrlname & "an").value, document.all(ctrlname & "ln").value, document.all(ctrlname & "zi").value)
 If Err.number <> 0 then
  Err.Clear 
 Else 
  GetDateFromSimpleControl = Data
 End If 
End Function


' Enable/Disable a simple date control
Sub SimpleControlDisabled(ctrlname, state)
 document.all(ctrlname & "an").disabled = state
 document.all(ctrlname & "ln").disabled = state
 document.all(ctrlname & "zi").disabled = state
End Sub
