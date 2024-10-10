x = MsgBox("Did you have your cafelutsa yet?" ,vbYesNo + vbQuestion, "Reminder cafelutsa")
If x = vbYes Then
                    MsgBox "O zi de colectie, suflet cald!" ,0 + vbInformation, "Reminder cafelutsa" 
                ElseIf vbNo Then
                    MsgBox "drink your cafelutsa babygirl." ,0 + vbCritical, "Reminder cafelutsa" 
                End If

