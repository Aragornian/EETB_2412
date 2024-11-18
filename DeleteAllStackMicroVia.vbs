Option Explicit
' Add any type libraries to be used. 
Scripting.AddTypeLibrary("MGCPCB.ExpeditionPCBApplication")
 
' Get the Application object 
Dim pcbAppObj
Set pcbAppObj = Application
  
' Get the active document
Dim pcbDocObj
Set pcbDocObj = pcbAppObj.ActiveDocument

' License the document
ValidateServer(pcbDocObj)

' Get the selected vias collection
Dim viaColl
Set viaColl = pcbDocObj.Vias(epcbSelectSelected)

' Number of selected position vias
Dim ViaArraySize
ViaArraySize = 0

Call DeleteStackMicroVias(viaColl)

' Delete all stackmicrovias
Sub DeleteStackMicroVias(ByVal viaColl)

    Dim viaselectedobj
    For Each viaselectedobj In viaColl
        
        ' If selected via is still valid in the design. Delete it
        If viaselectedobj.IsValid Then
            ' Get the Position of selected via
            Dim Px, Py 
            Px = viaselectedobj.PositionX
            Py = viaselectedobj.PositionY

            ' Get the start and end layer of selected via
            Dim Slayer, Elayer
            Slayer = viaselectedobj.StartLayer
            Elayer = viaselectedobj.EndLayer

            ' Get the selected via net
            Dim viaNetobj
            Set viaNetobj = viaselectedobj.Net

            ' Get all vias for the net
            Dim viaCollNet
            Set viaCollNet = viaNetobj.Vias

            Dim ViaPropertyArray
            ViaPropertyArray = GetViaProperties(viaCollNet, Px, Py)

            ' Sort vias
            Dim i, j, tmp
            For i = 0 To ViaArraySize - 2 Step 1
                For j = 0 To ViaArraySize - 2 Step 1
                    If ViaPropertyArray(j)(1) > ViaPropertyArray(j + 1)(1) Then
                        tmp = ViaPropertyArray(j + 1)
                        ViaPropertyArray(j + 1) = ViaPropertyArray(j)
                        ViaPropertyArray(j) = tmp
                    End If
                Next
            Next

            ' Delete stackvias
            Call DeleteViasBelow(ViaPropertyArray, Slayer, Elayer)
            Call DeleteViasAbove(ViaPropertyArray, Slayer, Elayer)
        End If

    Next

End Sub

' Get properties of selected position vias
Function GetViaProperties(ByVal viaCollNet, ByVal Px, ByVal Py)

    Dim viaobj, ViaPropertyArray(), ViaSpanArray(3)
    ViaArraySize = 0
    For Each viaobj In viaCollNet
        If viaobj.PositionX = Px And viaobj.PositionY = Py Then
            Set ViaSpanArray(0) = viaobj
            ViaSpanArray(1) = viaobj.StartLayer
            ViaSpanArray(2) = viaobj.EndLayer
            ReDim Preserve ViaPropertyArray(ViaArraySize)
            ViaPropertyArray(ViaArraySize) = ViaSpanArray
            ViaArraySize = ViaArraySize + 1
        End If
    Next
    GetViaProperties = ViaPropertyArray
    
End Function

' Delete stackmicrovias below
Sub DeleteViasBelow(ByVal ViaPropertyArray, ByVal Slayer, ByVal Elayer)

    Do
        Dim countInt, searchFlag
        ' Search flag
        searchFlag = 1
        For countInt = 0 To ViaArraySize - 1 Step 1
            ' Search stackmicrovias below
            If ViaPropertyArray(countInt)(1) = Slayer + 1 And ViaPropertyArray(countInt)(2) = Elayer + 1 Then
                Call ViaPropertyArray(countInt)(0).Delete()
                Slayer = Slayer + 1
                Elayer = Elayer + 1
            End If
        Next
        searchFlag = 0
    Loop Until searchFlag = 0

End Sub

' Delete stackmicrovias above
Sub DeleteViasAbove(ByVal ViaPropertyArray, ByVal Slayer, ByVal Elayer)

    Do
        Dim countInt, searchFlag
        ' Search flag
        searchFlag = 0
        For countInt = ViaArraySize - 1 To 0 Step -1
            ' Search stackmicrovias above
            If ViaPropertyArray(countInt)(1) = Slayer And ViaPropertyArray(countInt)(2) = Elayer Then
                Call ViaPropertyArray(countInt)(0).Delete()
                Slayer = Slayer - 1
                Elayer = Elayer - 1
            End If
        Next
        searchFlag = 0
    Loop Until searchFlag = 0

End Sub

' Server validation function
Function ValidateServer(docObj)
      
    Dim keyInt
    Dim licenseTokenInt
    Dim licenseServerObj
  
    ' Ask Expedition’s document for the key
    keyInt = docObj.Validate(0)
  
    ' Get license server
    Set licenseServerObj = CreateObject("MGCPCBAutomationLicensing.Application")
  
    ' Ask the license server for the license token
    licenseTokenInt = licenseServerObj.GetToken(keyInt)
  
    ' Release license server
    Set licenseServerObj = nothing
  
    ' Turn off error messages.  
    On Error Resume Next
    Err.Clear
  
    ' Ask the document to validate the license token
    docObj.Validate(licenseTokenInt)
    If Err Then
        ValidateServer = 0    
    Else
        ValidateServer = 1
    End If
  
End Function