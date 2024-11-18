Option Explicit

' Get the application object
Dim pcbAppObj
Set pcbAppObj = Application

' Get the active document
Dim pcbDocObj
Set pcbDocObj = pcbAppObj.ActiveDocument

' Get the GUI object
Dim pcbGuiObj
Set pcbGuiObj = pcbAppObj.Gui

' License the document
ValidateServer(pcbDocObj)

pcbAppObj.Gui.CursorBusy(True)
pcbDocObj.TransactionStart

If pcbAppObj.LockServer = True Then

    ' Get the display control object
    Dim displayCtrlObj, utilityObj
    Set displayCtrlObj = pcbDocObj.ActiveViewEx.DisplayControl
    Set utilityObj = pcbAppObj.Utility

    '--------------------------------------------------------------------------------
	' Conductor layer display control
	'--------------------------------------------------------------------------------
	Dim conductorLayersColl: Set conductorLayersColl = pcbDocObj.LayerStack(False)
	Dim i, conductorLayerCount
    conductorLayerCount = conductorLayersColl.Count
    Dim colorsArr(15,2)
    ' Dark Green
    colorsArr(0,0) = 0 
    colorsArr(0,1) = 102
    colorsArr(0,2) = 0
    ' Pink
    colorsArr(1,0) = 255 
    colorsArr(1,1) = 153
    colorsArr(1,2) = 255
    ' Teal
    colorsArr(2,0) = 102
    colorsArr(2,1) = 255
    colorsArr(2,2) = 255
    ' Blue
    colorsArr(3,0) = 51
    colorsArr(3,1) = 102
    colorsArr(3,2) = 153
    ' Green
    colorsArr(4,0) = 51
    colorsArr(4,1) = 204
    colorsArr(4,2) = 51
    ' Dark Blue Purple
    colorsArr(5,0) = 102
    colorsArr(5,1) = 102
    colorsArr(5,2) = 153
    ' Purple
    colorsArr(6,0) = 204
    colorsArr(6,1) = 0
    colorsArr(6,2) = 255
    ' Dark Yellow
    colorsArr(7,0) = 204
    colorsArr(7,1) = 153
    colorsArr(7,2) = 0
    ' Blue
    colorsArr(8,0) = 51
    colorsArr(8,1) = 102
    colorsArr(8,2) = 255
    ' Red
    colorsArr(9,0) = 255
    colorsArr(9,1) = 51
    colorsArr(9,2) = 0
    ' Yellow
    colorsArr(10,0) = 255
    colorsArr(10,1) = 255
    colorsArr(10,2) = 0
    ' Dark Yellow
    colorsArr(11,0) = 204
    colorsArr(11,1) = 153
    colorsArr(11,2) = 0
    ' Green Yellow
    colorsArr(12,0) = 102
    colorsArr(12,1) = 102
    colorsArr(12,2) = 51
    ' Cyan
    colorsArr(13,0) = 51
    colorsArr(13,1) = 153
    colorsArr(13,2) = 102
    ' Yellow Green
    colorsArr(14,0) = 153
    colorsArr(14,1) = 102
    colorsArr(14,2) = 51
    ' Little Pink
    colorsArr(15,0) = 153
    colorsArr(15,1) = 51
    colorsArr(15,2) = 102
    displayCtrlObj.Global.Color( "LayerControl.1" ) = utilityObj.NewColorPattern( 0, 102, 0, 100, 0, False, True )
    displayCtrlObj.Global.Color( "LayerControl." & conductorLayerCount ) = utilityObj.NewColorPattern( 204, 153, 0, 100, 0, False, True )
    Dim redColor, greenColor, blueColor
    For i = 1 To conductorLayerCount
        displayCtrlObj.Visible( "Copper.Trace." & i ) = epcbGraphicsItemStateOnEnabled
        displayCtrlObj.Visible( "Copper.Pad." & i ) = epcbGraphicsItemStateOnEnabled
        displayCtrlObj.Visible( "Copper.Plane.Data." & i ) = epcbGraphicsItemStateOnEnabled
        If i > 1 And i < conductorLayerCount Then
            redColor = colorsArr(i-1,0)
            greenColor = colorsArr(i-1,1)
            blueColor = colorsArr(i-1,2)
            displayCtrlObj.Global.Color( "LayerControl." & i ) = utilityObj.NewColorPattern( redColor, greenColor, blueColor, 100, 0, False, True )
        End If
		displayCtrlObj.Visible( "LayerControl." & i ) = epcbGraphicsItemStateOnEnabled	
	Next
    displayCtrlObj.Option( "Option.ActiveLayerOnly" ) = epcbGraphicsItemStateOnEnabled

    '--------------------------------------------------------------------------------
	' All other layers display control
	'--------------------------------------------------------------------------------
	With displayCtrlObj
        ' Global View & Interactive Selection
        .Option( "Option.Planning.Enabled" ) = epcbGraphicsItemStateOnEnabled 
                .Visible( "Group.Outline.Top" ) = epcbGraphicsItemStateOffEnabled 
                .Visible( "Group.Outline.Bubble.Top" ) = epcbGraphicsItemStateOffEnabled 
                .Visible( "Part.PlaceOutline.Top" ) = epcbGraphicsItemStateOnEnabled 
                .Option( "Option.PlaceObjects.ObstructsAndRooms.Top" ) = epcbGraphicsItemStateOnEnabled 
            .Option( "Option.PlaceObjects.Parts.Top" ) = epcbGraphicsItemStateOnEnabled 
                .Visible( "Group.Outline.Bottom" ) = epcbGraphicsItemStateOffEnabled 
                .Visible( "Group.Outline.Bubble.Bottom" ) = epcbGraphicsItemStateOffEnabled 
                .Visible( "Part.PlaceOutline.Bottom" ) = epcbGraphicsItemStateOnEnabled 
                .Option( "Option.PlaceObjects.ObstructsAndRooms.Bottom" ) = epcbGraphicsItemStateOnEnabled 
            .Option( "Option.PlaceObjects.Parts.Bottom" ) = epcbGraphicsItemStateOnEnabled 
        .Option( "Option.PlaceObjects" ) = epcbGraphicsItemStateOnEnabled 
            .Option( "Option.Traces.Enabled" ) = epcbGraphicsItemStateOnEnabled 
            .Option( "Option.Vias.Enabled" ) = epcbGraphicsItemStateOnEnabled 
            .Option( "Option.Pins.Enabled" ) = epcbGraphicsItemStateOnEnabled 
            .Option( "Option.Netlines.Enabled" ) = epcbGraphicsItemStateOnEnabled 
            .Option( "Option.Planes.Enabled" ) = epcbGraphicsItemStateOnEnabled 
            .Option( "Option.RouteObstructs.Enabled" ) = epcbGraphicsItemStateOnEnabled 
            .Option( "Option.RouteAreas.Enabled" ) = epcbGraphicsItemStateOnEnabled 
            .Option( "Option.ConductiveShapes.Enabled" ) = epcbGraphicsItemStateOnEnabled 
            .Option( "Option.Teardrops.Enabled" ) = epcbGraphicsItemStateOnEnabled 
        .Option( "Option.RouteObjects.Enabled" ) = epcbGraphicsItemStateOnEnabled 
            .Option( "Option.RFNodes.Enabled" ) = epcbGraphicsItemStateOnEnabled 
            .Option( "Option.RFShapes.Enabled" ) = epcbGraphicsItemStateOnEnabled 
            .Option( "Option.RFObstructs.Enabled" ) = epcbGraphicsItemStateOnEnabled 
        .Option( "Option.RFObjects.Enabled" ) = epcbGraphicsItemStateOnEnabled 
                .Option( "Option.DiePinsTop.Enabled" ) = epcbGraphicsItemStateOnEnabled 
                .Option( "Option.BondWiresTop.Enabled" ) = epcbGraphicsItemStateOnEnabled 
                .Visible( "WirebondObjects.WirebondGuides.Top" ) = epcbGraphicsItemStateOnEnabled 
            .Option( "Option.WirebondItemsTop.Enabled" ) = epcbGraphicsItemStateOnEnabled 
                .Option( "Option.DiePinsBottom.Enabled" ) = epcbGraphicsItemStateOnEnabled 
                .Option( "Option.BondWiresBottom.Enabled" ) = epcbGraphicsItemStateOnEnabled 
                .Visible( "WirebondObjects.WirebondGuides.Bottom" ) = epcbGraphicsItemStateOnEnabled 
            .Option( "Option.WirebondItemsBottom.Enabled" ) = epcbGraphicsItemStateOnEnabled 
        .Option( "Option.WirebondItems.Enabled" ) = epcbGraphicsItemStateOnEnabled 
            .Option( "Option.Fiducials.Enabled" ) = epcbGraphicsItemStateOnEnabled 
            .Option( "Option.Holes.Enabled" ) = epcbGraphicsItemStateOnEnabled
            .Option( "Option.BoardElements.Enabled" ) = epcbGraphicsItemStateOnEnabled 
        .Option( "Option.BoardObjects.Enabled" ) = epcbGraphicsItemStateOnEnabled 
            .Option( "Option.FabricationObjects" ) = epcbGraphicsItemStateOnEnabled 
            .Option( "Option.CopperBalancing.Enabled" ) = epcbGraphicsItemStateOnEnabled 
            .Option( "Option.Materials.Enabled" ) = epcbGraphicsItemStateOnEnabled 
            .Option( "Option.Fabrication.DrillDrawing" ) = epcbGraphicsItemStateOnEnabled 
            .Option( "Option.UserDraftLayers.Enabled" ) = epcbGraphicsItemStateOnEnabled 
        .Option( "Option.DrawFabObjects.Enabled" ) = epcbGraphicsItemStateOnEnabled 
        .Visible( "Fabrication.DetailViews" ) = epcbGraphicsItemStateOnEnabled 

		' Route/Multi Planning
                .Option( "Option.VirtualPins.Enabled" ) = epcbGraphicsItemStateOnEnabled
                .Option( "Option.DiffPairCenterlines.Enabled" ) = epcbGraphicsItemStateOnEnabled
                .Option( "Option.UnpackedAreas.Enabled" ) = epcbGraphicsItemStateOnEnabled 
                .Option( "Option.BusPaths.Enabled" ) = epcbGraphicsItemStateOnEnabled 
                .Option( "Option.TargetAreas.Enabled" ) = epcbGraphicsItemStateOnEnabled 
                .Option( "Option.RouteTargets.Enabled" ) = epcbGraphicsItemStateOnEnabled 
            .Option( "Option.RoutePlanning" ) = epcbGraphicsItemStateOnEnabled

                .Option( "Option.SketchPlans.Width" ) = epcbGraphicsItemStateOnEnabled
            .Option( "Option.SketchPlans.Enabled" ) = epcbGraphicsItemStateOnEnabled
            .Option( "Option.ReuseArea.Enabled" ) = epcbGraphicsItemStateOnEnabled
                    .Option( "Option.MultipleDesigners.TeamPCB.ShadowMode" ) = epcbGraphicsItemStateOffEnabled
				.Visible( "General.MultipleDesigners.TeamPCB.ReservedAreas" ) = epcbGraphicsItemStateOnEnabled 
				.Visible( "General.MultipleDesigners.Xtreme.ProtectedAreas" ) = epcbGraphicsItemStateOnEnabled 
				.Visible( "Board.Sandbox" ) = epcbGraphicsItemStateOnEnabled
			.Option( "Option.MultipleDesigners" ) = epcbGraphicsItemStateOffEnabled
		.Option( "Option.Planning.Enabled" ) = epcbGraphicsItemStateOnEnabled
        '.Lock
        .Global.Color( "Planning.ReuseArea" ) = utilityObj.NewColorPattern( 0, 255, 0, 100, 0, False, True ) 
        .Global.Color( "General.MultipleDesigners.TeamPCB.ReservedAreas" ) = utilityObj.NewColorPattern( 51, 153, 51, 100, 0, False, True ) 
        .Global.Color( "General.MultipleDesigners.Xtreme.ProtectedAreas" ) = utilityObj.NewColorPattern( 0, 128, 0, 100, 0, False, True ) 
        .Global.Color( "Board.Sandbox" ) = utilityObj.NewColorPattern( 0, 128, 0, 100, 0, False, True ) 
        '.Unlock

        ' Place
			.Visible( "Group.Outline.Top" ) = epcbGraphicsItemStateOffEnabled 
			.Visible( "Group.Outline.Bottom" ) = epcbGraphicsItemStateOffEnabled
			.Visible( "Group.Outline.Bubble.Top" ) = epcbGraphicsItemStateOffEnabled 
			.Visible( "Group.Outline.Bubble.Bottom" ) = epcbGraphicsItemStateOffEnabled 
			.Visible( "Place.Part.Text.RefDes.Top" ) = epcbGraphicsItemStateOffEnabled 
			.Visible( "Place.Part.Text.RefDes.Bottom" ) = epcbGraphicsItemStateOffEnabled
			
                .Option( "Option.SelectableInsidePartOutline" ) = epcbGraphicsItemStateOffEnabled
                .Option( "Option.FillPartOutlineOnSelection" ) = epcbGraphicsItemStateOffEnabled
            .Visible( "Part.PlaceOutline.Top" ) = epcbGraphicsItemStateOnEnabled 
			.Visible( "Part.PlaceOutline.Bottom" ) = epcbGraphicsItemStateOnEnabled
			
				.Visible( "Board.Obstruct.Part.Top" ) = epcbGraphicsItemStateOnEnabled 
				.Visible( "Board.Obstruct.Part.Bottom" ) = epcbGraphicsItemStateOnEnabled 
				.Visible( "Board.Obstruct.TestPoint.Top" ) = epcbGraphicsItemStateOnEnabled 
				.Visible( "Board.Obstruct.TestPoint.Bottom" ) = epcbGraphicsItemStateOnEnabled 
				.Visible( "Board.Room.Top" ) = epcbGraphicsItemStateOnEnabled 
				.Visible( "Board.Room.Bottom" ) = epcbGraphicsItemStateOnEnabled 	
			.Option( "Option.PlaceObjects.ObstructsAndRooms.Top" ) = epcbGraphicsItemStateOffEnabled 
			.Option( "Option.PlaceObjects.ObstructsAndRooms.Bottom" ) = epcbGraphicsItemStateOffEnabled 
		
		
				.Visible( "Part.InsertionOutline.Top" ) = epcbGraphicsItemStateOnEnabled 
				.Visible( "Part.InsertionOutline.Bottom" ) = epcbGraphicsItemStateOnEnabled 
				.Visible( "Part.Hazard.Top" ) = epcbGraphicsItemStateOnEnabled 
				.Visible( "Part.Hazard.Bottom" ) = epcbGraphicsItemStateOnEnabled 
				.Option( "Option.Pin.Number.Top" ) = epcbGraphicsItemStateOnEnabled 
				.Option( "Option.Pin.Number.Bottom" ) = epcbGraphicsItemStateOnEnabled 
				.Option( "Option.Pin.Type.Top" ) = epcbGraphicsItemStateOffEnabled 
				.Option( "Option.Pin.Type.Bottom" ) = epcbGraphicsItemStateOffEnabled 
                .Option( "Option.Pin.NetName.Top" ) = epcbGraphicsItemStateOnEnabled
                .Option( "Option.Pin.NetName.Bottom" ) = epcbGraphicsItemStateOnEnabled 
			.Option( "Option.PlaceObjects.PartItems.Top" ) = epcbGraphicsItemStateOnEnabled 
			.Option( "Option.PlaceObjects.PartItems.Bottom" ) = epcbGraphicsItemStateOnEnabled 
        .Option( "Option.PlaceObjects.Parts.Top" ) = epcbGraphicsItemStateOnEnabled 
        .Option( "Option.PlaceObjects.Parts.Bottom" ) = epcbGraphicsItemStateOnEnabled 
		.Option( "Option.PlaceObjects" ) = epcbGraphicsItemStateOnEnabled
        'Lock
        .Global.Color( "Group.Outline.Top" ) = utilityObj.NewColorPattern( 204, 0, 255, 100, 0, False, True ) 
        .Global.Color( "Group.Outline.Bottom" ) = utilityObj.NewColorPattern( 255, 0, 0, 100, 0, False, True )
        .Global.Color( "Group.Outline.Bubble.Top" ) = utilityObj.NewColorPattern( 255, 0, 0, 100, 0, False, True ) 
        .Global.Color( "Group.Outline.Bubble.Bottom" ) = utilityObj.NewColorPattern( 0, 0, 255, 100, 0, False, True ) 
        .Global.Color( "Place.Part.Text.RefDes.Top" ) = utilityObj.NewColorPattern( 255, 102, 0, 100, 0, False, True ) 
        .Global.Color( "Place.Part.Text.RefDes.Bottom" ) = utilityObj.NewColorPattern( 0, 0, 255, 100, 0, False, True ) 
        .Global.Color( "Part.PlaceOutline.Top" ) = utilityObj.NewColorPattern( 0, 255, 255, 100, 2, False, True ) 
        .Global.Color( "Part.PlaceOutline.Bottom" ) = utilityObj.NewColorPattern( 204, 0, 153, 100, 2, False, True ) 
        .Global.Color( "Board.Obstruct.Part.Top" ) = utilityObj.NewColorPattern( 255, 255, 153, 100, 0, False, True ) 
        .Global.Color( "Board.Obstruct.Part.Bottom" ) = utilityObj.NewColorPattern( 0, 255, 204, 100, 0, False, True ) 
        .Global.Color( "Board.Obstruct.TestPoint.Top" ) = utilityObj.NewColorPattern( 255, 0, 0, 100, 0, False, True ) 
        .Global.Color( "Board.Obstruct.TestPoint.Bottom" ) = utilityObj.NewColorPattern( 102, 0, 255, 100, 0, False, True ) 
        .Global.Color( "Board.Room.Top" ) = utilityObj.NewColorPattern( 0, 0, 255, 100, 0, False, True ) 
        .Global.Color( "Board.Room.Bottom" ) = utilityObj.NewColorPattern( 192, 192, 192, 100, 0, False, True ) 
        .Global.Color( "Part.InsertionOutline.Top" ) = utilityObj.NewColorPattern( 0, 255, 255, 100, 0, False, True ) 
        .Global.Color( "Part.InsertionOutline.Bottom" ) = utilityObj.NewColorPattern( 204, 0, 153, 100, 0, False, True ) 
        .Global.Color( "Part.Hazard.Top" ) = utilityObj.NewColorPattern( 255, 102, 0, 100, 0, False, True ) 
        .Global.Color( "Part.Hazard.Bottom" ) = utilityObj.NewColorPattern( 255, 102, 0, 100, 0, False, True ) 
        .Global.Color( "Part.Pin.NumberType.Top" ) = utilityObj.NewColorPattern( 255, 255, 255, 100, 0, False, True ) 
        .Global.Color( "Part.Pin.NumberType.Bottom" ) = utilityObj.NewColorPattern( 255, 255, 255, 100, 0, False, True ) 
        'Unlock

        ' Vias
			.Option( "Option.Pad.Via.AllSameColor" ) = epcbGraphicsItemStateOffEnabled 
			.Option( "Option.ViaPads.Enabled" ) = epcbGraphicsItemStateOnEnabled 
			.Visible( "Fabrication.Hole.Via" ) = epcbGraphicsItemStateOnEnabled 
			.Visible( "General.Via.SpanNumbers" ) = epcbGraphicsItemStateOnEnabled 
            .Option( "Option.Vias.NetNames" ) = epcbGraphicsItemStateOnEnabled 
			.Visible( "General.Via.InactiveBlindBuriedPad" ) = epcbGraphicsItemStateOffEnabled 
		.Option( "Option.Vias.Enabled" ) = epcbGraphicsItemStateOnEnabled
        'Lock
        .Global.Color( "General.Via.AllSameColor" ) = utilityObj.NewColorPattern( 255, 255, 0, 100, 0, False, True ) 
        .Global.Color( "Fabrication.Hole.Via" ) = utilityObj.NewColorPattern( 255, 153, 0, 100, 0, False, True ) 
        .Global.Color( "General.Via.SpanNumbers" ) = utilityObj.NewColorPattern( 255, 255, 255, 100, 0, False, True ) 
        .Global.Color( "General.Via.InternalSkipViaPad" ) = utilityObj.NewColorPattern( 221, 221, 221, 100, 0, False, True ) 
        .Global.Color( "General.Via.InactiveBlindBuriedPad" ) = utilityObj.NewColorPattern( 41, 41, 41, 100, 0, False, True ) 
        'Unlock

        ' Pins
			.Option( "Option.Pad.Through.AllSameColor" ) = epcbGraphicsItemStateOffEnabled 
			.Option( "Option.Pin.Through.Enabled" ) = epcbGraphicsItemStateOnEnabled 
			.Visible( "Fabrication.Hole.Pin" ) = epcbGraphicsItemStateOnEnabled 
			.Option( "Option.SMDPinPads.Enabled" ) = epcbGraphicsItemStateOnEnabled 
            .Option( "Option.BondPads.Enabled" ) = epcbGraphicsItemStateOffEnabled
			.Visible( "Copper.Pad.TestPoint.Top" ) = epcbGraphicsItemStateOnEnabled 
			.Visible( "Copper.Pad.TestPoint.Bottom" ) = epcbGraphicsItemStateOnEnabled 
		.Option( "Option.Pins.Enabled" ) = epcbGraphicsItemStateOnEnabled 
        'Lock
        .Global.Color( "General.Pin.AllSameColor" ) = utilityObj.NewColorPattern( 0, 128, 0, 100, 0, False, True ) 
        .Global.Color( "Fabrication.Hole.Pin" ) = utilityObj.NewColorPattern( 255, 153, 0, 100, 0, False, True ) 
        .Global.Color( "Copper.Pad.TestPoint.Top" ) = utilityObj.NewColorPattern( 102, 153, 255, 100, 0, False, True ) 
        .Global.Color( "Copper.Pad.TestPoint.Bottom" ) = utilityObj.NewColorPattern( 255, 51, 153, 100, 0, False, True ) 
        'Unlock

        ' Netlines
                .StringOption( "Option.Netlines.DynamicFiltering.Mode" ) = "BothEnds" 
                .Option( "Option.Netlines.DynamicFiltering.Freeze" ) = epcbGraphicsItemStateOffEnabled 
            .Option( "Option.Netlines.DynamicFiltering" ) = epcbGraphicsItemStateOffEnabled 
            
            .Visible( "Netline.NonOrderedOpen" ) = epcbGraphicsItemStateOnEnabled 
            .Visible( "Netline.OrderedOpen" ) = epcbGraphicsItemStateOnEnabled 
            .Visible( "Netline.OrderedAll" ) = epcbGraphicsItemStateOffEnabled 
            .Option( "Option.Netlines.DisplayFromFilteredNets" ) = epcbGraphicsItemStateOffEnabled
            .Option( "Option.Netlines.DisplayBetweenMarkedComponents" ) = epcbGraphicsItemStateOffEnabled 
            .Option( "Option.Netlines.DisplayFromMarkedComponents" ) = epcbGraphicsItemStateOffEnabled 
            .Option( "Option.Netlines.DisplayFromMarkedNets" ) = epcbGraphicsItemStateOffEnabled 
                .Option( "Option.Classlines.Netlines" ) = epcbGraphicsItemStateOnEnabled 
            .Option( "Option.Classlines" ) = epcbGraphicsItemStateOffEnabled 
        .Option( "Option.Netlines.Enabled" ) = epcbGraphicsItemStateOffEnabled
        'Lock
        .Global.Color( "Netline.NonOrderedOpen" ) = utilityObj.NewColorPattern( 192, 192, 192, 100, 0, False, True ) 
        .Global.Color( "Netline.OrderedOpen" ) = utilityObj.NewColorPattern( 255, 153, 204, 100, 0, False, True ) 
        .Global.Color( "Netline.OrderedAll" ) = utilityObj.NewColorPattern( 255, 204, 153, 100, 0, False, True ) 
        'Unlock

        ' Planes
			.Option( "Option.Planes.Data.Fill" ) = epcbGraphicsItemStateOnEnabled 
			.Option( "Option.Planes.Data.Enabled" ) = epcbGraphicsItemStateOnEnabled 
			.Option( "Option.Planes.Shape.Enabled" ) = epcbGraphicsItemStateOnEnabled 
			.Option( "Option.Planes.Sketch.Enabled" ) = epcbGraphicsItemStateOnEnabled 
		.Option( "Option.Planes.Enabled" ) = epcbGraphicsItemStateOnEnabled

        ' Route Obstructs
			.Option( "Option.RouteObstructs.Pad.Enabled" ) = epcbGraphicsItemStateOnEnabled 
			.Option( "Option.RouteObstructs.Plane.Enabled" ) = epcbGraphicsItemStateOnEnabled 
			.Option( "Option.RouteObstructs.Trace.Enabled" ) = epcbGraphicsItemStateOnEnabled 
			.Option( "Option.RouteObstructs.TraceVia.Enabled" ) = epcbGraphicsItemStateOnEnabled 
			.Option( "Option.RouteObstructs.Via.Enabled" ) = epcbGraphicsItemStateOnEnabled 
			.Option( "Option.RouteObstructs.TuningPattern.Enabled" ) = epcbGraphicsItemStateOnEnabled 
			    .Option( "Option.Spacers.ShadowModeEnabled" ) = epcbGraphicsItemStateOffEnabled 
			.Option( "Option.Spacers.Enabled" ) = epcbGraphicsItemStateOnEnabled 
		.Option( "Option.RouteObstructs.Enabled" ) = epcbGraphicsItemStateOnEnabled
        
        ' Route Areas
			.Visible( "Board.RouteBorder" ) = epcbGraphicsItemStateOnEnabled 
			.Visible( "Board.RouteFence.Hard" ) = epcbGraphicsItemStateOnEnabled 
			.Visible( "Board.RouteFence.Soft" ) = epcbGraphicsItemStateOnEnabled 
			.Option( "Option.RuleAreas.Enabled" ) = epcbGraphicsItemStateOnEnabled 
		.Option( "Option.RouteAreas.Enabled" ) = epcbGraphicsItemStateOnEnabled 
        'Lock
        .Global.Color( "Board.RouteBorder" ) = utilityObj.NewColorPattern( 153, 204, 255, 100, 0, False, True ) 
        .Global.Color( "Board.RouteFence.Hard" ) = utilityObj.NewColorPattern( 153, 204, 255, 100, 0, False, True ) 
        .Global.Color( "Board.RouteFence.Soft" ) = utilityObj.NewColorPattern( 153, 204, 255, 100, 0, False, True ) 
        'Unlock

        ' RF
            .Option( "Option.RFNodes.Enable" ) = epcbGraphicsItemStateOnEnabled 
            .Option( "Option.RFShapes.Enable" ) = epcbGraphicsItemStateOnEnabled
                .Option( "Option.RFObstructs.Pad.Enabled" ) = epcbGraphicsItemStateOnEnabled 
                .Option( "Option.RFObstructs.Plane.Enabled" ) = epcbGraphicsItemStateOnEnabled
                .Option( "Option.RFObstructs.Trace.Enabled" ) = epcbGraphicsItemStateOnEnabled 
                .Option( "Option.RFObstructs.TraceVia.Enabled" ) = epcbGraphicsItemStateOnEnabled 
                .Option( "Option.RFObstructs.Via.Enabled" ) = epcbGraphicsItemStateOnEnabled
            .Option( "Option.RFObstructs.Enabled" ) = epcbGraphicsItemStateOnEnabled 
        .Option( "Option.RFObjects.Enabled" ) = epcbGraphicsItemStateOnEnabled 

        ' Wirebond Objects
                    .Visible( "Fabrication.DiePins.Level0.Top" ) = epcbGraphicsItemStateOnEnabled 
                    .Visible( "Fabrication.DiePins.Level0.Bottom" ) = epcbGraphicsItemStateOnEnabled 
                    .Visible( "Fabrication.DiePins.Level1.Top" ) = epcbGraphicsItemStateOnEnabled 
                    .Visible( "Fabrication.DiePins.Level1.Bottom" ) = epcbGraphicsItemStateOnEnabled 
                    .Visible( "Fabrication.DiePins.Level2.Top" ) = epcbGraphicsItemStateOnEnabled 
                    .Visible( "Fabrication.DiePins.Level2.Bottom" ) = epcbGraphicsItemStateOnEnabled 
                    .Visible( "Fabrication.DiePins.Level3.Top" ) = epcbGraphicsItemStateOnEnabled 
                    .Visible( "Fabrication.DiePins.Level3.Bottom" ) = epcbGraphicsItemStateOnEnabled 
                    .Visible( "Fabrication.DiePins.Level4.Top" ) = epcbGraphicsItemStateOnEnabled 
                    .Visible( "Fabrication.DiePins.Level4.Bottom" ) = epcbGraphicsItemStateOnEnabled 
                .Option( "Option.DiePinsTop.Enabled" ) = epcbGraphicsItemStateOnEnabled 
                .Option( "Option.DiePinsBottom.Enabled" ) = epcbGraphicsItemStateOnEnabled 
                    .Visible( "Fabrication.BondWires.Level0.Top" ) = epcbGraphicsItemStateOnEnabled 
                    .Visible( "Fabrication.BondWires.Level0.Bottom" ) = epcbGraphicsItemStateOnEnabled 
                    .Visible( "Fabrication.BondWires.Level1.Top" ) = epcbGraphicsItemStateOnEnabled 
                    .Visible( "Fabrication.BondWires.Level1.Bottom" ) = epcbGraphicsItemStateOnEnabled 
                    .Visible( "Fabrication.BondWires.Level2.Top" ) = epcbGraphicsItemStateOnEnabled 
                    .Visible( "Fabrication.BondWires.Level2.Bottom" ) = epcbGraphicsItemStateOnEnabled 
                    .Visible( "Fabrication.BondWires.Level3.Top" ) = epcbGraphicsItemStateOnEnabled 
                    .Visible( "Fabrication.BondWires.Level3.Bottom" ) = epcbGraphicsItemStateOnEnabled 
                    .Visible( "Fabrication.BondWires.Level4.Top" ) = epcbGraphicsItemStateOnEnabled 
                    .Visible( "Fabrication.BondWires.Level4.Bottom" ) = epcbGraphicsItemStateOnEnabled 
                .Option( "Option.BondWiresTop.Enabled" ) = epcbGraphicsItemStateOnEnabled
                .Option( "Option.BondWiresBottom.Enabled" ) = epcbGraphicsItemStateOnEnabled 
                .Visible( "WirebondObjects.WirebondGuides.Top" ) = epcbGraphicsItemStateOnEnabled 
                .Visible( "WirebondObjects.WirebondGuides.Bottom" ) = epcbGraphicsItemStateOnEnabled
            .Option( "Option.WirebondItemsTop.Enabled" ) = epcbGraphicsItemStateOnEnabled 
            .Option( "Option.WirebondItemsBottom.Enabled" ) = epcbGraphicsItemStateOnEnabled 
        .Option( "Option.WirebondItems.Enabled" ) = epcbGraphicsItemStateOnEnabled 
        'Lock
        .Global.Color( "WirebondObjects.WirebondGuides.Top" ) = utilityObj.NewColorPattern( 255, 102, 0, 100, 0, False, True ) 
        .Global.Color( "WirebondObjects.WirebondGuides.Bottom" ) = utilityObj.NewColorPattern( 128, 0, 0, 100, 0, False, True ) 
        'Unlock

        ' Graphic options
        .Visible( "General.Color.SelectionShape" ) = epcbGraphicsItemStateOffEnabled 
        .Option( "Global.Option.Selection.DynamicHighlight" ) = epcbGraphicsItemStateOnEnabled 
            .StringOption( "Global.Option.DimMode" ) = "100" 
            .StringOption( "Global.Option.Transparency" ) = "80" 
        .Option( "Global.Option.Selection.DisplaySolid" ) = epcbGraphicsItemStateOffEnabled
        .Option( "Option.SelectionAndHighlights.EntireNetOnSelect" ) = epcbGraphicsItemStateOffEnabled 
        .Option( "Option.SelectionAndHighlights.DiffPairPinsOnSelect" ) = epcbGraphicsItemStateOffEnabled 
        .Option( "Option.SelectionAndHighlights.ElectricalNetOnSelect" ) = epcbGraphicsItemStateOffEnabled 
        .Option( "Option.NetlinesForSelectedItems.Enabled" ) = epcbGraphicsItemStateOnEnabled 
            .Option( "Option.ForceOutline" ) = epcbGraphicsItemStateOffEnabled 
            .Option( "Option.ForceSolid" ) = epcbGraphicsItemStateOffEnabled 
        .Option( "Option.FillPatterns" ) = epcbGraphicsItemStateOnEnabled 
            .Option( "Global.Option.HalfScreenCursor" ) = epcbGraphicsItemStateOffEnabled
            .Option( "Global.Option.FullScreenCursorDuringMoveOnly" ) = epcbGraphicsItemStateOffEnabled 
            .StringOption( "Global.Option.FullScreenCursorStyle" ) = "90Degree"
        .Option( "Global.Option.FullScreenCursor" ) = epcbGraphicsItemStateOnEnabled
            .StringOption( "Global.Option.PanSensitivity" ) = "7"
        .Option( "Global.Option.AutoPan" ) = epcbGraphicsItemStateOnEnabled 
        .Option( "Global.Option.PlaneDataBehindTraces" ) = epcbGraphicsItemStateOnEnabled
        .Option( "Global.Option.PlaneShapesOnTop" ) = epcbGraphicsItemStateOffEnabled 
        .Option( "Option.LegibleTextOnly" ) = epcbGraphicsItemStateOffEnabled 
        .Option( "Global.Option.NetNamesOnTraces" ) = epcbGraphicsItemStateOnEnabled 
        .Option( "Option.MirrorView" ) = epcbGraphicsItemStateOffEnabled 
        .Option( "Global.Option.TuningMeter" ) = epcbGraphicsItemStateOnEnabled 
        .Option( "Global.Option.LinkSketchNetlineDisplayToLayerVisibility" ) = epcbGraphicsItemStateOnEnabled 
        .Visible( "General.ActiveClearance" ) = epcbGraphicsItemStateOnEnabled 
        .StringOption( "Global.Option.ActiveClearanceRadius" ) = "100"
        '.Lock
        .Global.Color( "General.Color.SelectionShape" ) = utilityObj.NewColorPattern( 255, 255, 255, 100, 0, False, True ) 
        .Global.Color( "General.Color.Selection" ) = utilityObj.NewColorPattern( 255, 255, 0, 100, 0, False, True ) 
        .Global.Color( "General.Color.Highlight" ) = utilityObj.NewColorPattern( 255, 255, 255, 100, 0, False, True ) 
        .Global.Color( "General.Pattern.Fixed" ) = utilityObj.NewColorPattern( 255, 255, 255, 100, 34, False, True ) 
        .Global.Color( "General.Pattern.SemiFixed" ) = utilityObj.NewColorPattern( 255, 255, 255, 100, 6, False, True )
        .Global.Color( "General.Pattern.Locked" ) = utilityObj.NewColorPattern( 255, 255, 255, 100, 2, True, True )
        .Global.Color( "General.Color.Background" ) = utilityObj.NewColorPattern( 0, 0, 0, 100, 0, False, True ) 
        .Global.Color( "General.ActiveClearance" ) = utilityObj.NewColorPattern( 102, 153, 255, 50, 0, False, True ) 
        '.Unlock

        ' Grids
            .Visible( "Grid.Draw" ) = epcbGraphicsItemStateOnEnabled 
            .Visible( "Grid.Jumper" ) = epcbGraphicsItemStateOffEnabled 
            .Visible( "Grid.Part.Primary" ) = epcbGraphicsItemStateOffEnabled 
            .Visible( "Grid.Part.Secondary" ) = epcbGraphicsItemStateOffEnabled 
            .Visible( "Grid.Route" ) = epcbGraphicsItemStateOffEnabled 
            .Visible( "Grid.TestPoint" ) = epcbGraphicsItemStateOffEnabled 
            .Visible( "Grid.Via" ) = epcbGraphicsItemStateOffEnabled 
        .Option( "Option.Grids.Enabled" ) = epcbGraphicsItemStateOnEnabled 
        '.Lock
        .Global.Color( "Grid.Draw" ) = utilityObj.NewColorPattern( 255, 255, 255, 100, 0, False, True ) 
        .Global.Color( "Grid.Jumper" ) = utilityObj.NewColorPattern( 255, 255, 255, 100, 0, False, True ) 
        .Global.Color( "Grid.Part.Primary" ) = utilityObj.NewColorPattern( 255, 255, 255, 100, 0, False, True ) 
        .Global.Color( "Grid.Part.Secondary" ) = utilityObj.NewColorPattern( 255, 255, 255, 100, 0, False, True ) 
        .Global.Color( "Grid.Route" ) = utilityObj.NewColorPattern( 255, 255, 255, 100, 0, False, True ) 
        .Global.Color( "Grid.TestPoint" ) = utilityObj.NewColorPattern( 255, 255, 255, 100, 0, False, True ) 
        .Global.Color( "Grid.Via" ) = utilityObj.NewColorPattern( 255, 255, 255, 100, 0, False, True ) 
        '.Unlock

        ' Color by group
        .Option( "Global.Option.ColorByGroup.Enabled" ) = epcbGraphicsItemStateOnEnabled 

        ' Color by net or class
        .Option( "Global.Option.ColorByNetClass.PatternConstrainedNets" ) = epcbGraphicsItemStateOffEnabled
        .Option( "Global.Option.ColorByNetClass.PreserveLayerColorOnPlanes" ) = epcbGraphicsItemStateOffEnabled
        .Option( "Global.Option.ColorByNetClass.UseObjectColorAsPatternBackground" ) = epcbGraphicsItemStateOffEnabled
        .Option( "Global.Option.ColorByNetClass.UseCMColors" ) = epcbGraphicsItemStateOffEnabled 
        .Option( "Global.Option.ColorByNetClass.Netlines" ) = epcbGraphicsItemStateOnEnabled 
        .Option( "Global.Option.ColorByNetClass.Traces" ) = epcbGraphicsItemStateOnEnabled 
        .Option( "Global.Option.ColorByNetClass.Pads" ) = epcbGraphicsItemStateOnEnabled 
        .Option( "Global.Option.ColorByNetClass.Planes" ) = epcbGraphicsItemStateOnEnabled 
        .Option( "Global.Option.ColorByNetClass.Vias" ) = epcbGraphicsItemStateOnEnabled 
        .Option( "Global.Option.ColorByNetClass.ConductiveShapes" ) = epcbGraphicsItemStateOnEnabled 
            '.Visible( "[Net].GND" ) = epcbGraphicsItemStateOnEnabled 
        .Option( "Global.Option.ColorByNetClass.Nets.Enabled" ) = epcbGraphicsItemStateOnEnabled 
        .Option( "Global.Option.ColorByNetClass.NetClasses.Enabled" ) = epcbGraphicsItemStateOnEnabled 
        .Option( "Global.Option.ColorByNetClass.ConstraintClasses.Enabled" ) = epcbGraphicsItemStateOffEnabled

        ' Object appearance
        ' .Lock
        ' .Global.Color( "Copper.Trace.*" ) = utilityObj.NewColorPattern( 255, 255, 255, 100, 0, False, True )  
        ' .Global.Color( "Copper.Pad.*" ) = utilityObj.NewColorPattern( 255, 255, 255, 100, 0, False, True ) 
        ' .Global.Color( "Copper.Plane.Data.*" ) = utilityObj.NewColorPattern( 255, 255, 255, 100, 0, False, True ) 
        ' .Unlock

        ' Board Objects
                .Visible( "Copper.Pad.Fiducial.Top" ) = epcbGraphicsItemStateOnEnabled 
                .Visible( "Copper.Pad.Fiducial.Bottom" ) = epcbGraphicsItemStateOnEnabled 
            .Option( "Option.Fiducials.Enabled" ) = epcbGraphicsItemStateOnEnabled 
                .Option( "Option.MountingHolePads.Enabled" ) = epcbGraphicsItemStateOnEnabled
            .Option( "Option.Holes.Enabled" ) = epcbGraphicsItemStateOnEnabled 
            .Visible( "Fabrication.Hole.Mounting" ) = epcbGraphicsItemStateOnEnabled 
                .Visible( "Board.BoardOutline" ) = epcbGraphicsItemStateOnEnabled 
                .Visible( "Board.ManufacturingOutline" ) = epcbGraphicsItemStateOnEnabled 
                .Visible( "Board.FixtureOutline" ) = epcbGraphicsItemStateOnEnabled 
                .Visible( "Board.Cavity" ) = epcbGraphicsItemStateOnEnabled 
                .Visible( "Fabrication.Hole.Contour" ) = epcbGraphicsItemStateOnEnabled 
                .Visible( "Fabrication.Hole.Contour.SpanNumbers" ) = epcbGraphicsItemStateOnEnabled 
                .Visible( "Board.DRCWindow" ) = epcbGraphicsItemStateOnEnabled
                .Visible( "Board.Origin.Board" ) = epcbGraphicsItemStateOffEnabled
                .Visible( "Board.Origin.NCDrill" ) = epcbGraphicsItemStateOffEnabled 
                .Visible( "Fabrication.RedlineLayer" ) = epcbGraphicsItemStateOffEnabled 
            .Option( "Option.BoardElements.Enabled" ) = epcbGraphicsItemStateOnEnabled 
                .Option( "Option.TextItems.PinProperties.Enabled" ) = epcbGraphicsItemStateOnEnabled 
                .Option( "Option.TextItems.CellProperties.Enabled" ) = epcbGraphicsItemStateOnEnabled
            .Option( "Option.TextItems.Enabled" ) = epcbGraphicsItemStateOnEnabled 
        .Option( "Option.BoardObjects.Enabled" ) = epcbGraphicsItemStateOnEnabled 
        '.Lock
        .Global.Color( "Copper.Pad.Fiducial.Top" ) = utilityObj.NewColorPattern( 0, 255, 255, 100, 0, False, True ) 
        .Global.Color( "Copper.Pad.Fiducial.Bottom" ) = utilityObj.NewColorPattern( 255, 204, 255, 100, 0, False, True ) 
        .Global.Color( "Fabrication.Hole.Mounting" ) = utilityObj.NewColorPattern( 204, 153, 0, 100, 0, False, True ) 
        .Global.Color( "Board.BoardOutline" ) = utilityObj.NewColorPattern( 255, 0, 0, 100, 0, False, True ) 
        .Global.Color( "Board.ManufacturingOutline" ) = utilityObj.NewColorPattern( 102, 0, 102, 100, 0, False, True ) 
        .Global.Color( "Board.FixtureOutline" ) = utilityObj.NewColorPattern( 0, 0, 153, 100, 0, False, True ) 
        .Global.Color( "Board.Cavity" ) = utilityObj.NewColorPattern( 102, 0, 102, 100, 0, False, True ) 
        .Global.Color( "Fabrication.Hole.Contour" ) = utilityObj.NewColorPattern( 204, 153, 0, 100, 0, False, True ) 
        .Global.Color( "Board.DRCWindow" ) = utilityObj.NewColorPattern( 153, 102, 255, 100, 0, False, True ) 
        .Global.Color( "Board.Origin.Board" ) = utilityObj.NewColorPattern( 255, 255, 255, 100, 0, False, True ) 
        .Global.Color( "Board.Origin.NCDrill" ) = utilityObj.NewColorPattern( 255, 255, 255, 100, 0, False, True ) 
        .Global.Color( "Fabrication.RedlineLayer" ) = utilityObj.NewColorPattern( 255, 0, 0, 100, 0, False, True )
        '.Unlock

        ' Fabrication objects	
        .Visible( "Fabrication.SolderMask.Bottom" ) = epcbGraphicsItemStateOffEnabled 
		.Visible( "Fabrication.SolderMask.Top" ) = epcbGraphicsItemStateOffEnabled 
		.Visible( "Fabrication.SolderPaste.Bottom" ) = epcbGraphicsItemStateOffEnabled 
		.Visible( "Fabrication.SolderPaste.Top" ) = epcbGraphicsItemStateOffEnabled 	
            .Visible( "Fabrication.Assembly.Part.Outline.Bottom" ) = epcbGraphicsItemStateOnEnabled 
            .Visible( "Fabrication.Assembly.Part.Outline.Top" ) = epcbGraphicsItemStateOnEnabled 
            .Visible( "Fabrication.Assembly.Part.Text.PartNumber.Bottom" ) = epcbGraphicsItemStateOffEnabled 
            .Visible( "Fabrication.Assembly.Part.Text.PartNumber.Top" ) = epcbGraphicsItemStateOffEnabled
            .Visible( "Fabrication.Assembly.Part.Text.RefDes.Bottom" ) = epcbGraphicsItemStateOffEnabled 
            .Visible( "Fabrication.Assembly.Part.Text.RefDes.Top" ) = epcbGraphicsItemStateOffEnabled
        .Option( "Option.Fabrication.AssemblyItems.Top" ) = epcbGraphicsItemStateOnEnabled 
        .Option( "Option.Fabrication.AssemblyItems.Bottom" ) = epcbGraphicsItemStateOnEnabled  
            .Visible( "Fabrication.Silkscreen.Part.Outline.Bottom" ) = epcbGraphicsItemStateOnEnabled 
            .Visible( "Fabrication.Silkscreen.Part.Outline.Top" ) = epcbGraphicsItemStateOnEnabled 
            .Visible( "Fabrication.Silkscreen.Part.Text.PartNumber.Bottom" ) = epcbGraphicsItemStateOffEnabled 
            .Visible( "Fabrication.Silkscreen.Part.Text.PartNumber.Top" ) = epcbGraphicsItemStateOffEnabled 
            .Visible( "Fabrication.Silkscreen.Part.Text.RefDes.Top" ) = epcbGraphicsItemStateOffEnabled 
            .Visible( "Fabrication.Silkscreen.Part.Text.RefDes.Bottom" ) = epcbGraphicsItemStateOffEnabled 
            .Visible( "Fabrication.Silkscreen.Generated.Bottom" ) = epcbGraphicsItemStateOnEnabled 
            .Visible( "Fabrication.Silkscreen.Generated.Top" ) = epcbGraphicsItemStateOnEnabled 
        .Option( "Option.Fabrication.SilkscreenItems.Top" ) = epcbGraphicsItemStateOnEnabled 
        .Option( "Option.Fabrication.SilkscreenItems.Bottom" ) = epcbGraphicsItemStateOnEnabled 
            .Visible( "Fabrication.Assembly.TestPoint.Text.RefDes.Bottom" ) = epcbGraphicsItemStateOnEnabled 
            .Visible( "Fabrication.Assembly.TestPoint.Text.RefDes.Top" ) = epcbGraphicsItemStateOnEnabled 
            .Visible( "Fabrication.Silkscreen.TestPoint.Text.RefDes.Bottom" ) = epcbGraphicsItemStateOnEnabled 
            .Visible( "Fabrication.Silkscreen.TestPoint.Text.RefDes.Top" ) = epcbGraphicsItemStateOnEnabled 
            .Visible( "Fabrication.Silkscreen.TestPoint.Probe.Bottom" ) = epcbGraphicsItemStateOnEnabled 
            .Visible( "Fabrication.Silkscreen.TestPoint.Probe.Top" ) = epcbGraphicsItemStateOnEnabled 
        .Option( "Option.Fabrication.TestPointItems.Top" ) = epcbGraphicsItemStateOffEnabled 
        .Option( "Option.Fabrication.TestPointItems.Bottom" ) = epcbGraphicsItemStateOffEnabled 
        ' .Lock
        .Global.Color( "Fabrication.Soldermask.Top" ) = utilityObj.NewColorPattern( 51, 51, 204, 100, 0, False, True ) 
        .Global.Color( "Fabrication.Soldermask.Bottom" ) = utilityObj.NewColorPattern( 204, 0, 0, 100, 0, False, True ) 
        .Global.Color( "Fabrication.Solderpaste.Top" ) = utilityObj.NewColorPattern( 204, 153, 0, 100, 0, False, True )
        .Global.Color( "Fabrication.Solderpaste.Bottom" ) = utilityObj.NewColorPattern( 102, 0, 204, 60, 0, False, True ) 
        .Global.Color( "Fabrication.Assembly.Part.Outline.Top" ) = utilityObj.NewColorPattern( 51, 204, 51, 100, 0, False, True ) 
        .Global.Color( "Fabrication.Assembly.Part.Text.PartNumber.Top" ) = utilityObj.NewColorPattern( 51, 204, 51, 100, 0, False, True ) 
        .Global.Color( "Fabrication.Assembly.Part.Text.RefDes.Top" ) = utilityObj.NewColorPattern( 51, 204, 51, 100, 0, False, True ) 
        .Global.Color( "Fabrication.Assembly.Part.Outline.Bottom" ) = utilityObj.NewColorPattern( 51, 153, 255, 100, 0, False, True ) 
        .Global.Color( "Fabrication.Assembly.Part.Text.PartNumber.Bottom" ) = utilityObj.NewColorPattern( 51, 153, 255, 100, 0, False, True ) 
        .Global.Color( "Fabrication.Assembly.Part.Text.RefDes.Bottom" ) = utilityObj.NewColorPattern( 51, 153, 255, 100, 0, False, True ) 
        .Global.Color( "Fabrication.Silkscreen.Part.Outline.Top" ) = utilityObj.NewColorPattern( 51, 153, 255, 100, 0, False, True ) 
        .Global.Color( "Fabrication.Silkscreen.Part.Text.PartNumber.Top" ) = utilityObj.NewColorPattern( 51, 153, 255, 100, 0, False, True ) 
        .Global.Color( "Fabrication.Silkscreen.Part.Text.RefDes.Top" ) = utilityObj.NewColorPattern( 51, 153, 255, 100, 0, False, True ) 
        .Global.Color( "Fabrication.Silkscreen.Generated.Top" ) = utilityObj.NewColorPattern( 51, 204, 255, 100, 0, False, True ) 
        .Global.Color( "Fabrication.Silkscreen.Part.Outline.Bottom" ) = utilityObj.NewColorPattern( 255, 153, 255, 100, 0, False, True ) 
        .Global.Color( "Fabrication.Silkscreen.Part.Text.PartNumber.Bottom" ) = utilityObj.NewColorPattern( 255, 153, 255, 100, 0, False, True ) 
        .Global.Color( "Fabrication.Silkscreen.Part.Text.RefDes.Bottom" ) = utilityObj.NewColorPattern( 255, 153, 255, 100, 0, False, True ) 
        .Global.Color( "Fabrication.Silkscreen.Generated.Bottom" ) = utilityObj.NewColorPattern( 255, 204, 255, 100, 0, False, True )
        .Global.Color( "Part.Cell.GlueSpot.Top" ) = utilityObj.NewColorPattern( 255, 255, 255, 100, 0, False, True ) 
        .Global.Color( "Part.Cell.GlueSpot.Bottom" ) = utilityObj.NewColorPattern( 255, 255, 255, 100, 0, False, True ) 
        .Global.Color( "Part.Cell.Origin.Top" ) = utilityObj.NewColorPattern( 255, 0, 102, 100, 0, False, True ) 
        .Global.Color( "Part.Cell.Origin.Bottom" ) = utilityObj.NewColorPattern( 255, 0, 102, 100, 0, False, True ) 
        .Global.Color( "Fabrication.Assembly.TestPoint.Text.RefDes.Top" ) = utilityObj.NewColorPattern( 255, 255, 255, 100, 0, False, True ) 
        .Global.Color( "Fabrication.Silkscreen.TestPoint.Text.RefDes.Top" ) = utilityObj.NewColorPattern( 255, 255, 255, 100, 0, False, True ) 
        .Global.Color( "Fabrication.Assembly.TestPoint.Text.RefDes.Bottom" ) = utilityObj.NewColorPattern( 255, 255, 255, 100, 0, False, True ) 
        .Global.Color( "Fabrication.Silkscreen.TestPoint.Text.RefDes.Bottom" ) = utilityObj.NewColorPattern( 255, 255, 255, 100, 0, False, True ) 
        .Global.Color( "Fabrication.Silkscreen.TestPoint.Probe.Top" ) = utilityObj.NewColorPattern( 255, 255, 0, 100, 0, False, True ) 
        .Global.Color( "Fabrication.Silkscreen.TestPoint.Probe.Bottom" ) = utilityObj.NewColorPattern( 0, 255, 0, 100, 0, False, True ) 
        ' .Unlock
		
		' Cell layer graphics
            .Visible( "Part.Cell.GlueSpot.Top" ) = epcbGraphicsItemStateOnEnabled 
            .Visible( "Part.Cell.GlueSpot.Bottom" ) = epcbGraphicsItemStateOnEnabled 
            .Visible( "Part.Cell.Origin.Top" ) = epcbGraphicsItemStateOnEnabled 
            .Visible( "Part.Cell.Origin.Bottom" ) = epcbGraphicsItemStateOnEnabled 
        .Option( "Option.Fabrication.CellItems.Top" ) = epcbGraphicsItemStateOffEnabled 
        .Option( "Option.Fabrication.CellItems.Bottom" ) = epcbGraphicsItemStateOffEnabled 
		
		' Copper Balancing
            .Option( "Option.CopperBalancing.Data" ) = epcbGraphicsItemStateOnEnabled 
            .Option( "Option.CopperBalancing.Shapes" ) = epcbGraphicsItemStateOnEnabled
        .Option( "Option.CopperBalancing.Enabled" ) = epcbGraphicsItemStateOffEnabled 

        ' Materials
        .Option( "Option.Materials.Enabled" ) = epcbGraphicsItemStateOffEnabled
        
        ' Drill Drawing
		.Option( "Option.Fabrication.DrillDrawing" ) = epcbGraphicsItemStateOffEnabled 

        ' User draft layers
        .Option( "Option.UserDraftLayers.Enabled" ) = epcbGraphicsItemStateOffEnabled 

        ' Detail View
		.Visible( "Fabrication.DetailViews" ) = epcbGraphicsItemStateOnEnabled 

        ' Hazards
		'.Option( "Global.Option.DRC.ColorByHazard.Enabled" ) = epcbGraphicsItemStateOnEnabled 
		'.Option( "Global.Option.DRC.ColorByHazard.ShadowMode" ) = epcbGraphicsItemStateOnEnabled 

		' Online Hazards
		'.Visible( "Hazard.Online.Component" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Online.MissingArc" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Online.OpenFanouts" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Online.NetOpenCounts" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Online.TraceWidth" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Online.PadEntry" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Online.NodeEntry" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Online.LayerConstraints" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Online.ViasPerNet" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Online.ViaUsage" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Online.Teardrop" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Online.MultiViaPadEntry" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Online.ViaGrid" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Online.RouteGrid" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Online.JumperGrid" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Online.TestPointGrid" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Online.PlacementGrid" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Online.PlacementRotation" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Online.Stub" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Online.TJunction" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Online.TopologyBalancing" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Online.TopologyMismatch" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Online.MaxLength" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Online.GroupMatchLength" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Online.MaxDelay" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Online.GroupMatchDelay" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Online.Formulas" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Online.Parallelism" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Online.DpConvergence" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Online.DpClearance" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Online.DpDelayTolerance" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Online.DpLengthTolerance" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Online.DpConvergenceTol" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Online.DpPhaseMatching" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Online.CrossTalk" ) = epcbGraphicsItemStateOnEnabled 

		' Batch Hazards
		'.Visible( "Hazard.Batch.Proximity" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Batch.MissingPastePads" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Batch.Hangers" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Batch.TraceLoops" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Batch.TraceWidths" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Batch.PlaneIslands" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Batch.EPViolation" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Batch.SinglePointNets" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Batch.PartialNets" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Batch.Dangling" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Batch.UnplatedConnectorPin" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Batch.ViasUnderParts" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Batch.ViasUnderSMDPads" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Batch.MinAnnularRing" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Batch.MissingCondPads" ) = epcbGraphicsItemStateOnEnabled 
		'.Visible( "Hazard.Batch.MissingMaskPads" ) = epcbGraphicsItemStateOnEnabled 
    End With

    ' ' Save the display scheme
    ' displayCtrlObj.SaveSchemeEx("Sys:MyRouting")

    ' Refresh view
    ' pcbGuiObj.ProcessCommand("View->Fit All")
    pcbGuiObj.ProcessCommand("View->Previous View")
    pcbGuiObj.ProcessCommand("View->Next View")
    ' Application.Addins("iDCAddinCtrl").Visible = False
    ' Application.Addins("iDCAddinCtrl").Visible = True

    pcbAppObj.UnlockServer          
	pcbDocObj.TransactionEnd        
	pcbAppObj.Gui.CursorBusy(False)

End If

' Server validation function
Private Function ValidateServer(doc)
    
    dim key, licenseServer, licenseToken

    ' Ask Expedition's document for the key
    key = doc.Validate(0)

    ' Get license server
    Set licenseServer = CreateObject("MGCPCBAutomationLicensing.Application")

    ' Ask the license server for the license token
    licenseToken = licenseServer.GetToken(key)

    ' Release license server
    Set licenseServer = nothing

    ' Turn off error messages.  Validate may fail if the token is incorrect
    On Error Resume Next
    Err.Clear

    ' Ask the document to validate the license token
    doc.Validate(licenseToken)
    If Err Then
        ValidateServer = 0    
    Else
        ValidateServer = 1
    End If

End Function
