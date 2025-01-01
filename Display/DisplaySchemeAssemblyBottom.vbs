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
    For i = 1 To conductorLayerCount
		displayCtrlObj.Visible( "LayerControl." & i ) = epcbGraphicsItemStateOffEnabled	
	Next
    With displayCtrlObj
        .Visible( "Copper.Trace." & conductorLayerCount ) = epcbGraphicsItemStateOnEnabled
        .Visible( "Copper.Pad." & conductorLayerCount ) = epcbGraphicsItemStateOnEnabled
        .Visible( "Copper.Plane.Data." & conductorLayerCount ) = epcbGraphicsItemStateOnEnabled
        .Global.Color( "Copper.Trace." & conductorLayerCount ) = utilityObj.NewColorPattern( 255, 0, 0, 100, 0, False, True ) 
        .Global.Color( "Copper.Pad." & conductorLayerCount ) = utilityObj.NewColorPattern( 192, 192, 192, 100, 0, False, True ) 
        .Global.Color( "Copper.Plane.Data." & conductorLayerCount ) = utilityObj.NewColorPattern( 255, 0, 0, 100, 0, False, True ) 
	    .Visible( "LayerControl." & conductorLayerCount  ) = epcbGraphicsItemStateOnEnabled
        .Option( "Option.ActiveLayerOnly" ) = epcbGraphicsItemStateOffEnabled
    End With

    '--------------------------------------------------------------------------------
	' All other layers display control
	'--------------------------------------------------------------------------------
	With displayCtrlObj
        ' Global View & Interactive Selection
        .Option( "Option.Planning.Enabled" ) = epcbGraphicsItemStateOffEnabled 
                .Visible( "Group.Outline.Top" ) = epcbGraphicsItemStateOffEnabled 
                .Visible( "Group.Outline.Bubble.Top" ) = epcbGraphicsItemStateOffEnabled 
                .Visible( "Part.PlaceOutline.Top" ) = epcbGraphicsItemStateOffEnabled 
                .Option( "Option.PlaceObjects.ObstructsAndRooms.Top" ) = epcbGraphicsItemStateOffEnabled 
            .Option( "Option.PlaceObjects.Parts.Top" ) = epcbGraphicsItemStateOffEnabled 
                .Visible( "Group.Outline.Bottom" ) = epcbGraphicsItemStateOffEnabled 
                .Visible( "Group.Outline.Bubble.Bottom" ) = epcbGraphicsItemStateOffEnabled 
                .Visible( "Part.PlaceOutline.Bottom" ) = epcbGraphicsItemStateOffEnabled 
                .Option( "Option.PlaceObjects.ObstructsAndRooms.Bottom" ) = epcbGraphicsItemStateOffEnabled 
            .Option( "Option.PlaceObjects.Parts.Bottom" ) = epcbGraphicsItemStateOnEnabled 
        .Option( "Option.PlaceObjects" ) = epcbGraphicsItemStateOffEnabled 
            .Option( "Option.Traces.Enabled" ) = epcbGraphicsItemStateOffEnabled 
            .Option( "Option.Vias.Enabled" ) = epcbGraphicsItemStateOffEnabled 
            .Option( "Option.Pins.Enabled" ) = epcbGraphicsItemStateOnEnabled 
            .Option( "Option.Netlines.Enabled" ) = epcbGraphicsItemStateOffEnabled 
            .Option( "Option.Planes.Enabled" ) = epcbGraphicsItemStateOffEnabled 
            .Option( "Option.RouteObstructs.Enabled" ) = epcbGraphicsItemStateOffEnabled 
            .Option( "Option.RouteAreas.Enabled" ) = epcbGraphicsItemStateOffEnabled 
            .Option( "Option.ConductiveShapes.Enabled" ) = epcbGraphicsItemStateOffEnabled 
            .Option( "Option.Teardrops.Enabled" ) = epcbGraphicsItemStateOffEnabled 
        .Option( "Option.RouteObjects.Enabled" ) = epcbGraphicsItemStateOnEnabled 
            .Option( "Option.RFNodes.Enabled" ) = epcbGraphicsItemStateOffEnabled 
            .Option( "Option.RFShapes.Enabled" ) = epcbGraphicsItemStateOffEnabled 
            .Option( "Option.RFObstructs.Enabled" ) = epcbGraphicsItemStateOffEnabled 
        .Option( "Option.RFObjects.Enabled" ) = epcbGraphicsItemStateOffEnabled 
                .Option( "Option.DiePinsTop.Enabled" ) = epcbGraphicsItemStateOffEnabled 
                .Option( "Option.BondWiresTop.Enabled" ) = epcbGraphicsItemStateOffEnabled 
                .Visible( "WirebondObjects.WirebondGuides.Top" ) = epcbGraphicsItemStateOffEnabled 
            .Option( "Option.WirebondItemsTop.Enabled" ) = epcbGraphicsItemStateOffEnabled 
                .Option( "Option.DiePinsBottom.Enabled" ) = epcbGraphicsItemStateOffEnabled 
                .Option( "Option.BondWiresBottom.Enabled" ) = epcbGraphicsItemStateOffEnabled 
                .Visible( "WirebondObjects.WirebondGuides.Bottom" ) = epcbGraphicsItemStateOffEnabled 
            .Option( "Option.WirebondItemsBottom.Enabled" ) = epcbGraphicsItemStateOffEnabled 
        .Option( "Option.WirebondItems.Enabled" ) = epcbGraphicsItemStateOffEnabled 
            .Option( "Option.Fiducials.Enabled" ) = epcbGraphicsItemStateOnEnabled 
            .Option( "Option.Holes.Enabled" ) = epcbGraphicsItemStateOnEnabled
            .Option( "Option.BoardElements.Enabled" ) = epcbGraphicsItemStateOnEnabled 
        .Option( "Option.BoardObjects.Enabled" ) = epcbGraphicsItemStateOnEnabled 
            .Option( "Option.FabricationObjects" ) = epcbGraphicsItemStateOnEnabled 
            .Option( "Option.CopperBalancing.Enabled" ) = epcbGraphicsItemStateOffEnabled 
            .Option( "Option.Materials.Enabled" ) = epcbGraphicsItemStateOffEnabled 
            .Option( "Option.Fabrication.DrillDrawing" ) = epcbGraphicsItemStateOffEnabled 
            .Option( "Option.UserDraftLayers.Enabled" ) = epcbGraphicsItemStateOffEnabled 
        .Option( "Option.DrawFabObjects.Enabled" ) = epcbGraphicsItemStateOnEnabled 
        .Visible( "Fabrication.DetailViews" ) = epcbGraphicsItemStateOffEnabled 

		' Route/Multi Planning
                .Option( "Option.VirtualPins.Enabled" ) = epcbGraphicsItemStateOffEnabled
                .Option( "Option.DiffPairCenterlines.Enabled" ) = epcbGraphicsItemStateOffEnabled
                .Option( "Option.UnpackedAreas.Enabled" ) = epcbGraphicsItemStateOffEnabled 
                .Option( "Option.BusPaths.Enabled" ) = epcbGraphicsItemStateOffEnabled 
                .Option( "Option.TargetAreas.Enabled" ) = epcbGraphicsItemStateOffEnabled 
                .Option( "Option.RouteTargets.Enabled" ) = epcbGraphicsItemStateOffEnabled 
            .Option( "Option.RoutePlanning" ) = epcbGraphicsItemStateOffEnabled

                .Option( "Option.SketchPlans.Width" ) = epcbGraphicsItemStateOffEnabled
            .Option( "Option.SketchPlans.Enabled" ) = epcbGraphicsItemStateOffEnabled
            .Option( "Option.ReuseArea.Enabled" ) = epcbGraphicsItemStateOffEnabled
                    .Option( "Option.MultipleDesigners.TeamPCB.ShadowMode" ) = epcbGraphicsItemStateOffEnabled
				.Visible( "General.MultipleDesigners.TeamPCB.ReservedAreas" ) = epcbGraphicsItemStateOffEnabled 
				.Visible( "General.MultipleDesigners.Xtreme.ProtectedAreas" ) = epcbGraphicsItemStateOffEnabled 
				.Visible( "Board.Sandbox" ) = epcbGraphicsItemStateOffEnabled
			.Option( "Option.MultipleDesigners" ) = epcbGraphicsItemStateOffEnabled
		.Option( "Option.Planning.Enabled" ) = epcbGraphicsItemStateOffEnabled

        ' Place
			.Visible( "Group.Outline.Top" ) = epcbGraphicsItemStateOffEnabled 
			.Visible( "Group.Outline.Bottom" ) = epcbGraphicsItemStateOffEnabled
			.Visible( "Group.Outline.Bubble.Top" ) = epcbGraphicsItemStateOffEnabled 
			.Visible( "Group.Outline.Bubble.Bottom" ) = epcbGraphicsItemStateOffEnabled 
			.Visible( "Place.Part.Text.RefDes.Top" ) = epcbGraphicsItemStateOffEnabled 
			.Visible( "Place.Part.Text.RefDes.Bottom" ) = epcbGraphicsItemStateOnEnabled
			
                .Option( "Option.SelectableInsidePartOutline" ) = epcbGraphicsItemStateOffEnabled
                .Option( "Option.FillPartOutlineOnSelection" ) = epcbGraphicsItemStateOffEnabled
            .Visible( "Part.PlaceOutline.Top" ) = epcbGraphicsItemStateOffEnabled 
			.Visible( "Part.PlaceOutline.Bottom" ) = epcbGraphicsItemStateOffEnabled
			
				.Visible( "Board.Obstruct.Part.Top" ) = epcbGraphicsItemStateOffEnabled 
				.Visible( "Board.Obstruct.Part.Bottom" ) = epcbGraphicsItemStateOffEnabled 
				.Visible( "Board.Obstruct.TestPoint.Top" ) = epcbGraphicsItemStateOffEnabled 
				.Visible( "Board.Obstruct.TestPoint.Bottom" ) = epcbGraphicsItemStateOffEnabled 
				.Visible( "Board.Room.Top" ) = epcbGraphicsItemStateOffEnabled 
				.Visible( "Board.Room.Bottom" ) = epcbGraphicsItemStateOffEnabled 	
			.Option( "Option.PlaceObjects.ObstructsAndRooms.Top" ) = epcbGraphicsItemStateOffEnabled 
			.Option( "Option.PlaceObjects.ObstructsAndRooms.Bottom" ) = epcbGraphicsItemStateOffEnabled 
		
				.Visible( "Part.InsertionOutline.Top" ) = epcbGraphicsItemStateOffEnabled 
				.Visible( "Part.InsertionOutline.Bottom" ) = epcbGraphicsItemStateOffEnabled 
				.Visible( "Part.Hazard.Top" ) = epcbGraphicsItemStateOffEnabled 
				.Visible( "Part.Hazard.Bottom" ) = epcbGraphicsItemStateOffEnabled 
				.Option( "Option.Pin.Number.Top" ) = epcbGraphicsItemStateOnEnabled 
				.Option( "Option.Pin.Number.Bottom" ) = epcbGraphicsItemStateOnEnabled 
				.Option( "Option.Pin.Type.Top" ) = epcbGraphicsItemStateOffEnabled 
				.Option( "Option.Pin.Type.Bottom" ) = epcbGraphicsItemStateOffEnabled 
                .Option( "Option.Pin.NetName.Top" ) = epcbGraphicsItemStateOffEnabled
                .Option( "Option.Pin.NetName.Bottom" ) = epcbGraphicsItemStateOffEnabled 
			.Option( "Option.PlaceObjects.PartItems.Top" ) = epcbGraphicsItemStateOffEnabled 
			.Option( "Option.PlaceObjects.PartItems.Bottom" ) = epcbGraphicsItemStateOnEnabled 
        .Option( "Option.PlaceObjects.Parts.Top" ) = epcbGraphicsItemStateOffEnabled 
        .Option( "Option.PlaceObjects.Parts.Bottom" ) = epcbGraphicsItemStateOnEnabled 
		.Option( "Option.PlaceObjects" ) = epcbGraphicsItemStateOnEnabled
        'Lock
        .Global.Color( "Part.Pin.NumberType.Bottom" ) = utilityObj.NewColorPattern( 255, 0, 153, 100, 0, False, True ) 
        'Unlock

        ' Vias
			.Option( "Option.Pad.Via.AllSameColor" ) = epcbGraphicsItemStateOffEnabled 
			.Option( "Option.ViaPads.Enabled" ) = epcbGraphicsItemStateOnEnabled 
			.Visible( "Fabrication.Hole.Via" ) = epcbGraphicsItemStateOnEnabled 
			.Visible( "General.Via.SpanNumbers" ) = epcbGraphicsItemStateOnEnabled 
            .Option( "Option.Vias.NetNames" ) = epcbGraphicsItemStateOnEnabled 
			.Visible( "General.Via.InactiveBlindBuriedPad" ) = epcbGraphicsItemStateOffEnabled 
		.Option( "Option.Vias.Enabled" ) = epcbGraphicsItemStateOffEnabled

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
        .Global.Color( "Fabrication.Hole.Pin" ) = utilityObj.NewColorPattern( 0, 0, 153, 100, 0, False, True )  
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

        ' Planes
			.Option( "Option.Planes.Data.Fill" ) = epcbGraphicsItemStateOnEnabled 
			.Option( "Option.Planes.Data.Enabled" ) = epcbGraphicsItemStateOnEnabled 
			.Option( "Option.Planes.Shape.Enabled" ) = epcbGraphicsItemStateOnEnabled 
			.Option( "Option.Planes.Sketch.Enabled" ) = epcbGraphicsItemStateOnEnabled 
		.Option( "Option.Planes.Enabled" ) = epcbGraphicsItemStateOffEnabled

        ' Route Obstructs
			.Option( "Option.RouteObstructs.Pad.Enabled" ) = epcbGraphicsItemStateOnEnabled 
			.Option( "Option.RouteObstructs.Plane.Enabled" ) = epcbGraphicsItemStateOnEnabled 
			.Option( "Option.RouteObstructs.Trace.Enabled" ) = epcbGraphicsItemStateOnEnabled 
			.Option( "Option.RouteObstructs.TraceVia.Enabled" ) = epcbGraphicsItemStateOnEnabled 
			.Option( "Option.RouteObstructs.Via.Enabled" ) = epcbGraphicsItemStateOnEnabled 
			.Option( "Option.RouteObstructs.TuningPattern.Enabled" ) = epcbGraphicsItemStateOnEnabled 
			    .Option( "Option.Spacers.ShadowModeEnabled" ) = epcbGraphicsItemStateOffEnabled 
			.Option( "Option.Spacers.Enabled" ) = epcbGraphicsItemStateOnEnabled 
		.Option( "Option.RouteObstructs.Enabled" ) = epcbGraphicsItemStateOffEnabled
        
        ' Route Areas
			.Visible( "Board.RouteBorder" ) = epcbGraphicsItemStateOnEnabled 
			.Visible( "Board.RouteFence.Hard" ) = epcbGraphicsItemStateOnEnabled 
			.Visible( "Board.RouteFence.Soft" ) = epcbGraphicsItemStateOnEnabled 
			.Option( "Option.RuleAreas.Enabled" ) = epcbGraphicsItemStateOnEnabled 
		.Option( "Option.RouteAreas.Enabled" ) = epcbGraphicsItemStateOffEnabled 

        ' RF
            .Option( "Option.RFNodes.Enable" ) = epcbGraphicsItemStateOnEnabled 
            .Option( "Option.RFShapes.Enable" ) = epcbGraphicsItemStateOnEnabled
                .Option( "Option.RFObstructs.Pad.Enabled" ) = epcbGraphicsItemStateOnEnabled 
                .Option( "Option.RFObstructs.Plane.Enabled" ) = epcbGraphicsItemStateOnEnabled
                .Option( "Option.RFObstructs.Trace.Enabled" ) = epcbGraphicsItemStateOnEnabled 
                .Option( "Option.RFObstructs.TraceVia.Enabled" ) = epcbGraphicsItemStateOnEnabled 
                .Option( "Option.RFObstructs.Via.Enabled" ) = epcbGraphicsItemStateOnEnabled
            .Option( "Option.RFObstructs.Enabled" ) = epcbGraphicsItemStateOnEnabled 
        .Option( "Option.RFObjects.Enabled" ) = epcbGraphicsItemStateOffEnabled 

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
        .Option( "Option.WirebondItems.Enabled" ) = epcbGraphicsItemStateOffEnabled 

        ' Graphic options
        .Visible( "General.Color.SelectionShape" ) = epcbGraphicsItemStateOffEnabled 
        .Option( "Global.Option.Selection.DynamicHighlight" ) = epcbGraphicsItemStateOnEnabled 
            .StringOption( "Global.Option.DimMode" ) = "100" 
            .StringOption( "Global.Option.Transparency" ) = "80" 
        .Option( "Global.Option.Selection.DisplaySolid" ) = epcbGraphicsItemStateOffEnabled
        .Option( "Option.SelectionAndHighlights.EntireNetOnSelect" ) = epcbGraphicsItemStateOffEnabled 
        .Option( "Option.SelectionAndHighlights.DiffPairPinsOnSelect" ) = epcbGraphicsItemStateOffEnabled 
        .Option( "Option.SelectionAndHighlights.ElectricalNetOnSelect" ) = epcbGraphicsItemStateOffEnabled 
        .Option( "Option.NetlinesForSelectedItems.Enabled" ) = epcbGraphicsItemStateOffEnabled 
            .Option( "Option.ForceOutline" ) = epcbGraphicsItemStateOffEnabled 
            .Option( "Option.ForceSolid" ) = epcbGraphicsItemStateOffEnabled 
        .Option( "Option.FillPatterns" ) = epcbGraphicsItemStateOnEnabled 
            .Option( "Global.Option.HalfScreenCursor" ) = epcbGraphicsItemStateOffEnabled
            .Option( "Global.Option.FullScreenCursorDuringMoveOnly" ) = epcbGraphicsItemStateOffEnabled 
            .StringOption( "Global.Option.FullScreenCursorStyle" ) = "90Degree"
        .Option( "Global.Option.FullScreenCursor" ) = epcbGraphicsItemStateOnEnabled
            .StringOption( "Global.Option.PanSensitivity" ) = "7"
        .Option( "Global.Option.AutoPan" ) = epcbGraphicsItemStateOnEnabled 
        .Option( "Global.Option.PlaneDataBehindTraces" ) = epcbGraphicsItemStateOffEnabled
        .Option( "Global.Option.PlaneShapesOnTop" ) = epcbGraphicsItemStateOffEnabled 
        .Option( "Option.LegibleTextOnly" ) = epcbGraphicsItemStateOffEnabled 
        .Option( "Global.Option.NetNamesOnTraces" ) = epcbGraphicsItemStateOffEnabled 
        .Option( "Option.MirrorView" ) = epcbGraphicsItemStateOnEnabled 
        .Option( "Global.Option.TuningMeter" ) = epcbGraphicsItemStateOffEnabled 
        .Option( "Global.Option.LinkSketchNetlineDisplayToLayerVisibility" ) = epcbGraphicsItemStateOffEnabled 
            .StringOption( "Global.Option.ActiveClearanceRadius" ) = "100"
        .Visible( "General.ActiveClearance" ) = epcbGraphicsItemStateOnEnabled 
        '.Lock
        .Global.Color( "General.Pattern.Fixed" ) = utilityObj.NewColorPattern( 255, 255, 255, 100, 1, False, True ) 
        .Global.Color( "General.Pattern.SemiFixed" ) = utilityObj.NewColorPattern( 255, 255, 255, 100, 1, False, True )
        .Global.Color( "General.Pattern.Locked" ) = utilityObj.NewColorPattern( 255, 255, 255, 100, 1, True, True )
        .Global.Color( "General.Color.Background" ) = utilityObj.NewColorPattern( 255, 255, 255, 100, 0, False, True ) 
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
        .Option( "Option.Grids.Enabled" ) = epcbGraphicsItemStateOffEnabled 

        ' Color by group
        .Option( "Global.Option.ColorByGroup.Enabled" ) = epcbGraphicsItemStateOffEnabled 

        ' Color by net or class
        .Option( "Global.Option.ColorByNetClass.PatternConstrainedNets" ) = epcbGraphicsItemStateOffEnabled
        .Option( "Global.Option.ColorByNetClass.PreserveLayerColorOnPlanes" ) = epcbGraphicsItemStateOffEnabled
        .Option( "Global.Option.ColorByNetClass.UseObjectColorAsPatternBackground" ) = epcbGraphicsItemStateOffEnabled
        .Option( "Global.Option.ColorByNetClass.UseCMColors" ) = epcbGraphicsItemStateOffEnabled 
        .Option( "Global.Option.ColorByNetClass.Netlines" ) = epcbGraphicsItemStateOffEnabled 
        .Option( "Global.Option.ColorByNetClass.Traces" ) = epcbGraphicsItemStateOffEnabled 
        .Option( "Global.Option.ColorByNetClass.Pads" ) = epcbGraphicsItemStateOffEnabled 
        .Option( "Global.Option.ColorByNetClass.Planes" ) = epcbGraphicsItemStateOffEnabled 
        .Option( "Global.Option.ColorByNetClass.Vias" ) = epcbGraphicsItemStateOffEnabled 
        .Option( "Global.Option.ColorByNetClass.ConductiveShapes" ) = epcbGraphicsItemStateOffEnabled 
            '.Visible( "[Net].GND" ) = epcbGraphicsItemStateOnEnabled 
        .Option( "Global.Option.ColorByNetClass.Nets.Enabled" ) = epcbGraphicsItemStateOffEnabled 
        .Option( "Global.Option.ColorByNetClass.NetClasses.Enabled" ) = epcbGraphicsItemStateOffEnabled 
        .Option( "Global.Option.ColorByNetClass.ConstraintClasses.Enabled" ) = epcbGraphicsItemStateOffEnabled

        ' Object appearance
        ' .Lock
        ' .Global.Color( "Copper.Trace.*" ) = utilityObj.NewColorPattern( 255, 255, 255, 100, 0, False, True )  
        ' .Global.Color( "Copper.Pad.*" ) = utilityObj.NewColorPattern( 255, 255, 255, 100, 0, False, True ) 
        ' .Global.Color( "Copper.Plane.Data.*" ) = utilityObj.NewColorPattern( 255, 255, 255, 100, 0, False, True ) 
        ' .Unlock

        ' Board Objects
                .Visible( "Copper.Pad.Fiducial.Top" ) = epcbGraphicsItemStateOffEnabled 
                .Visible( "Copper.Pad.Fiducial.Bottom" ) = epcbGraphicsItemStateOnEnabled 
            .Option( "Option.Fiducials.Enabled" ) = epcbGraphicsItemStateOnEnabled   
                .Option( "Option.MountingHolePads.Enabled" ) = epcbGraphicsItemStateOnEnabled
            .Option( "Option.Holes.Enabled" ) = epcbGraphicsItemStateOnEnabled 
            .Visible( "Fabrication.Hole.Mounting" ) = epcbGraphicsItemStateOnEnabled 
                .Visible( "Board.BoardOutline" ) = epcbGraphicsItemStateOnEnabled 
                .Visible( "Board.ManufacturingOutline" ) = epcbGraphicsItemStateOffEnabled 
                .Visible( "Board.FixtureOutline" ) = epcbGraphicsItemStateOffEnabled 
                .Visible( "Board.Cavity" ) = epcbGraphicsItemStateOffEnabled 
                .Visible( "Fabrication.Hole.Contour" ) = epcbGraphicsItemStateOnEnabled 
                .Visible( "Fabrication.Hole.Contour.SpanNumbers" ) = epcbGraphicsItemStateOnEnabled 
                .Visible( "Board.DRCWindow" ) = epcbGraphicsItemStateOffEnabled
                .Visible( "Board.Origin.Board" ) = epcbGraphicsItemStateOffEnabled
                .Visible( "Board.Origin.NCDrill" ) = epcbGraphicsItemStateOffEnabled 
                .Visible( "Fabrication.RedlineLayer" ) = epcbGraphicsItemStateOffEnabled 
            .Option( "Option.BoardElements.Enabled" ) = epcbGraphicsItemStateOnEnabled 
                .Option( "Option.TextItems.PinProperties.Enabled" ) = epcbGraphicsItemStateOnEnabled 
                .Option( "Option.TextItems.CellProperties.Enabled" ) = epcbGraphicsItemStateOnEnabled
            .Option( "Option.TextItems.Enabled" ) = epcbGraphicsItemStateOffEnabled 
        .Option( "Option.BoardObjects.Enabled" ) = epcbGraphicsItemStateOnEnabled 
        '.Lock
        .Global.Color( "Copper.Pad.Fiducial.Bottom" ) = utilityObj.NewColorPattern( 204, 153, 0, 100, 0, False, True ) 
        .Global.Color( "Fabrication.Hole.Mounting" ) = utilityObj.NewColorPattern(  0, 128, 0, 100, 0, False, True ) 
        .Global.Color( "Board.BoardOutline" ) = utilityObj.NewColorPattern( 128, 128, 128, 100, 0, False, True ) 
        .Global.Color( "Board.Cavity" ) = utilityObj.NewColorPattern( 204, 153, 0, 100, 0, False, True ) 
        .Global.Color( "Fabrication.Hole.Contour" ) = utilityObj.NewColorPattern( 204, 153, 0, 100, 0, False, True ) 
        '.Unlock

        ' Fabrication objects	
        .Visible( "Fabrication.SolderMask.Bottom" ) = epcbGraphicsItemStateOnEnabled 
		.Visible( "Fabrication.SolderMask.Top" ) = epcbGraphicsItemStateOffEnabled 
		.Visible( "Fabrication.SolderPaste.Bottom" ) = epcbGraphicsItemStateOffEnabled 
		.Visible( "Fabrication.SolderPaste.Top" ) = epcbGraphicsItemStateOffEnabled 	
            .Visible( "Fabrication.Assembly.Part.Outline.Bottom" ) = epcbGraphicsItemStateOnEnabled 
            .Visible( "Fabrication.Assembly.Part.Outline.Top" ) = epcbGraphicsItemStateOnEnabled 
            .Visible( "Fabrication.Assembly.Part.Text.PartNumber.Bottom" ) = epcbGraphicsItemStateOffEnabled 
            .Visible( "Fabrication.Assembly.Part.Text.PartNumber.Top" ) = epcbGraphicsItemStateOffEnabled
            .Visible( "Fabrication.Assembly.Part.Text.RefDes.Bottom" ) = epcbGraphicsItemStateOffEnabled 
            .Visible( "Fabrication.Assembly.Part.Text.RefDes.Top" ) = epcbGraphicsItemStateOffEnabled
        .Option( "Option.Fabrication.AssemblyItems.Top" ) = epcbGraphicsItemStateOffEnabled 
        .Option( "Option.Fabrication.AssemblyItems.Bottom" ) = epcbGraphicsItemStateOnEnabled  
            .Visible( "Fabrication.Silkscreen.Part.Outline.Bottom" ) = epcbGraphicsItemStateOnEnabled 
            .Visible( "Fabrication.Silkscreen.Part.Outline.Top" ) = epcbGraphicsItemStateOnEnabled 
            .Visible( "Fabrication.Silkscreen.Part.Text.PartNumber.Bottom" ) = epcbGraphicsItemStateOffEnabled 
            .Visible( "Fabrication.Silkscreen.Part.Text.PartNumber.Top" ) = epcbGraphicsItemStateOffEnabled 
            .Visible( "Fabrication.Silkscreen.Part.Text.RefDes.Top" ) = epcbGraphicsItemStateOffEnabled 
            .Visible( "Fabrication.Silkscreen.Part.Text.RefDes.Bottom" ) = epcbGraphicsItemStateOffEnabled 
            .Visible( "Fabrication.Silkscreen.Generated.Bottom" ) = epcbGraphicsItemStateOffEnabled 
            .Visible( "Fabrication.Silkscreen.Generated.Top" ) = epcbGraphicsItemStateOffEnabled 
        .Option( "Option.Fabrication.SilkscreenItems.Top" ) = epcbGraphicsItemStateOffEnabled 
        .Option( "Option.Fabrication.SilkscreenItems.Bottom" ) = epcbGraphicsItemStateOnEnabled 
            .Visible( "Fabrication.Assembly.TestPoint.Text.RefDes.Bottom" ) = epcbGraphicsItemStateOffEnabled 
            .Visible( "Fabrication.Assembly.TestPoint.Text.RefDes.Top" ) = epcbGraphicsItemStateOffEnabled 
            .Visible( "Fabrication.Silkscreen.TestPoint.Text.RefDes.Bottom" ) = epcbGraphicsItemStateOffEnabled 
            .Visible( "Fabrication.Silkscreen.TestPoint.Text.RefDes.Top" ) = epcbGraphicsItemStateOffEnabled 
            .Visible( "Fabrication.Silkscreen.TestPoint.Probe.Bottom" ) = epcbGraphicsItemStateOffEnabled 
            .Visible( "Fabrication.Silkscreen.TestPoint.Probe.Top" ) = epcbGraphicsItemStateOffEnabled 
        .Option( "Option.Fabrication.TestPointItems.Top" ) = epcbGraphicsItemStateOffEnabled 
        .Option( "Option.Fabrication.TestPointItems.Bottom" ) = epcbGraphicsItemStateOnEnabled 
        ' .Lock
        .Global.Color( "Fabrication.Soldermask.Bottom" ) = utilityObj.NewColorPattern( 204, 153, 0, 100, 0, False, True ) 
        .Global.Color( "Fabrication.Assembly.Part.Outline.Bottom" ) = utilityObj.NewColorPattern( 119, 119, 119, 100, 0, False, True ) 
        .Global.Color( "Fabrication.Silkscreen.Part.Outline.Bottom" ) = utilityObj.NewColorPattern( 0, 0, 0, 100, 0, False, True ) 
        .Global.Color( "Fabrication.Silkscreen.Part.Text.RefDes.Bottom" ) = utilityObj.NewColorPattern( 0, 0, 0, 100, 0, False, True ) 
        .Global.Color( "Fabrication.Silkscreen.TestPoint.Text.RefDes.Bottom" ) = utilityObj.NewColorPattern( 0, 0, 0, 100, 0, False, True ) 
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
		.Option( "Option.Fabrication.DrillDrawing" ) =  epcbGraphicsItemStateOffEnabled 

        ' User draft layers
        .Option( "Option.UserDraftLayers.Enabled" ) = epcbGraphicsItemStateOffEnabled 

        ' Detail View
		.Visible( "Fabrication.DetailViews" ) = epcbGraphicsItemStateOffEnabled 

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
    ' displayCtrlObj.SaveSchemeEx("Sys:MyAssemblyBottom")

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
