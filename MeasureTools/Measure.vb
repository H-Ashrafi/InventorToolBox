Imports Inventor

Public Class Measure
    Public Sub InteractiveMeasureDistance()
        ' Create a new Measure object.
        Dim oMeasure As New Measure

        ' Call the Measure method of the Measure object
        Call oMeasure.Measure(kDistanceMeasure)
    End Sub


    Public Sub InteractiveMeasureAngle()
        ' Create a new Measure object.
        Dim oMeasure As New Measure

        ' Call the Measure method of the Measure object
        Call oMeasure.Measure(kAngleMeasure)
    End Sub


    '*************************************************************
    ' The declarations and functions below need to be copied into
    ' a class module whose name is "Measure". The name can be
    ' changed but you'll need to change the declaration in the
    ' calling function "InteractiveMeasureDistance" and
    ' "InteractiveMeasureAngle" to use the new name.

    ' Declare the event objects
    Private WithEvents oInteractEvents As InteractionEvents
    Private WithEvents oMeasureEvents As MeasureEvents

    ' Declare a flag that's used to determine when measuring stops.
    Private bStillMeasuring As Boolean
    Private eMeasureType As MeasureTypeEnum

    Public Sub Measure(MeasureType As MeasureTypeEnum)

        eMeasureType = MeasureType

        ' Initialize flag.
        bStillMeasuring = True

    ' Create an InteractionEvents object.
    Set oInteractEvents = ThisApplication.CommandManager.CreateInteractionEvents

    ' Set a reference to the measure events.
    Set oMeasureEvents = oInteractEvents.MeasureEvents
    oMeasureEvents.Enabled = True

        ' Start the InteractionEvents object.
        oInteractEvents.Start()

        ' Start measure tool
        If eMeasureType = kDistanceMeasure Then
            oMeasureEvents.Measure kDistanceMeasure
    Else
            oMeasureEvents.Measure kAngleMeasure
    End If

        ' Loop until a selection is made.
        Do While bStillMeasuring
            DoEvents
        Loop

        ' Stop the InteractionEvents object.
        oInteractEvents.Stop()
        
    ' Clean up.
    Set oMeasureEvents = Nothing
    Set oInteractEvents = Nothing
End Sub

    Private Sub oInteractEvents_OnTerminate()
        ' Set the flag to indicate we're done.
        bStillMeasuring = False
    End Sub

    Private Sub oMeasureEvents_OnMeasure(ByVal MeasureType As MeasureTypeEnum, ByVal MeasuredValue As Double, ByVal Context As NameValueMap)

        Dim strMeasuredValue As String

        If eMeasureType = kDistanceMeasure Then
            strMeasureValue = ThisApplication.ActiveDocument.UnitsOfMeasure.GetStringFromValue(MeasuredValue, kDefaultDisplayLengthUnits)
            MsgBox "Distance = " & strMeasureValue, vbOKOnly, "Measure Distance"
    Else
            strMeasureValue = ThisApplication.ActiveDocument.UnitsOfMeasure.GetStringFromValue(MeasuredValue, kDefaultDisplayAngleUnits)
            MsgBox "Angle = " & strMeasureValue, vbOKOnly, "Measure Angle"
    End If

        ' Set the flag to indicate we're done.
        bStillMeasuring = False

    End Sub
End Class