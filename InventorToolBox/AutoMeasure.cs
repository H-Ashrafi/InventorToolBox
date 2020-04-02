using Inventor;
namespace InventorToolBox
{
    public partial class AutoMeasure
    {
        //instance of this inventor application

        private Application ThisApplication { get; set; }
        //constructor
        public AutoMeasure(Application app)
        {
            ThisApplication = app;
        }

        // Declare the event objects
        private InteractionEvents oInteractEvents { get; set; }
        private MeasureEvents oMeasureEvents { get; set; }

        // Declare a flag that's used to determine when measuring stops.
        private bool bStillMeasuring;
        private MeasureTypeEnum eMeasureType;

        public void Measure(MeasureTypeEnum MeasureType)
        {
            eMeasureType = MeasureType;

            // Initialize flag.
            bStillMeasuring = true;

            // Create an InteractionEvents object.
            oInteractEvents = ThisApplication.CommandManager.CreateInteractionEvents();

            //Set a reference to the measure events.
            oMeasureEvents = oInteractEvents.MeasureEvents;
            oMeasureEvents.Enabled = true;

            // Start the InteractionEvents object.
            oInteractEvents.Start();

            // Start measure tool
            if (eMeasureType == MeasureTypeEnum.kDistanceMeasure)
            {
                oMeasureEvents.Measure(MeasureTypeEnum.kDistanceMeasure);
            }
            else
            {
                oMeasureEvents.Measure(MeasureTypeEnum.kAngleMeasure);
            }

            // Loop until a selection is made.
            while (bStillMeasuring)
                ThisApplication.UserInterfaceManager.DoEvents();

            // Stop the InteractionEvents object.
            oInteractEvents.Stop();

            //Clean up.
            oMeasureEvents = null;
            oInteractEvents = null;
        }

        private void oInteractEvents_OnTerminate()
        {
            // Set the flag to indicate we're done.
            bStillMeasuring = false;
        }

        private void oMeasureEvents_OnMeasure(MeasureTypeEnum MeasureType, double MeasuredValue, NameValueMap Context)
        {
            string strMeasuredValue;

            if (eMeasureType == MeasureTypeEnum.kDistanceMeasure)
            {
                strMeasuredValue = ThisApplication.ActiveDocument.UnitsOfMeasure.GetStringFromValue(MeasuredValue, UnitsTypeEnum.kDefaultDisplayLengthUnits);
                System.Windows.Forms.MessageBox.Show("Distance = " + strMeasuredValue); 
            }
            else
            {
                strMeasuredValue = ThisApplication.ActiveDocument.UnitsOfMeasure.GetStringFromValue(MeasuredValue, UnitsTypeEnum.kDefaultDisplayAngleUnits);
                System.Windows.Forms.MessageBox.Show("Angle = " + strMeasuredValue); 
            }
            // Set the flag to indicate we're done.
            bStillMeasuring = false;

        }
    }
}
