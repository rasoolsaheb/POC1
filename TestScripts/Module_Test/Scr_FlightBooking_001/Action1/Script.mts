'#########################################################################################
'ScriptName				:	Scr_FlightBooking_001
'Author					:	csc
'Created On 			:	30/10/2015
'Description			:	This Script is used to execute book a flight from the application
'Pre-Conditions         :   None
'History				:
'#########################################################################################
Option Explicit

Dim ObjRS 			'Stores RecordSet Object 
Dim blnFailStatus	'Stores True/False Status
Dim iIteration      'Stores iteration number


'Initializing the Test Script
Call gfOnInitialize(Environment("TestName"))


'Open the Browser
Call gfOpenBrowser(Environment("browserType"),Environment("AppURL"))


'Login to application
Call gfLogin()


For iIteration = 1 to (gfGetTDRowCount(objRS)-1)

	
	'Retreving the recordset for the iteration
    Call gfGetRS(iIteration, objRS)
	
 @@ hightlight id_;_Browser("Welcome: Mercury Tours").Page("Find a Flight: Mercury").Image("findFlights")_;_script infofile_;_ZIP::ssf15.xml_;_
	'Entering Flight details
	Call gfEnterFlightFinderDetials(ObjRS)
	
 @@ hightlight id_;_Browser("Welcome: Mercury Tours").Page("Select a Flight: Mercury").Image("reserveFlights")_;_script infofile_;_ZIP::ssf20.xml_;_
	'Entering the Booking details
	Call gfSelectFlight(ObjRS)
	
	
	'Entering the Booking details
	Call gfEnterBookingDetials(ObjRS)


Next

'Closing the application
Call gfSignOffnClose()


'Terminating the Test Script
Call gfOnTerminate()


 @@ hightlight id_;_2093261384_;_script infofile_;_ZIP::ssf51.xml_;_