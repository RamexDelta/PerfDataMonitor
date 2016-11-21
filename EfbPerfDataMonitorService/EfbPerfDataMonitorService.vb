Option Explicit On
Imports System.Threading
Imports System.Timers
'Imports EfbCommsControllerDll
Imports System.Reflection
Imports System.ServiceProcess
Imports System.Xml


Public Class EfbPerfDataMonitorService

    Dim PdiStatus As Boolean    'Records whether the PDI is responding to requests for data

    Dim AlpsWeight As Integer
    Dim AlpsFlex As Integer
    Dim AlpsFlap As String

    Dim CommsAllowed As Boolean = False
    Dim PreviousCommsAllowed As Boolean = False
    Dim ProposedCommsStatus As Boolean
    Dim ProposedCommsStatusTime As Date = Now


    'Public DllInitialised As Boolean
    Public tBagCheckTimer As System.Timers.Timer
    Public runOnceTimer As System.Timers.Timer
    Public valueMonitorTimer As System.Timers.Timer
    Public makeComparisonTimer As System.Timers.Timer
    Public modemControlTimer As System.Timers.Timer
    Public dllMonitorTimer As System.Timers.Timer
    Public A717ReaderTimer As System.Timers.Timer

    Dim valueMonitorHeartBeat As Integer = 0

    Dim ModemResponse As Integer

    Dim LogStream As System.IO.StreamWriter
    Dim ResultsStream As System.IO.StreamWriter
    Dim LogPreviousDiscreteState As String = ""
    Dim LogWriteTime As Date = Now

    Private stopping As Boolean
    Private stoppedEvent As ManualResetEvent

    Public LabelCount As Integer
    Public cycle As Integer
    Public pTimeStamp As Integer
    Public A717RnFrame As New A717Frame
    Public A717RnSuperFrame As A717SuperFrame

    Public FrameCounterSubFrame As Integer
    Public FrameCounterWord As Integer
    Public WordCount As Integer

    Public GroundSpeedLabel As LabelInstance
    Public GroundSpeed As Integer
    Public Eng1StartTime As DateTime
    Public Eng2StartTime As DateTime

    Public Sub New()
        InitializeComponent()
        Me.stopping = False
        Me.stoppedEvent = New ManualResetEvent(False)
    End Sub


    ''' <summary>
    ''' The function is executed when a Start command is sent to the service
    ''' by the SCM or when the operating system starts (for a service that 
    ''' starts automatically). It specifies actions to take when the service 
    ''' starts. In this code sample, OnStart logs a service-start message to 
    ''' the Application log, and queues the main service function for 
    ''' execution in a thread pool worker thread.
    ''' </summary>
    ''' <param name="args">Command line arguments</param>
    ''' <remarks>
    ''' A service application is designed to be long running. Therefore, it 
    ''' usually polls or monitors something in the system. The monitoring is 
    ''' set up in the OnStart method. However, OnStart does not actually do 
    ''' the monitoring. The OnStart method must return to the operating 
    ''' system after the service's operation has begun. It must not loop 
    ''' forever or block. To set up a simple monitoring mechanism, one 
    ''' general solution is to create a timer in OnStart. The timer would 
    ''' then raise events in your code periodically, at which time your 
    ''' service could do its monitoring. The other solution is to spawn a 
    ''' new thread to perform the main service functions, which is 
    ''' demonstrated in this code sample.
    ''' 
    ''' </remarks>
    ''' 


    ' Program Flow
    ' 1. OnStart
    ' 2. tBagCheckTick
    ' 3. runOnceTick
    ' 4. valueMonitorTick. A717ReaderTimerTick also started
    ' 5. makeComparisonTick

    Protected Overrides Sub OnStart(ByVal args() As String)
        Dim navApiStatus As Boolean

        ' Log a service start message to the Application log.
        Me.EventLog1.WriteEntry("EfbPerfDataMonitorService in OnStart. Version " & Assembly.GetExecutingAssembly().GetName().Version.ToString)

        ' Open Log File
        LogStream = My.Computer.FileSystem.OpenTextFileWriter("C:\EFB\Logs\" & Format(DateAndTime.Now, "yyyyMMdd-HHmmss") & "-PerfMonitor.txt", True)
        LogStream.AutoFlush = True
        Call WriteToLog("EFB Performace Data Monitor Service version:  " & Assembly.GetExecutingAssembly().GetName().Version.ToString)

        'Set up all timers (note that they are not all needed yet)

        'Set up and run the timer used to check for the TBag Service. 
        'This is the only timer that is enabled at this point
        tBagCheckTimer = New System.Timers.Timer(2000)  'Runs every 2 seconds
        AddHandler tBagCheckTimer.Elapsed, AddressOf Me.tBagCheckTick
        tBagCheckTimer.AutoReset = False
        tBagCheckTimer.Enabled = True

        'Set up and run the initializing timer. This will only run once, but in doing so will enable the timer that runs indefinately
        runOnceTimer = New System.Timers.Timer(1)
        AddHandler runOnceTimer.Elapsed, AddressOf Me.runOnceTick
        runOnceTimer.AutoReset = False
        runOnceTimer.Enabled = False

        'Set up the main navAero monitoring timer. Note that it is not yet enabled. It's needed to be set up here because the WriteToLog procedure needs it's interval value 
        valueMonitorTimer = New System.Timers.Timer(1000)
        AddHandler valueMonitorTimer.Elapsed, AddressOf Me.valueMonitorTick
        valueMonitorTimer.AutoReset = False
        valueMonitorTimer.Enabled = False

        'Set up the main Comparision timer. Note that it is not yet enabled. It's needed to be set up here because the WriteToLog procedure needs it's interval value 
        makeComparisonTimer = New System.Timers.Timer(10000)
        AddHandler makeComparisonTimer.Elapsed, AddressOf Me.makeComparisonTick
        makeComparisonTimer.AutoReset = False
        makeComparisonTimer.Enabled = False

        'Set up the A717 Databus reading timeer. Note that it is not yet enabled.
        A717ReaderTimer = New System.Timers.Timer(900) 'runs every 900 miliseconds
        AddHandler A717ReaderTimer.Elapsed, AddressOf Me.A717ReaderTimerTick
        A717ReaderTimer.AutoReset = True
        A717ReaderTimer.Enabled = False

    End Sub

    Private Sub tBagCheckTick(sender As Object, e As Timers.ElapsedEventArgs)
        'This timer is the only one enabled by the onStart method.
        'This will loop until the tBag Service is found to be running
        Dim controller As New ServiceController("tBag Plugin Service")

        If controller.Status = ServiceControllerStatus.Running Then
            'tBag service is running
            Call WriteToLog("tBag Service Status: " & controller.Status.ToString)
            'Initialise the navAeroDll
            Dim navApiStatus As Boolean
            navApiStatus = DatabusClass.navAPI_Initialize
            If navApiStatus = True Then
                'Successfully Initialised navAPI
                Me.EventLog1.WriteEntry("navAPI Initialization: " & navApiStatus.ToString & vbCrLf & "navAPI Version: " & ApiVersion())
                Call WriteToLog("navAPI Status:       " & navApiStatus.ToString())
                Call WriteToLog("navAPI Version:      " & ApiVersion())
                Call WriteToLog("Device:              " & Environment.MachineName.ToString)
                'Disable this timer
                tBagCheckTimer.Enabled = False
                'Enable the runOnce timer
                runOnceTimer.Enabled = True
            Else
                'Keep this timer running
                tBagCheckTimer.Enabled = True
                'navAPI failed to initialize
                Me.EventLog1.WriteEntry("navAPI Initialization failed.")
                Call WriteToLog("navAPI Initialization failed.")
            End If
        Else
            'tBag service is stil not running
            Call WriteToLog("tBag Service Status: " & controller.Status.ToString)
            'Keep this timer running
            tBagCheckTimer.Enabled = True
        End If
    End Sub

    Private Sub runOnceTick(sender As Object, e As Timers.ElapsedEventArgs)
        Call WriteToLog("In runOnceTick")
        'This is the second block to be run, and will be run once only
        'Dim doc As New XmlDocument()

        ''Ezxtract data from the ALPS output file
        'If My.Computer.FileSystem.FileExists(My.Computer.FileSystem.SpecialDirectories.ProgramFiles & "\Alps\TempData\InputRecords\\AlpsSelectedData.txt") = True Then
        '    doc.Load(My.Computer.FileSystem.SpecialDirectories.ProgramFiles & "\Alps\TempData\InputRecords\\AlpsSelectedData.txt")
        '    'http://stackoverflow.com/questions/17276580/looping-through-xml-file-using-vb-net
        '    Dim nodelist As XmlNodeList
        '    nodelist = doc.GetElementsByTagName("pctoRes")
        '    'for each structure used below, but there should only be one node in the file
        '    For Each node As XmlElement In nodelist
        '        Try
        '            Debug.WriteLine(node("mtow").InnerText)
        '            Debug.WriteLine(node("flaps").InnerText)

        '            AlpsWeight = Val(node("mtow").InnerText)
        '            AlpsFlex = Val(node("flexTemp").InnerText)
        '            AlpsFlap = node("flaps").InnerText
        '            Call WriteToLog("ALPS extracted data :")
        '            Call WriteToLog(" - Weight : " & AlpsWeight & "kg (" & Int(Val(AlpsWeight) / 0.453592) & "lbs)")
        '            Call WriteToLog(" - Flaps  : " & AlpsFlap)
        '            Call WriteToLog(" - Flex   : " & AlpsFlex)
        '        Catch
        '            Call WriteToLog("Error extracting data from ALPS output")
        '        End Try
        '    Next
        'Else
        '    Call WriteToLog(My.Computer.FileSystem.SpecialDirectories.ProgramFiles.ToString & "\Alps\TempData\InputRecords\\AlpsSelectedData.txt could not be found.")
        'End If



        'Extract Label List
        ReDim DatabusClass.LabelList(-1)
        Dim doc2 As New XmlDocument()
        If My.Computer.FileSystem.FileExists("C:\EFB\PerfDataMonitor\LabelListPerfMon.xml") = True Then
            doc2.Load("C:\EFB\PerfDataMonitor\LabelListPerfMon.xml")

            'http://stackoverflow.com/questions/17276580/looping-through-xml-file-using-vb-net
            Dim nodelist2 As XmlNodeList

            'Read All A717 nodes from the xml
            Call WriteToLog("Reading A717 Labels from XML file")
            nodelist2 = doc2.GetElementsByTagName("Label717")

            For Each node As XmlElement In nodelist2
                Debug.WriteLine(node("Description").InnerText)
                ReDim Preserve DatabusClass.LabelList(UBound(DatabusClass.LabelList) + 1)
                DatabusClass.LabelList(UBound(DatabusClass.LabelList)) = New LabelInstance
                DatabusClass.LabelList(UBound(DatabusClass.LabelList)).LabelType = DatabusClass.LabelType.A717
                DatabusClass.LabelList(UBound(DatabusClass.LabelList)).Description = node("Description").InnerText
                DatabusClass.LabelList(UBound(DatabusClass.LabelList)).Encoding = DirectCast([Enum].Parse(GetType(DatabusClass.Encoding), node("Encoding").InnerText), Integer) 'From http://stackoverflow.com/questions/852141/parse-a-string-to-an-enum-value-in-vb-net
                DatabusClass.LabelList(UBound(DatabusClass.LabelList)).Scale = node("Scale").InnerText

                DatabusClass.LabelList(UBound(DatabusClass.LabelList)).Cycle = node("Cycle").InnerText
                DatabusClass.LabelList(UBound(DatabusClass.LabelList)).SubFrame = node("SubFrame").InnerText
                DatabusClass.LabelList(UBound(DatabusClass.LabelList)).Word = node("Word").InnerText
                DatabusClass.LabelList(UBound(DatabusClass.LabelList)).Msb = node("Msb").InnerText
                DatabusClass.LabelList(UBound(DatabusClass.LabelList)).Lsb = node("Lsb").InnerText
                DatabusClass.LabelList(UBound(DatabusClass.LabelList)).SignBit = node("SignBit").InnerText
                Call WriteToLog("A717 Label Read as follows:" & vbCrLf &
                            "- Description     : " & DatabusClass.LabelList(UBound(DatabusClass.LabelList)).Description & vbCrLf &
                            "- Encoding        : " & DatabusClass.LabelList(UBound(DatabusClass.LabelList)).Encoding.ToString & vbCrLf &
                            "- Scale           : " & DatabusClass.LabelList(UBound(DatabusClass.LabelList)).Scale & vbCrLf &
                            "- Cycle           : " & DatabusClass.LabelList(UBound(DatabusClass.LabelList)).Cycle & vbCrLf &
                            "- SubFrame        : " & DatabusClass.LabelList(UBound(DatabusClass.LabelList)).SubFrame & vbCrLf &
                            "- Word            : " & DatabusClass.LabelList(UBound(DatabusClass.LabelList)).Word & vbCrLf &
                            "- MSB             : " & DatabusClass.LabelList(UBound(DatabusClass.LabelList)).Msb & vbCrLf &
                            "- LSB             : " & DatabusClass.LabelList(UBound(DatabusClass.LabelList)).Lsb & vbCrLf &
                            "- SignBit         : " & DatabusClass.LabelList(UBound(DatabusClass.LabelList)).SignBit)
            Next


            'Read All A429 nodes from the xml
            Call WriteToLog("Reading A429 Labels from XML file")
            nodelist2 = Nothing
            nodelist2 = doc2.GetElementsByTagName("Label429")
            For Each node As XmlElement In nodelist2
                Debug.WriteLine(node("Description").InnerText)
                ReDim Preserve DatabusClass.LabelList(UBound(DatabusClass.LabelList) + 1)
                DatabusClass.LabelList(UBound(DatabusClass.LabelList)) = New LabelInstance
                DatabusClass.LabelList(UBound(DatabusClass.LabelList)).LabelType = DatabusClass.LabelType.A429
                DatabusClass.LabelList(UBound(DatabusClass.LabelList)).Description = node("Description").InnerText
                DatabusClass.LabelList(UBound(DatabusClass.LabelList)).Encoding = DirectCast([Enum].Parse(GetType(DatabusClass.Encoding), node("Encoding").InnerText), Integer) 'From http://stackoverflow.com/questions/852141/parse-a-string-to-an-enum-value-in-vb-net
                DatabusClass.LabelList(UBound(DatabusClass.LabelList)).Scale = node("Scale").InnerText

                DatabusClass.LabelList(UBound(DatabusClass.LabelList)).Channel = node("ChannelNo").InnerText
                DatabusClass.LabelList(UBound(DatabusClass.LabelList)).Label = node("LabelNo").InnerText
                DatabusClass.LabelList(UBound(DatabusClass.LabelList)).SigBits = node("SigBits").InnerText
                'DatabusClass.LabelList(UBound(DatabusClass.LabelList)).DecodeMechanism = CType(node("DecodeMechanism").InnerText, DatabusClass.DecodeMechanism)
                Try
                    DatabusClass.LabelList(UBound(DatabusClass.LabelList)).DecodeMechanism = DirectCast([Enum].Parse(GetType(DatabusClass.DecodeMechanism), node("DecodeMechanism").InnerText), Integer)
                Catch
                    'Do nothing
                End Try
                Call WriteToLog("A429 Label Read as follows:" & vbCrLf &
                            "- Description     : " & DatabusClass.LabelList(UBound(DatabusClass.LabelList)).Description & vbCrLf &
                            "- Encoding        : " & DatabusClass.LabelList(UBound(DatabusClass.LabelList)).Encoding.ToString & vbCrLf &
                            "- Scale           : " & DatabusClass.LabelList(UBound(DatabusClass.LabelList)).Scale & vbCrLf &
                            "- Channel         : " & DatabusClass.LabelList(UBound(DatabusClass.LabelList)).Channel & vbCrLf &
                            "- Label           : " & DatabusClass.LabelList(UBound(DatabusClass.LabelList)).Label & vbCrLf &
                            "- SigBits         : " & DatabusClass.LabelList(UBound(DatabusClass.LabelList)).SigBits & vbCrLf &
                            "- DecodeMechanism : " & DatabusClass.LabelList(UBound(DatabusClass.LabelList)).DecodeMechanism)
            Next

            'Read the WordCount node from the xml
            Call WriteToLog("Reading WordCount from XML file")
            nodelist2 = Nothing
            nodelist2 = doc2.GetElementsByTagName("WordCount")
            For Each node As XmlElement In nodelist2
                WordCount = node("Count").InnerText
                Call WriteToLog("WordCount  Read as follows:" & WordCount)
            Next


            'Set the values for the FrameCounter SubFrame and Word
            'from Stack OverFlow 7795149
            Dim matches = From c In DatabusClass.LabelList
                          Where c.Description = "Frame Counter"
                          Select c
            Try
                Call WriteToLog("Frame Counter Word figured as " & matches(0).Word)
                Call WriteToLog("Frame Counter Sub Frame figured as " & matches(0).SubFrame)
                FrameCounterSubFrame = matches(0).SubFrame
                FrameCounterWord = matches(0).Word
            Catch ex As Exception
                Call WriteToLog("Error figuring out Frame Counter Location")
            End Try

            'Figure out location (in the array of labels) of the Ground Speed word
            matches = From c In DatabusClass.LabelList
                      Where c.Description = "Ground Speed"
                      Select c
            Try
                GroundSpeedLabel = matches(0)
                Call WriteToLog("Ground Speed Sub Frame figured as " & GroundSpeedLabel.SubFrame)
                Call WriteToLog("Ground Speed Word figured as " & GroundSpeedLabel.Word)
            Catch ex As Exception
                Call WriteToLog("Error figuring out Ground Speed Location")
            End Try

            'Set Up the Engine discreets
            DatabusClass.PreviousEng1Lop = False
            DatabusClass.PreviousEng2Lop = False


            valueMonitorTimer.Enabled = True
            A717ReaderTimer.Enabled = True

        Else
            'Couldn't find the XML file that has the list of labels
            'Can't proceed
            Call WriteToLog("Couldn't find C:\EFB\PerfDataMonitor\LabelListPerfMon.xml")
            Call WriteToLog("Ending")
            Call OnStop()
        End If

        'If AlpsWeight <> 0 Then
        '    'Proceed
        '    valueMonitorTimer.Enabled = True
        '    A717ReaderTimer.Enabled = True
        'Else
        '    'failed to read ALPS weight
        '    'Shut down the service
        '    Call WriteToLog("Stopping Service from runOnceTick")
        '    Call OnStop()
        'End If

    End Sub


    Private Sub valueMonitorTick(sender As Object, e As Timers.ElapsedEventArgs)
        'Call WriteToLog("In valueMonitorTick")

        Dim Eng1RunTime As Integer = 0
        Dim Eng2RunTime As Integer = 0

        'Decode Ground Speed
        Try
            GroundSpeed = DatabusClass.BnrDecode717B(A717RnFrame.SubFrames(GroundSpeedLabel.SubFrame).WordsRaw(GroundSpeedLabel.Word), GroundSpeedLabel.Lsb, GroundSpeedLabel.Msb, GroundSpeedLabel.SignBit, GroundSpeedLabel.Scale).DecodedPayload
            Call WriteToLog("Ground Speed = " & GroundSpeed)
        Catch ex As Exception
            Call WriteToLog("Error attempting to decode Ground Speed")
        End Try

        Try
            'Check Engine Oil Pressure discretes
            'True = Engine Operating 
            DatabusClass.Eng1Lop = DatabusClass.DiscreteStatus(2)
            DatabusClass.Eng2Lop = DatabusClass.DiscreteStatus(3)
            'Check Weight on Wheels
            'True = In Air
            DatabusClass.WoW = DatabusClass.DiscreteStatus(1)
        Catch ex As Exception
            Call WriteToLog("Error attempting to retrieve Discetes")
        End Try


        Try
            If DatabusClass.PreviousEng1Lop = False And DatabusClass.Eng1Lop = True Then
                'Engine 1 has been switched on
                Eng1StartTime = DateTime.Now
                Call WriteToLog("Eng 1 now operating")
            End If
        Catch ex As Exception
            Call WriteToLog("Error attempting to derive Eng1StartTime")
        End Try
        Try
            If DatabusClass.PreviousEng2Lop = False And DatabusClass.Eng2Lop = True Then
                'Engine 2 has been switched on
                Eng2StartTime = DateTime.Now
                Call WriteToLog("Eng 2 now operating")
            End If
        Catch ex As Exception
            Call WriteToLog("Error attempting to derive Eng2StartTime")
        End Try
        Try
            If DatabusClass.PreviousEng1Lop = True And DatabusClass.Eng1Lop = True Then
                'Engine 1 is running
                Eng1RunTime = DateDiff(DateInterval.Second, Eng1StartTime, Now)
                'Call WriteToLog("Eng1RunTime = " & Eng1RunTime)
            End If
        Catch ex As Exception
            Call WriteToLog("Error attempting to derive Eng1RunTime")
        End Try
        Try
            If DatabusClass.PreviousEng2Lop = True And DatabusClass.Eng2Lop = True Then
                'Engine 2 is running
                Eng2RunTime = DateDiff(DateInterval.Second, Eng2StartTime, Now)
                'Call WriteToLog("Eng2RunTime = " & Eng2RunTIme)
            End If
        Catch ex As Exception
            Call WriteToLog("Error attempting to derive Eng2RunTime")
        End Try

        Try
            'Check Current Status
            If (GroundSpeed > 10 And GroundSpeed < 900) And Eng1RunTime > 10 And Eng2RunTime > 10 Then
                'If Eng1RunTime > 10 And GroundSpeed < 900 And Eng2RunTime > 10 Then
                'Proceed with the check
                Call WriteToLog("Ground Speed = " & GroundSpeed & " | Both Engines Running")
                Call WriteToLog("Proceeding")
                'Enable the next timer
                makeComparisonTimer.Enabled = True
            ElseIf (GroundSpeed > 100 And GroundSpeed < 900) Or DatabusClass.WoW = True Then
                'Ground Speed > 100 or aircraft is in air.
                'stop the service
                Call WriteToLog("Ground Speed = " & GroundSpeed & " | Wow = " & DatabusClass.WoW)
                Call WriteToLog("Stopping Service because Ground Speed > 100, or aircraft is In Air")
                'Stop The Service
                [Stop]()
            Else
                'Neither the Continue condition or the Stop condition were met
                'Keep looping
                'Set the Previous Eng LOP values
                DatabusClass.PreviousEng1Lop = DatabusClass.Eng1Lop
                DatabusClass.PreviousEng2Lop = DatabusClass.Eng2Lop
                valueMonitorTimer.Enabled = True
            End If
        Catch ex As Exception
            Call WriteToLog("Error attempting to check current status.")
            valueMonitorTimer.Enabled = True
        End Try

    End Sub


    Private Sub makeComparisonTick(sender As Object, e As Timers.ElapsedEventArgs)

        'Dim AircraftWeight429 As Integer
        'Dim AircraftWeight717 As Integer
        'Dim AircraftFlaps As Integer
        Dim rString As String = ""

        Call WriteToLog("Will make comparision now.")
        Call WriteToLog("Reading data from ALPS file")


        Dim doc As New XmlDocument()

        'Ezxtract data from the ALPS output file
        If My.Computer.FileSystem.FileExists(My.Computer.FileSystem.SpecialDirectories.ProgramFiles & "\Alps\TempData\InputRecords\AlpsSelectedData.txt") = True Then
            doc.Load(My.Computer.FileSystem.SpecialDirectories.ProgramFiles & "\Alps\TempData\InputRecords\AlpsSelectedData.txt")
            'http://stackoverflow.com/questions/17276580/looping-through-xml-file-using-vb-net
            Dim nodelist As XmlNodeList
            nodelist = doc.GetElementsByTagName("pctoRes")
            'for each structure used below, but there should only be one node in the file
            For Each node As XmlElement In nodelist
                Try
                    Debug.WriteLine(node("mtow").InnerText)
                    Debug.WriteLine(node("flaps").InnerText)

                    AlpsWeight = Val(node("mtow").InnerText)
                    AlpsFlex = Val(node("flexTemp").InnerText)
                    AlpsFlap = node("flaps").InnerText
                    Call WriteToLog("ALPS extracted data :")
                    Call WriteToLog(" - Weight : " & AlpsWeight & "kg (" & Int(Val(AlpsWeight) / 0.453592) & "lbs)")
                    Call WriteToLog(" - Flaps  : " & AlpsFlap)
                    Call WriteToLog(" - Flex   : " & AlpsFlex)
                Catch
                    Call WriteToLog("Error extracting data from ALPS output")
                End Try
            Next
        Else
            Call WriteToLog(My.Computer.FileSystem.SpecialDirectories.ProgramFiles.ToString & "\Alps\TempData\InputRecords\AlpsSelectedData.txt could not be found.")
        End If

        'Loop through the list of labels
        Dim h As Integer
        For h = 0 To UBound(DatabusClass.LabelList)
            Call WriteToLog("Comparing " & DatabusClass.LabelList(h).Description)
            rString = rString & DatabusClass.LabelList(h).Description & ","
            If DatabusClass.LabelList(h).LabelType = DatabusClass.LabelType.A717 Then
                'Handling 717 type
                Call WriteToLog("This is a 717 type")
                'get the 717 data
                If DatabusClass.LabelList(h).Cycle = -1 Then
                    'No cycle is specified, so SuperFrame concept doens't apply
                    Call WriteToLog("No Cycle Specified")
                    Try
                        Call WriteToLog("- Raw Value: " & A717RnFrame.SubFrames(DatabusClass.LabelList(h).SubFrame).WordsRaw(DatabusClass.LabelList(h).Word))
                        rString = rString & A717RnFrame.SubFrames(DatabusClass.LabelList(h).SubFrame).WordsRaw(DatabusClass.LabelList(h).Word) & ","
                    Catch ex As Exception
                        Call WriteToLog(" - Error extracting " & DatabusClass.LabelList(h).Description)
                        rString = rString & "Raw Extract Error,"
                    End Try
                    'Decode the 717 data
                    Try
                        Call WriteToLog(" - Decoded Value: " & DatabusClass.BnrDecode717B(A717RnFrame.SubFrames(DatabusClass.LabelList(h).SubFrame).WordsRaw(DatabusClass.LabelList(h).Word), DatabusClass.LabelList(h).Lsb, DatabusClass.LabelList(h).Msb, DatabusClass.LabelList(h).SignBit, DatabusClass.LabelList(h).Scale).DecodedPayload)
                        rString = rString & DatabusClass.BnrDecode717B(A717RnFrame.SubFrames(DatabusClass.LabelList(h).SubFrame).WordsRaw(DatabusClass.LabelList(h).Word), DatabusClass.LabelList(h).Lsb, DatabusClass.LabelList(h).Msb, DatabusClass.LabelList(h).SignBit, DatabusClass.LabelList(h).Scale).DecodedPayload.ToString & ","
                    Catch ex As Exception
                        Call WriteToLog("- Error Decoding " & DatabusClass.LabelList(h).Description)
                        rString = rString & "Decode Error,"
                    End Try
                Else
                    'A cycle is specified, so need to extract from correct cycle
                    Call WriteToLog("A Cycle Was Specified")
                    Try
                        Call WriteToLog("- Raw Value: " & A717RnSuperFrame.Frames(DatabusClass.LabelList(h).Cycle).SubFrames(DatabusClass.LabelList(h).SubFrame).WordsRaw(DatabusClass.LabelList(h).Word))
                        rString = rString & A717RnSuperFrame.Frames(DatabusClass.LabelList(h).Cycle).SubFrames(DatabusClass.LabelList(h).SubFrame).WordsRaw(DatabusClass.LabelList(h).Word) & ","
                    Catch ex As Exception
                        Call WriteToLog(" - Error extracting " & DatabusClass.LabelList(h).Description)
                        rString = rString & "Raw Extract Error,"
                    End Try
                End If


            ElseIf DatabusClass.LabelList(h).LabelType = DatabusClass.LabelType.A429 Then
                'Handling 429 type
                Call WriteToLog("This is a 429 type")
                'Get the data
                Try
                    Call WriteToLog(" - Raw Value: " & DatabusClass.Get429RawData(DatabusClass.LabelList(h).Label, DatabusClass.LabelList(h).Channel).PacketData.ToString)
                    rString = rString & DatabusClass.Get429RawData(DatabusClass.LabelList(h).Label, DatabusClass.LabelList(h).Channel).PacketData.ToString & ","
                Catch ex As Exception
                    Call WriteToLog(" - Error Extracting " & DatabusClass.LabelList(h).Description)
                    rString = rString & "Raw Extract Error,"
                End Try
                'Decode the data
                Try
                    Call WriteToLog(" - Decoded Value: " & DatabusClass.BnrDecode429B(DatabusClass.Get429RawData(DatabusClass.LabelList(h).Label, DatabusClass.LabelList(h).Channel).PacketData, DatabusClass.LabelList(h).DecodeMechanism, DatabusClass.LabelList(h).SigBits, DatabusClass.LabelList(h).Scale).DecodedPayload)
                    rString = rString & DatabusClass.BnrDecode429B(DatabusClass.Get429RawData(DatabusClass.LabelList(h).Label, DatabusClass.LabelList(h).Channel).PacketData, DatabusClass.LabelList(h).DecodeMechanism, DatabusClass.LabelList(h).SigBits, DatabusClass.LabelList(h).Scale).DecodedPayload.ToString & ","
                Catch ex As Exception
                    Call WriteToLog("- Error Decoding " & DatabusClass.LabelList(h).Description)
                    rString = rString & "Raw Decode Error,"
                End Try
            End If
        Next h

        'write results to dedicated file
        ResultsStream = My.Computer.FileSystem.OpenTextFileWriter("C:\EFB\Logs\" & "PerfMonitorResults.txt", True)
        ResultsStream.AutoFlush = True
        Dim temp As String = Format(Now, "MMdd-HH:mm:ss") & ","
        temp = temp & DiscreteDecode(DatabusClass.WoW) & "," & DiscreteDecode(DatabusClass.Eng1Lop) & "," & DiscreteDecode(DatabusClass.Eng2Lop) & "," & GroundSpeed & "," & AlpsWeight & "," & AlpsFlap & "," & AlpsFlex & ","
        temp = temp & rString & Assembly.GetExecutingAssembly().GetName().Version.ToString
        ResultsStream.WriteLine(temp)
        ResultsStream.Close()
        ResultsStream = Nothing

        'Stop The Service
        [Stop]()

    End Sub

    Private Sub A717ReaderTimerTick(sender As Object, e As Timers.ElapsedEventArgs)
        'Call WriteToLog("In A717ReaderTimerTick")

        Dim TempSubFrame As A717SubFrame
        Dim i As Integer
        Dim temp As String
        'Dim cycle As Integer
        Dim tempfs As New DecodedPacket
        Dim tempFrame As Integer

        'Dim WordCount As Integer = 256
        'Dim FrameCounterSubFrame = 4
        'Dim FrameCounterWord = 13
        'DEB: Frame 2, Word 113 (has only 128 words)
        'Some Others: Frame 4, Word 13 (these have 256 words)

        'above values now read from the LabelList xml file
        Try
            TempSubFrame = DatabusClass.Get717SubFrame(WordCount)
            Call WriteToLog("TempSubFrame has been populated. Status is " & TempSubFrame.SuccessStatus & ". Word1 is " & TempSubFrame.WordsRaw(1).ToString)
        Catch ex As Exception
            Call WriteToLog("Error attempting to populate TempSubFrame (in A717ReaderTimerTick block)")
        End Try

        'Call WriteToLog(TempSubFrame.Timestamp & "- " & TempSubFrame.WordsRaw(1).ToString)
        If TempSubFrame.Timestamp <> pTimeStamp And TempSubFrame.SuccessStatus = True Then
            'Call WriteToLog(TempSubFrame.WordsRaw(1) & " - Frame Counter: " & TempSubFrame.WordsRaw(FrameCounterWord))
            Select Case TempSubFrame.WordsRaw(1)
                Case 583
                    'Subframe 1
                    A717RnFrame.SubFrames(1) = TempSubFrame
                    tempFrame = 1
                Case 1464
                    'Subframe 2
                    A717RnFrame.SubFrames(2) = TempSubFrame
                    tempFrame = 2
                Case 2631
                    'Subframe 3
                    A717RnFrame.SubFrames(3) = TempSubFrame
                    tempFrame = 3
                Case 3512
                    'Subframe 4
                    A717RnFrame.SubFrames(4) = TempSubFrame
                    tempFrame = 4
            End Select

            'Call WriteToLog("tempframe: " & tempFrame)
            If tempFrame = FrameCounterSubFrame Then
                'dealing with the frame the holds the cycle
                'figure out cycle
                'Call WriteToLog("Sending " & A717RnFrame.SubFrames(FrameCounterSubFrame).WordsRaw(FrameCounterWord))
                'Call WriteToLog("Figuring Cycle - " & A717RnFrame.SubFrames(FrameCounterSubFrame).WordsRaw(FrameCounterWord) & " - " & DatabusClass.BnrDecode717B(A717RnFrame.SubFrames(FrameCounterSubFrame).WordsRaw(FrameCounterWord), 1, 12, 0, 4096).DecodedPayload)
                tempfs = DatabusClass.BnrDecode717B(A717RnFrame.SubFrames(FrameCounterSubFrame).WordsRaw(FrameCounterWord), 1, 4, 0, 16)
                'Call WriteToLog("Cycle is: " & DatabusClass.BnrDecode717B(A717RnFrame.SubFrames(FrameCounterSubFrame).WordsRaw(FrameCounterWord), 1, 4, 0, 16).DecodedPayload)
                'tempfs = DatabusClass.BnrDecode717B(1056, 1, 4, 0, 16)
                cycle = tempfs.DecodedPayload
                Call WriteToLog("tempframe matches FrameCounterSubFrame (" & tempFrame & "/" & FrameCounterSubFrame & ") Cycle is " & cycle)
            End If
            If tempFrame = 4 Then
                Call WriteToLog("Populating cycle " & cycle & ". Frame is " & A717RnFrame.SubFrames(FrameCounterSubFrame).WordsRaw(FrameCounterWord))
                If A717RnFrame.SubFrames(1) IsNot Nothing Then A717RnSuperFrame.Frames(cycle).SubFrames(1) = A717RnFrame.SubFrames(1)
                If A717RnFrame.SubFrames(2) IsNot Nothing Then A717RnSuperFrame.Frames(cycle).SubFrames(2) = A717RnFrame.SubFrames(2)
                If A717RnFrame.SubFrames(3) IsNot Nothing Then A717RnSuperFrame.Frames(cycle).SubFrames(3) = A717RnFrame.SubFrames(3)
                If A717RnFrame.SubFrames(4) IsNot Nothing Then A717RnSuperFrame.Frames(cycle).SubFrames(4) = A717RnFrame.SubFrames(4)
            End If
            pTimeStamp = TempSubFrame.Timestamp  'Store the timestamp of the current subframe
        Else
            'Call WriteToLog("Same Frame")
            'Same frame. Do Nothing
            '            textBox2.Text = "Same frame " & pTimeStamp & vbCrLf & textBox2.Text
        End If

    End Sub



    Private Function ApiVersion() As String
        Dim versionMajor As Byte
        Dim versionMinor As Byte
        Dim configuration As UInteger
        Dim ResultFlag As Boolean

        ResultFlag = DatabusClass.navAPI_GetVersion(versionMajor, versionMinor, configuration)
        If ResultFlag = True Then
            ApiVersion = Convert.ToString(versionMajor) & "." & Convert.ToString(versionMinor) & " (" & configuration & ") Then"
        Else
            ApiVersion = "Unknown"
        End If
    End Function

    Private Function DiscreteStatus(DiscreteChannel As Integer) As Boolean?
        'DiscreteChannel is expected to be recieved as:
        ' 1: Weight on Wheels
        ' 2: Eng 1 Oil Pressure
        ' 3: Eng 2 Oil Pressure
        ' Numbers above correspond to labels on box (in simulator) and labels on AID Status web page
        ' As per testing, the actual number that the API dll is expecting is zero-based, and so numbers are reduced by one below
        Dim SuccessStatus As Boolean
        Dim outChannel As Byte
        Dim outState As Boolean
        Dim outTimeStamp As UInt32
        outChannel = DiscreteChannel - 1
        SuccessStatus = DatabusClass.navAPI_GetAIDDiscreteIn(outChannel, outState, outTimeStamp)
        If SuccessStatus = True Then
            DiscreteStatus = outState
        Else
            'DiscreteStatus = False
            DiscreteStatus = Nothing
            'Call WriteToLog(False, "DiscreteStatus " & DiscreteChannel & " Is Nothing", False)
        End If
    End Function


    ' Rn Commented Out. Have set up timer instead of thread
    ' ''' <summary>
    ' ''' The method performs the main function of the service. It runs on a 
    ' ''' thread pool worker thread.
    ' ''' </summary>
    ' ''' <param name="state"></param>
    Private Sub ServiceWorkerThread(ByVal state As Object)
        ' Periodically check if the service is stopping.
        Do While Not Me.stopping
            ' Perform main service function here...

            Thread.Sleep(2000)  ' Simulate some lengthy operations.
            'Me.EventLog1.WriteEntry("Thread Ticked")       'This line works indefinately
            'Call WriteToLog("Thread has ticked", False)    'This line works onece or twice, then stops working

        Loop

        ' Signal the stopped event.
        Me.stoppedEvent.Set()
        'Call WriteToLog("Thread has stopped", False)
        Me.EventLog1.WriteEntry("Thread has stopped")
    End Sub


    ''' <summary>
    ''' The function is executed when a Stop command is sent to the service 
    ''' by SCM. It specifies actions to take when a service stops running. In 
    ''' this code sample, OnStop logs a service-stop message to the 
    ''' Application log, and waits for the finish of the main service 
    ''' function.
    ''' </summary>
    Protected Overrides Sub OnStop()
        ' Log a service stop message to the Application log.
        Me.EventLog1.WriteEntry("EfbPerfDataMonitor In OnStop.")

        'Call WriteToLog("+----------+----------+----------+----------+")
        Call WriteToLog("Service Is stopping")
        Call WriteToLog("****************************************************************************")
        LogStream.Close()
        LogStream.Dispose()

        ' Indicate that the service is stopping and wait for the finish of 
        ' the main service function (ServiceWorkerThread).
        'Me.stopping = True

        'Rn Commented out, as I think this is associated with thread that I'm not using
        'Me.stoppedEvent.WaitOne()
    End Sub

    Protected Overrides Sub OnShutdown()
        Call WriteToLog("Device Is Shutting Down")
        Call OnStop()
    End Sub

    Private Sub WriteToLog(Message As String)
        LogStream.WriteLine(Format(Now, "MMdd-HH:mm:ss") & " " & Message)
    End Sub


    Private Function DiscreteDecode(navAeroValue As Boolean?) As String
        If navAeroValue Is Nothing Then
            DiscreteDecode = "Null"
        Else
            If navAeroValue = True Then
                DiscreteDecode = "No  (T)"
            ElseIf navAeroValue = False Then
                DiscreteDecode = "Yes (F)"
            Else
                DiscreteDecode = "Unknown"
            End If
        End If

    End Function

End Class

Public Class A429DataSet
    Public PacketData As UInt32
    Public TimeStamp As UInt32
    Public Success As Boolean
    Public Age As Single 'seconds
End Class

Public Class A717SuperFrame
    Public Frames(15) As A717Frame
End Class

Public Class A717Frame
    Public SubFrames(4) As A717SubFrame
    Public FrameCounter As Integer
    Public WordCount As Integer
End Class

Public Class A717SubFrame
    Public WordsRaw() As UInt32
    Public Timestamp As UInt32
    Public SuccessStatus As Boolean
    Public outFrameSize As Integer
    Public Age As Single
End Class

Public Class DecodedPacket
    Public LabelNumber As Integer
    Public DecodedPayload As Single
    Public DecodedTime As DateTime
    Public DecodedDate As DateTime
    Public DecodedSignStatus As DatabusClass.SignStatus
    Public ConsideredValid As Boolean 'Specifies if payload is valid, according to Sign Status value
    Public Success As Boolean 'will be set to false if a conversion of a payload to a time or date fails
End Class

Public Class LabelInstance
    '717 characteristics
    Public Cycle As Integer
    Public SubFrame As Integer
    Public Word As Integer
    Public Msb As Integer
    Public Lsb As Integer
    Public SignBit As Integer
    '429 characteristics
    Public Channel As Integer
    Public Label As Integer
    Public SigBits As Integer
    Public DecodeMechanism As DatabusClass.DecodeMechanism
    'Common characteristics
    Public Description As String
    Public LabelType As DatabusClass.LabelType
    Public Encoding As DatabusClass.Encoding
    Public Scale As Integer
End Class

Public Class DatabusClass
    Public Declare Function navAPI_Initialize Lib "C:\navaero\navApi\64bit\PerfDataMonitorService\navAPI.dll" () As Boolean
    Public Declare Function navAPI_Release Lib "C:\navaero\navApi\64bit\PerfDataMonitorService\navAPI.dll" () As Boolean
    Public Declare Function navAPI_GetVersion Lib "C:\navaero\navApi\64bit\PerfDataMonitorService\navAPI.dll" (ByRef versionmajor As Byte, ByRef versionminor As Byte, ByRef configuration As UInteger) As Boolean
    Public Declare Function navAPI_GetPDIPositionID Lib "C:\navaero\navApi\64bit\PerfDataMonitorService\navAPI.dll" (ByRef posid As UInt32, ByRef timestamp As UInt32) As Boolean
    Public Declare Function navAPI_GetAIDDiscreteIn Lib "C:\navaero\navApi\64bit\PerfDataMonitorService\navAPI.dll" (ByVal channel As Byte, ByRef state As Boolean, ByRef timestamp As UInt32) As Boolean
    Public Declare Function navAPI_GetAIDDiscreteLogical Lib "C:\navaero\navApi\64bit\PerfDataMonitorService\navAPI.dll" (ByVal channel As Byte, ByRef state As Boolean, ByRef timestamp As UInt32) As Boolean
    Public Declare Function navAPI_GetA429Label Lib "C:\navaero\navApi\64bit\PerfDataMonitorService\navAPI.dll" (ByVal label As Byte, ByVal channel As Byte, ByRef A429data As UInt32, ByRef timestamp As UInt32) As Boolean
    Public Declare Function navAPI_GetA717SubFrame Lib "C:\navaero\navApi\64bit\PerfDataMonitorService\navAPI.dll" (ByRef buffer As UInt32, ByVal buffersize As UInt32, ByRef timestamp As UInt32, ByRef framesize As UInt32) As Boolean

    Public Shared WoW As Boolean? '= True 'Weight on Wheels
    Public Shared Eng1Lop As Boolean? '= True 'Engine 1 Oil Pressure
    Public Shared Eng2Lop As Boolean? '= True 'Engine 2 Oil Pressure
    Public Shared PreviousWoW As Boolean? '= True 'Weight on Wheels : True indicates No Weight on Wheels
    Public Shared PreviousEng1Lop As Boolean? '= True 'Engine 1 Low Oil Pressure : True indicates No Low Oil Pressure (i.e. Eng Running)
    Public Shared PreviousEng2Lop As Boolean? '= True 'Engine 2 Low Oil Pressure : True indicates No Low Oil Pressure (i.e. Eng Running)
    Public Shared navApiStatus As Boolean
    'Public valueMonitorTimer As System.Timers.Timer

    Public Shared LabelCount As Integer

    Public Shared LabelList() As LabelInstance
    Public Shared A429LabelList As Integer() = {1, 2, 3, 4, 5, 6, 7, 10, 11, 12, 13, 14, 15, 16, 17, 20, 21, 22, 23,
                                         24, 25, 26, 27, 30, 31, 32, 33, 34, 35, 36, 37, 40, 41,
                                         42, 43, 44, 45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 60, 61, 62, 63, 64,
                                         65, 66, 67, 70, 71, 72, 73, 74, 75, 76,
                                         77, 100, 101, 102, 103, 104,
                                         105, 106, 107, 110, 111, 112, 113, 114, 115, 116,
                                         117, 120, 121, 122, 123, 124, 125, 126, 127, 130, 131, 132, 133,
                                         134, 135, 136, 137, 140, 141, 142, 143, 144, 145, 146, 147, 150,
                                         151, 152, 153, 154, 155, 156, 157,
                                         160, 161, 162, 163, 164, 165, 166, 167, 170, 171, 172, 173,
                                         174, 175, 176, 177, 200, 201, 202, 203, 204,
                                         205, 206, 207, 210, 211, 212, 213, 214, 215, 216,
                                         217, 220, 221, 222, 223, 224, 225, 226, 227, 230, 231, 232, 233, 234,
                                         235, 236, 237, 240, 241, 242, 243, 244,
                                         245, 246, 247, 250, 251, 252, 253, 254,
                                         255, 256, 257, 260, 261,
                                         262, 263, 264, 265, 266,
                                         267, 270, 271, 272, 273, 274, 275,
                                         276, 277, 300, 301, 302, 303, 304,
                                         305, 306, 307, 310, 311, 312, 313, 314,
                                         315, 316, 317, 320, 321, 322,
                                         323, 324, 325, 326, 327, 330, 331,
                                         332, 333, 334, 335, 336, 337, 340,
                                         341, 342, 343, 344, 345, 346, 347, 350,
                                         351, 352, 353, 354, 355,
                                         356, 357, 360, 361, 362, 363, 364, 365, 366, 367, 370, 371,
                                         372, 373, 374, 375, 376, 377}

    Public Enum DecodeMechanism
        Table609FlightNumber '
        Table618Date        '1
        Table625General
        Table625Time
    End Enum

    Public Enum Encoding
        BCD
        BNR
    End Enum

    Public Enum LabelType
        A429
        A717
    End Enum

    Public Enum SignStatus
        PlusNorth
        NoComputedData
        FunctionalTest
        MinusSouth
        FailureWarning
        NormalOperation
        Unknown
    End Enum

    Public Enum OctalCoding
        To429Oct
        Fr429Oct
    End Enum

    Public Shared Function Get717SubFrame(ByVal WordCount As Integer) As A717SubFrame
        Dim ds As New A717SubFrame
        Dim SuccessStatus As Boolean

        'Dim outBuffer(WordCount) As UInt32
        ReDim ds.WordsRaw(WordCount)
        Dim outTimestamp As UInt32
        Dim outFramesize As UInt32
        Dim Temp As String
        Dim i As Integer

        Try
            SuccessStatus = DatabusClass.navAPI_GetA717SubFrame(ds.WordsRaw(0), WordCount * 4, outTimestamp, outFramesize)
        Catch ex As Exception
            'MsgBox(ex.Message)
        End Try
        'Shift all Words upwards by one. This is required becasue navAero populate a zero-based sequence, whereas Airbus refer to them as one-based
        For i = WordCount To 1 Step -1
            ds.WordsRaw(i) = ds.WordsRaw(i - 1)
        Next
        ds.SuccessStatus = SuccessStatus
        ds.Timestamp = outTimestamp
        ds.outFrameSize = outFramesize
        ds.Age = (Environment.TickCount - outTimestamp) / 1000
        Return ds
    End Function

    Public Shared Function Get429RawData(LabelNumber As Integer, Channel As Integer) As A429DataSet
        Dim ds = New A429DataSet
        Dim SuccessStatus As Boolean

        Dim outA429Data As UInt32
        Dim outTimestamp As UInt32

        Dim Divisor As Integer = 2
        Dim Result As Double
        Dim BitResult As Single
        Try
            SuccessStatus = navAPI_GetA429Label(Convert.ToByte(LabelNumber, 8), Convert.ToByte(Channel, 16), outA429Data, outTimestamp)
            'ds.PacketData = UCase(Convert.ToString(outA429Data, 16))
            ds.PacketData = outA429Data
            'ds.Age = Convert.ToString(outTimestamp, 10)
            ds.TimeStamp = outTimestamp
            ds.Success = SuccessStatus
            ds.Age = (Environment.TickCount - outTimestamp) / 1000
        Catch ex As Exception

        End Try
        Return ds
    End Function

    Public Shared Function TwosComplement(rString As String, sb As Integer, Raw As UInt32) As String
        Dim i As Integer
        Debug.WriteLine("Incoming = " & rString)


        'Put the string into an array, backwards. This means that the array indexs will correspond to the bit positions (which are backwards ... i.e. they go from 32 to 1 reading left to right)
        Dim rArray(32) As String
        For i = 1 To 32
            rArray(33 - i) = Mid(rString, i, 1)
        Next

        'Least Significant Bit
        Dim lsb As Integer = 29 - sb
        'ListBox1.Items.Add("Lsb = " & lsb)

        'Invert the Bit Positions that contain payload
        For i = lsb To 28
            If rArray(i) = 1 Then
                rArray(i) = 0
            ElseIf rArray(i) = 0 Then
                rArray(i) = 1
            End If
        Next
        Dim invertedString As String
        For i = 1 To 32
            invertedString = invertedString & rArray(33 - i)
        Next

        'Add Binary 1 to the payload
        For i = lsb To 28
            If rArray(i) = 0 Then
                rArray(i) = 1
                Exit For
            ElseIf rArray(i) = 1 Then
                rArray(i) = 0
            End If
        Next

        'Reassemble array into string
        Dim resultString As String
        For i = 1 To 32
            resultString = resultString & rArray(33 - i)
            'rArray(33 - i) = Mid(rString, i, 1)
        Next
        'ListBox1.Items.Add("33322222222221111111111")
        'ListBox1.Items.Add("21098765432109876543210987654321")
        'ListBox1.Items.Add(rString)
        'ListBox1.Items.Add(invertedString)
        'ListBox1.Items.Add(resultString)

        Debug.WriteLine("Two's C = " & resultString)
        TwosComplement = resultString
    End Function


    Public Shared Function TwosComplement717B(rString As String) As String
        Dim i As Integer
        Debug.WriteLine("Incoming = " & rString)

        'The string that's been recieved is:
        ' - only the payload bits, excluding the sign bit

        'Invert the digits
        For i = 1 To Len(rString)
            If Mid(rString, i, 1) = 0 Then
                Mid(rString, i, 1) = 1
            ElseIf Mid(rString, i, 1) = 1 Then
                Mid(rString, i, 1) = 0
            End If
        Next
        Debug.WriteLine("Inverted = " & rString)

        'Add Binary 1 to the payload
        For i = Len(rString) To 1 Step -1
            If Mid(rString, i, 1) = 0 Then
                'The fistr occurrance of a 0 is changed to 1. Once this is done, exit the Loop
                Mid(rString, i, 1) = 1
                Exit For
            ElseIf Mid(rString, i, 1) = 1 Then
                'Each occurrance of a 1 is changed to a 0
                Mid(rString, i, 1) = 0
            End If
        Next

        Debug.WriteLine("Two's C  = " & rString)
        TwosComplement717B = rString
    End Function


    Public Shared Function BnrDecode(BnrInput As UInt32, dm As DecodeMechanism, SigBits As Integer, Scale As Integer) As DecodedPacket
        'Bit Posn:    | 32 | 31 | 30 | 29 | 28 | 27 | 26 | 25 | 24 | 23 | 22 | 21 | 20 | 19 | 18 | 17 | 16 | 15 | 14 | 13 | 12 | 11 | 10 |  9 |
        'String Posn: |  1 |  2 |  3 |  4 |  5 |  6 |  7 |  8 |  9 | 10 | 11 | 12 | 13 | 14 | 15 | 16 | 17 | 18 | 19 | 20 | 21 | 22 | 23 | 24 |
        'Sig Bits:                        |  1 |  2 |  3 |  4 |  5 |  6 |  7 |  8 |  9 | 10 | 11 | 12 | 13 | 14 | 15 | 16 | 17 | 18 | 19 | 20 |
        Dim fs As New DecodedPacket
        Dim rnString As String
        Dim Divisor As Integer = 2
        Dim Result As Double
        Dim BitResult As Single

        'ListBox1.Items.Clear()
        'lblRaw.Text = outA429Data.ToString
        'lblHex.Text = UCase(Convert.ToString(outA429Data, 16))
        'lblBinary.Text = Convert.ToString(outA429Data, 2)

        rnString = Convert.ToString(BnrInput, 2)
        rnString = rnString.PadLeft(32, "0")
        Debug.WriteLine("rnString = " & rnString)

        Select Case Mid(rnString, 2, 2)
            Case "00"
                fs.DecodedSignStatus = SignStatus.FailureWarning
                fs.ConsideredValid = False
            Case "01"
                fs.DecodedSignStatus = SignStatus.NoComputedData
                fs.ConsideredValid = False
            Case "10"
                fs.DecodedSignStatus = SignStatus.FunctionalTest
                fs.ConsideredValid = False
            Case "11"
                fs.DecodedSignStatus = SignStatus.NormalOperation
                fs.ConsideredValid = True
            Case Else
                fs.DecodedSignStatus = SignStatus.Unknown
                fs.ConsideredValid = False
        End Select

        'Bit posn 29 determnines if it's a negative or positive number
        '1 = Negative Number; Two's Complement to be applied
        '0 = Positive Number; no additional handling required
        'Bit posn 29 is 4th Character
        If Mid(rnString, 4, 1) = "1" Then
            rnString = DatabusClass.TwosComplement(rnString, SigBits, BnrInput)
            Debug.WriteLine("Two's Comp rnString = " & rnString)
        End If

        'Put the string into an array, backwards. This means that the array indexs will correspond to the bit positions (which are backwards ... i.e. they go from 32 to 1 reading left to right)
        Dim sArray(32) As String
        For i = 1 To 32
            sArray(33 - i) = Mid(rnString, i, 1)
        Next

        'Least Significant Bit
        Dim lsb As Integer = 29 - SigBits
        'ListBox1.Items.Add("Lsb = " & lsb)

        For i = 28 To lsb Step -1
            BitResult = Convert.ToSingle(sArray(i)) / Divisor * Scale
            Debug.WriteLine("BitResult = " & BitResult & "(Value=" & sArray(i) & ", Divisor=" & Divisor & ")")
            Result = Result + BitResult
            Divisor = Divisor * 2
        Next
        fs.DecodedPayload = Result
        Debug.WriteLine("Final Result = " & Result)
        Return fs
    End Function

    Public Shared Function BnrDecode717B(BnrInput As UInt32, LSB As Integer, MSB As Integer, SignBit As Integer, Scale As Integer) As DecodedPacket
        '                        MSB                                                    LSB
        '         Bit Posn:    | 12 | 11 | 10 |  9 |  8 |  7 |  6 |  5 |  4 |  3 |  2 |  1 |
        '         String Posn: |  1 |  2 |  3 |  4 |  5 |  6 |  7 |  8 |  9 | 10 | 11 | 12 |
        'Reversed String Posn: | 12 | 11 | 10 |  9 |  8 |  7 |  6 |  5 |  4 |  3 |  2 |  1 |

        Dim fs As New DecodedPacket
        Dim rnString As String
        Dim rnPayLoad As String
        Dim IsNegative As Boolean

        Dim Divisor As Integer = 2
        Dim Result As Double
        Dim BitResult As Single

        Debug.WriteLine("Input  = " & BnrInput)

        rnString = Convert.ToString(BnrInput, 2)
        rnString = rnString.PadLeft(12, "0")
        Debug.WriteLine("rnString            = " & rnString)
        rnPayLoad = Mid(rnString, 12 - MSB + 1, MSB - LSB + 1)
        Debug.WriteLine("Payload in " & LSB & " to " & MSB)
        Debug.WriteLine("rnPayLoad = " & rnPayLoad)

        'The value in the SignBit location determines whether it's a positive or negative value
        'If SignBit  has been passed as 0, then this check doesn't apply, and so IsNegative is set to False
        '1 = Negative Number; Two's Complement to be applied
        '0 = Positive Number; no additional handling required
        If SignBit > 0 Then
            If Mid(rnString, 12 - SignBit + 1, 1) = "1" Then
                Debug.WriteLine("Applying Two's Complement")
                IsNegative = True
                rnPayLoad = DatabusClass.TwosComplement717B(rnPayLoad)
                'Debug.WriteLine("Two's Comp rnString = " & rnString)
            Else
                IsNegative = False
            End If
        Else
            IsNegative = False
            Debug.WriteLine("No sign bit specified; IsNegative set to False")
        End If
        Debug.WriteLine("IsNegative = " & IsNegative.ToString)

        For i = 1 To Len(rnPayLoad)
            BitResult = Convert.ToSingle(Val(Mid(rnPayLoad, i, 1))) / Divisor * Scale
            Debug.WriteLine("BitResult = " & BitResult & "(Value=" & Val(Mid(rnPayLoad, i, 1)) & ", Divisor=" & Divisor & ")")
            Result = Result + BitResult
            Divisor = Divisor * 2
        Next

        If IsNegative = True Then Result = Result * -1

        fs.DecodedPayload = Result
        Debug.WriteLine("Final Result = " & Result)
        'Call EfbPerfDataMonitorService.WriteToLog("717B: " & )
        Return fs
    End Function

    Public Shared Function BnrDecode429B(BnrInput As UInt32, dm As DecodeMechanism, SigBits As Integer, Scale As Integer) As DecodedPacket
        'Bit Posn:    | 32 | 31 | 30 | 29 | 28 | 27 | 26 | 25 | 24 | 23 | 22 | 21 | 20 | 19 | 18 | 17 | 16 | 15 | 14 | 13 | 12 | 11 | 10 |  9 |
        'String Posn: |  1 |  2 |  3 |  4 |  5 |  6 |  7 |  8 |  9 | 10 | 11 | 12 | 13 | 14 | 15 | 16 | 17 | 18 | 19 | 20 | 21 | 22 | 23 | 24 |
        'Sig Bits:                        |  1 |  2 |  3 |  4 |  5 |  6 |  7 |  8 |  9 | 10 | 11 | 12 | 13 | 14 | 15 | 16 | 17 | 18 | 19 | 20 |
        Dim fs As New DecodedPacket
        Dim rnString As String

        Dim Divisor As Integer = 2
        Dim Result As Double
        Dim BitResult As Single
        Dim rnPayload As String
        Dim IsNegative As Boolean

        rnString = Convert.ToString(BnrInput, 2)
        rnString = rnString.PadLeft(32, "0")
        Debug.WriteLine("rnString = " & rnString)
        'Payload will be from Bit 28, with a Length as passed in SigBits
        rnPayload = Mid(rnString, 5, SigBits)
        Debug.WriteLine("rnPayLoad = " & rnPayload)

        'The value in the Bit 29 determines whether it's a positive or negative value
        'Bit 29 is String posn 4
        '1 = Negative Number; Two's Complement to be applied
        '0 = Positive Number; no additional handling required
        If Mid(rnString, 4, 1) = "1" Then
            Debug.WriteLine("Applying Two's Complement")
            IsNegative = True
            rnPayload = DatabusClass.TwosComplement717B(rnPayload)
            'Debug.WriteLine("Two's Comp rnString = " & rnString)
        Else
            IsNegative = False
        End If
        Debug.WriteLine("IsNegative = " & IsNegative.ToString)



        Select Case Mid(rnString, 2, 2)
            Case "00"
                fs.DecodedSignStatus = SignStatus.FailureWarning
                fs.ConsideredValid = False
            Case "01"
                fs.DecodedSignStatus = SignStatus.NoComputedData
                fs.ConsideredValid = False
            Case "10"
                fs.DecodedSignStatus = SignStatus.FunctionalTest
                fs.ConsideredValid = False
            Case "11"
                fs.DecodedSignStatus = SignStatus.NormalOperation
                fs.ConsideredValid = True
            Case Else
                fs.DecodedSignStatus = SignStatus.Unknown
                fs.ConsideredValid = False
        End Select

        'Least Significant Bit
        Dim lsb As Integer = 29 - SigBits
        'ListBox1.Items.Add("Lsb = " & lsb)

        For i = 1 To Len(rnPayload)
            BitResult = Convert.ToSingle(Val(Mid(rnPayload, i, 1))) / Divisor * Scale
            Debug.WriteLine("BitResult = " & BitResult & "(Value=" & Val(Mid(rnPayload, i, 1)) & ", Divisor=" & Divisor & ")")
            Result = Result + BitResult
            Divisor = Divisor * 2
        Next

        If IsNegative = True Then Result = Result * -1

        fs.DecodedPayload = Result
        Debug.WriteLine("Final Result = " & Result)
        Return fs
    End Function


    Public Shared Function BcdDecode(BcdInput As UInt32, dm As DecodeMechanism, SigBits As Integer) As DecodedPacket
        Dim fs As New DecodedPacket

        Dim rn As String
        Dim rnPadded As String
        Dim Digit1 As Int32
        Dim Digit2 As Int32
        Dim Digit3 As Int32
        Dim Digit4 As Int32
        Dim Digit5 As Int32
        Dim Digit6 As Int32

        'ListBox1.Items.Clear()

        rn = Convert.ToString(BcdInput, 2)
        rnPadded = rn.PadLeft(32, "0")
        'MsgBox(BcdInput & vbCrLf & rn & vbCrLf & rnPadded & vbCrLf)
        'ListBox1.Items.Add("33322222222221111111111")
        'ListBox1.Items.Add("21098765432109876543210987654321")

        'ListBox1.Items.Add(rnPadded)
        Select Case Mid(rnPadded, 2, 2)
            Case "00"
                fs.DecodedSignStatus = SignStatus.PlusNorth
                fs.ConsideredValid = True
            Case "01"
                fs.DecodedSignStatus = SignStatus.NoComputedData
                fs.ConsideredValid = False
            Case "10"
                fs.DecodedSignStatus = SignStatus.FunctionalTest
                fs.ConsideredValid = False
            Case "11"
                fs.DecodedSignStatus = SignStatus.MinusSouth
                fs.ConsideredValid = True
            Case Else
                fs.DecodedSignStatus = SignStatus.Unknown
                fs.ConsideredValid = False
        End Select

        Select Case dm
            Case DecodeMechanism.Table609FlightNumber

            Case DecodeMechanism.Table618Date
                'Digit1 - Day x 10
                'Bits 29-28
                'Mid: Start 4, Length 2
                'MsgBox(Mid(rnPadded, 4, 2))
                Digit1 = Convert.ToInt32(Mid(rnPadded, 4, 2), 2)
                'MsgBox(Digit1)

                'Digit2 - Day x 1
                'Bits 27-26-25-24
                'Mid: Start 6, Length 4
                'MsgBox(Mid(rnPadded, 6, 4))
                Digit2 = Convert.ToInt32(Mid(rnPadded, 6, 4), 2)
                'MsgBox(Digit2)

                'Digit3 - Month x 10
                'Bits 23
                'Mid: Start 10, Length 1
                'MsgBox(Mid(rnPadded, 10, 1))
                Digit3 = Convert.ToInt32(Mid(rnPadded, 10, 1), 2)
                'MsgBox(Digit3)

                'Digit4 - Month x 1
                'Bits 22-21-20-19
                'Mid: Start 11, Length 4
                'MsgBox(Mid(rnPadded, 11, 4))
                Digit4 = Convert.ToInt32(Mid(rnPadded, 11, 4), 2)
                'MsgBox(Digit4)

                'Digit5 - Year x 10
                'Bits 18-17-16-15
                'Mid: Start 15, Length 4
                'MsgBox(Mid(rnPadded, 15, 4))
                Digit5 = Convert.ToInt32(Mid(rnPadded, 15, 4), 2)
                'MsgBox(Digit5)

                'Digit6 - Year x 1
                'Bits 14-13-12-11
                'Mid: Start 19, Length 4
                'MsgBox(Mid(rnPadded, 19, 4))
                Digit6 = Convert.ToInt32(Mid(rnPadded, 19, 4), 2)
                'MsgBox(Digit6)

                'MsgBox("Day: " & Digit1 * 10 + Digit2 & vbCrLf & "Month: " & Digit3 * 10 + Digit4 & vbCrLf & "Year: " & Digit5 * 10 + Digit6)
                Try
                    fs.DecodedDate = New DateTime(Convert.ToInt32(Digit5 * 10 + Digit6 + 2000), Convert.ToInt32(Digit3 * 10 + Digit4), Convert.ToInt32(Digit1 * 10 + Digit2), 0, 0, 0)
                    fs.Success = True
                Catch ex As Exception
                    'the decoded values can't be converted to a date
                    fs.Success = False
                End Try
            Case DecodeMechanism.Table625General, DecodeMechanism.Table625Time
                'Digit1
                'Bits 29-28-27
                'Mid: Start 4, Length 3
                'MsgBox(Mid(rnPadded, 4, 3))
                Digit1 = Convert.ToInt32(Mid(rnPadded, 4, 3), 2)
                Debug.WriteLine("Digit 1 :" & Digit1)
                'MsgBox(Digit1)

                'Digit2
                'Bits 26-25-24-23
                'Mid: Start 7, Length 4
                'MsgBox(Mid(rnPadded, 7, 4))
                Digit2 = Convert.ToInt32(Mid(rnPadded, 7, 4), 2)
                Debug.WriteLine("Digit 2 :" & Digit2)
                'MsgBox(Digit2)

                'Digit3
                'Bits 22-21-20-19
                'Mid: Start 11, Length 4
                'MsgBox(Mid(rnPadded, 11, 4))
                Digit3 = Convert.ToInt32(Mid(rnPadded, 11, 4), 2)
                Debug.WriteLine("Digit 3 :" & Digit3)
                'MsgBox(Digit3)

                'Digit4
                'Bits 18-17-16-15
                'Mid: Start 15, Length 4
                'MsgBox(Mid(rnPadded, 15, 4))
                Digit4 = Convert.ToInt32(Mid(rnPadded, 15, 4), 2)
                Debug.WriteLine("Digit 4 :" & Digit4)
                'MsgBox(Digit4)

                'Digit5
                'Bits 14-13-12-11
                'Mid: Start 19, Length 4
                'MsgBox(Mid(rnPadded, 19, 4))
                Digit5 = Convert.ToInt32(Mid(rnPadded, 19, 4), 2)
                Debug.WriteLine("Digit 5 :" & Digit5)
                'MsgBox(Digit5)

                Select Case dm
                    Case DecodeMechanism.Table625General
                        Try
                            fs.DecodedPayload = Convert.ToSingle(Digit1 & Digit2 & Digit3 & Digit4 & Digit5)
                            fs.Success = True
                        Catch ex As Exception
                            fs.Success = False
                        End Try
                    Case DecodeMechanism.Table625Time
                        Try
                            fs.DecodedTime = New DateTime(1, 1, 1, Convert.ToInt32(Digit1 & Digit2), Convert.ToInt32(Digit3 & Digit4), Convert.ToInt32(Digit5) * 6)
                            fs.Success = True
                        Catch ex As Exception
                            fs.Success = False
                        End Try
                End Select
            'BcdDecode = Convert.ToSingle(Digit1 & Digit2 & Digit3 & Digit4 & Digit5)
            Case Else
        End Select
        'BcdToDigital = Result
        Return fs
    End Function

    Public Shared Function DiscreteStatus(DiscreteChannel As Integer) As Boolean?
        'DiscreteChannel is expected to be recieved as:
        ' 1: Weight on Wheels
        ' 2: Eng 1 Oil Pressure
        ' 3: Eng 2 Oil Pressure
        ' Numbers above correspond to labels on box (in simulator) and labels on AID Status web page
        ' As per testing, the actual number that the API dll is expecting is zero-based, and so numbers are reduced by one below
        Dim SuccessStatus As Boolean
        Dim outChannel As Byte
        Dim outState As Boolean
        Dim outTimeStamp As UInt32
        outChannel = DiscreteChannel - 1
        SuccessStatus = navAPI_GetAIDDiscreteIn(outChannel, outState, outTimeStamp)
        If SuccessStatus = True Then
            DiscreteStatus = outState
        Else
            'DiscreteStatus = False
            DiscreteStatus = Nothing
            'Call WriteToLog(False, "DiscreteStatus " & DiscreteChannel & " is Nothing", False)
        End If
    End Function

End Class

