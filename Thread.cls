VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Thread"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Not unique to External or Internal Thread Properties
Private m_Name As String
Private m_Allowance As Double
Private m_FractionalThreadName As String

'External Thread Properties
Private m_ExtThreadClass As String
Private m_ExtMajorDMax As Double
Private m_ExtMajorDMin As Double
Private m_ExtMajorDMinNote3 As Double
Private m_ExtPitchDMax As Double
Private m_ExtPitchDmin As Double
Private m_ExtPitchDTol As Double
Private m_ExtUNRMinorDMaxRef As Double
Private m_ExtMinorThreadDMin As Double
Private m_ExtMinorThreadDMax As Double

'Internal Thread Properties
Private m_IntThreadClass As String
Private m_IntMinorDMin As Double
Private m_IntMinorDMax As Double
Private m_IntPitchDmin As Double
Private m_IntPitchDMax As Double
Private m_IntPitchDTol As Double
Private m_IntMajorDMin As Double


Public Sub InitiateThread(ThreadRow() As Variant)
    'This sub is just used to initialize a thread.  The numbers are the respective columns in the thread sheet
    m_Name = ThreadRow(1)
    m_FractionalThreadName = ThreadRow(2)
    m_ExtThreadClass = ThreadRow(6)
    m_Allowance = ThreadRow(7)
    m_ExtMajorDMax = ThreadRow(8)
    m_ExtMajorDMin = ThreadRow(9)
    m_ExtMajorDMinNote3 = ThreadRow(10)
    m_ExtPitchDMax = ThreadRow(11)
    m_ExtPitchDmin = ThreadRow(12)
    m_ExtPitchDTol = ThreadRow(13)
    m_ExtUNRMinorDMaxRef = ThreadRow(14)
    m_IntThreadClass = ThreadRow(15)
    m_IntMinorDMin = ThreadRow(16)
    m_IntMinorDMax = ThreadRow(17)
    m_IntPitchDmin = ThreadRow(18)
    m_IntPitchDMax = ThreadRow(19)
    m_IntPitchDTol = ThreadRow(20)
    m_IntMajorDMin = ThreadRow(21)
    
    'Need to check that these minor diameter calculations are correct
    m_ExtMinorThreadDMin = 2 * m_ExtPitchDmin - m_ExtMajorDMax
    m_ExtMinorThreadDMax = 2 * m_ExtPitchDMax - m_ExtMajorDMin
    
End Sub


Property Get ExtMinorThreadDMin() As Double
    ExtMinorThreadDMin = m_ExtMinorThreadDMin
End Property

Property Get ExtMinorThreadDMax() As Double
    ExtMinorThreadDMax = m_ExtMinorThreadDMax
End Property

Property Get Name() As String
    Name = m_Name
End Property

Property Get ExtThreadClass() As String
    ExtThreadClass = m_ExtThreadClass
End Property

Property Get Allowance() As Double
    Allowance = m_Allowance
End Property

Property Get ExtMajorDMax() As Double
    ExtMajorDMax = m_ExtMajorDMax
End Property

Property Get ExtMajorDMin() As Double
    ExtMajorDMin = m_ExtMajorDMin
End Property

Property Get ExtMajorDMinNote3() As Double
    ExtMajorDMinNote3 = m_ExtMajorDMinNote3
End Property

Property Get ExtPitchDMax() As Double
    ExtPitchDMax = m_ExtPitchDMax
End Property

Property Get ExtPitchDmin() As Double
    ExtPitchDmin = m_ExtPitchDmin
End Property

Property Get ExtPitchDTol() As Double
    ExtPitchDTol = m_ExtPitchDTol
End Property

Property Get ExtUNRMinorDMaxRef() As Double
    ExtUNRMinorDMaxRef = m_ExtUNRMinorDMaxRef
End Property

Property Get IntThreadClass() As String
    IntThreadClass = m_IntThreadClass
End Property

Property Get IntMinorDMin() As Double
    IntMinorDMin = m_IntMinorDMin
End Property

Property Get IntMinorDMax() As Double
    IntMinorDMax = m_IntMinorDMax
End Property

Property Get IntPitchDmin() As Double
    IntPitchDmin = m_IntPitchDmin
End Property

Property Get IntPitchDMax() As Double
    IntPitchDMax = m_IntPitchDMax
End Property

Property Get IntPitchDTol() As Double
    IntPitchDTol = m_IntPitchDTol
End Property

Property Get IntMajorDMin() As Double
    IntMajorDMin = m_IntMajorDMin
End Property

Property Get Pitch() As Double
    Pitch = 1 / CDbl(Split(m_FractionalThreadName, "-")(1))
End Property

Property Get InternalThreadGeometry(MaterialCondition As String, Depth As Double, LeftOrRight As String) As Variant
    Dim GeometryArray As Variant
    Dim MinorDiameter As Double
    Dim BoltScrewLength As Double
    
    
    If MaterialCondition = "MMC" Then
        MinorDiameter = m_ExtMinorThreadDMax
        BoltScrewLength = m_ScrewLengthMax
    ElseIf MaterialCondition = "LMC" Then
        MinorDiameter = m_Thread.ExtMinorThreadDMin
        BoltScrewLength = m_ScrewLengthMin
    ElseIf MaterialCondition = "Average" Then
        MinorDiameter = (m_Thread.ExtMinorThreadDMin + m_Thread.ExtMinorThreadDMax) / 2
        BoltScrewLength = (m_ScrewLengthMax + m_ScrewLengthMin) / 2
    End If
    
    
    If LeftOrRight = "Left" Then
        GeometryArray = Array( _
                                    Array(-MinorDiameter / 2, -BoltScrewLength + m_ThreadedLengthMin), _
                                    Array(-MinorDiameter / 2, -BoltScrewLength))
    ElseIf LeftOrRight = "Right" Then
        GeometryArray = Array( _
                                    Array(MinorDiameter / 2, -BoltScrewLength + m_ThreadedLengthMin), _
                                    Array(MinorDiameter / 2, -BoltScrewLength))
    End If
    
    ThreadGeometryGenerator = Application.Index(GeometryArray, 0, 0)
    
End Property
