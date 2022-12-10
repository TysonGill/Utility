Imports System.Threading
Imports System.Threading.Tasks

''' <summary>Class for reporting program resource utilization.</summary>
Public Class clsRescDiag

    ''' <summary>The sample size for averaging resource levels.</summary>
    Public SampleSize As Integer = 10

    ''' <summary>The delete in milliseconds between resource sample measurements.</summary>
    Public SamplingRate As Integer = 100

    ''' <summary>The sample log.</summary>
    Public Log As DataTable = Nothing

    Private QueueCPU As New Queue
    Private LoggingEnabled As Boolean = False
    Private swLog As New Stopwatch
    Private CPUCounter As System.Diagnostics.PerformanceCounter = New System.Diagnostics.PerformanceCounter

    ''' <summary>Instantiate a new SysResources class.</summary>
    ''' <remarks>Initiates the CPU sampling background process</remarks>
    Public Sub New()

        CPUCounter.CategoryName = "Processor"
        CPUCounter.CounterName = "% Processor Time"
        CPUCounter.InstanceName = "_Total"
        CPUCounter.NextValue()
        Dim TaskCPU As New task(Sub()
                                    Dim Sample As Single = 0
                                    Do
                                        Dim LogLock As New Object
                                        Sample = CPUCounter.NextValue()
                                        If QueueCPU.Count >= SampleSize Then QueueCPU.Dequeue()
                                        QueueCPU.Enqueue(Sample)
                                        If LoggingEnabled AndAlso Log.Rows.Count < 10000 Then
                                            SyncLock LogLock
                                                Log.Rows.Add(swLog.ElapsedMilliseconds, GetAvailableCPUPercent(), GetAvailableMemoryMB(), GetMemoryUsedByMeKB)
                                                Log.AcceptChanges()
                                            End SyncLock
                                        End If
                                        Thread.Sleep(SamplingRate)
                                    Loop
                                End Sub)
        TaskCPU.Start()

        ' Initialize log table
        If Log Is Nothing Then
            Log = New DataTable("System Log")
            Log.Columns.Add("Elapsed MS", GetType(System.Int64))
            Log.Columns.Add("CPU % Free", GetType(System.Single))
            Log.Columns.Add("Memory MB Free", GetType(System.Int32))
            Log.Columns.Add("Memory KB Used", GetType(System.Int32))
        End If

    End Sub

    ''' <summary>Get the available system CPU.</summary>
    ''' <returns>Available CPU as a percent</returns>
    ''' <remarks>If called immediately after instantiating class, will delay for sampling.</remarks>
    Public Function GetAvailableCPUPercent() As Single

        ' Wait for sampling to fill the cache
        While QueueCPU.Count < SampleSize
            Thread.Sleep(SamplingRate)
        End While

        ' Move the cache into an array
        Dim QueueArray As Array = QueueCPU.ToArray
        Dim QueueCount As Integer = QueueArray.Length

        ' Calculate the average
        Dim QueueTotal As Single = 0
        For i As Integer = 0 To QueueCount - 1
            QueueTotal += QueueArray(i)
        Next

        ' Return the available percent CPU
        Return 100 - (QueueTotal / QueueCount)

    End Function

    ''' <summary>Get the available system memory.</summary>
    ''' <returns>Available system memory in MB</returns>
    Public Function GetAvailableMemoryMB() As Integer

        Return My.Computer.Info.AvailablePhysicalMemory() / 1000000

    End Function

    ''' <summary>Get the total system memory.</summary>
    ''' <returns>Available system memory in MB</returns>
    Public Function GetTotalMemoryMB() As Integer

        Return My.Computer.Info.TotalPhysicalMemory() / 1000000

    End Function

    ''' <summary>Get the available system memory.</summary>
    ''' <returns>Available system memory as a percent of total</returns>
    Public Function GetAvailableMemoryPercent() As Single

        Return GetAvailableMemoryMB() / GetTotalMemoryMB() * 100

    End Function

    ''' <summary>Get the memory used by this process.</summary>
    ''' <returns>Memory used in KB</returns>
    Public Function GetMemoryUsedByMeKB()
        Return Process.GetCurrentProcess.PrivateMemorySize64() / 1000
    End Function

    ''' <summary>Blocks the calling application until minimum system resources are available.</summary>
    ''' <param name="MinimumAvailableCPUPercent">The minimum CPU required.</param>
    ''' <param name="MinimumAvailableMemoryMB">The minimum memory required.</param>
    ''' <param name="MaximumWaitSeconds">The maximum time to wait in seconds.</param>
    ''' <returns>True if resources were reached or False if it timed out.</returns>
    Public Function WaitForResources(ByVal MinimumAvailableCPUPercent As Single, ByVal MinimumAvailableMemoryMB As Integer, ByVal MaximumWaitSeconds As Single) As Boolean

        ' Prepare parameters
        If MinimumAvailableCPUPercent = Nothing Then MinimumAvailableCPUPercent = 0
        If MinimumAvailableCPUPercent > 100 Then Return False
        If MinimumAvailableMemoryMB = Nothing Then MinimumAvailableMemoryMB = 0
        If MinimumAvailableMemoryMB > GetTotalMemoryMB() Then Return False
        If MaximumWaitSeconds = Nothing Then MaximumWaitSeconds = 5
        MaximumWaitSeconds *= 1000

        ' Wait for resources or timeout
        Dim sw As New Stopwatch()
        sw.Start()
        Do
            If GetAvailableCPUPercent() >= MinimumAvailableCPUPercent AndAlso GetAvailableMemoryMB() >= MinimumAvailableMemoryMB Then Return True
            Thread.Sleep(SamplingRate * 4)
        Loop While sw.ElapsedMilliseconds < MaximumWaitSeconds
        Return False

    End Function

    ''' <summary>Begin logging.</summary>
    ''' <param name="ContinuePrevious">True to continue previous logging session</param>
    ''' <remarks>Log results are stored in the Log datatable</remarks>
    Public Sub LogBegin(Optional ByVal ContinuePrevious As Boolean = False)

        LoggingEnabled = False
        If Not ContinuePrevious Then
            swLog.Reset()
            Log.Rows.Clear()
        End If
        swLog.Start()
        LoggingEnabled = True

    End Sub

    ''' <summary>Halt logging.</summary>
    ''' <returns>The elapsed time since first begun in MS</returns>
    ''' <remarks>Log results are stored in the Log datatable</remarks>
    Public Function LogHalt() As Long

        If Not LoggingEnabled Then Return 0

        LoggingEnabled = False
        swLog.Stop()
        Return swLog.ElapsedMilliseconds

    End Function

End Class
