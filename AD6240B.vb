Imports System.Threading.Thread
Imports System.Windows.Forms.Application


Public Class AD6240B
    Inherits NI_VISA   'Class NS-VISAÇåpè≥

    Public Sub send_message(ByVal message As String)
        dev.Write(message)
    End Sub

    Public Function read_string(ByVal message As String) As String
        Return ReadStringValue(message)
    End Function

    Public Sub set_gen_function(ByVal gen_function As Integer)
        Select Case gen_function
            Case INSTRUMENT.gen_voltage
                send_message("VF")
            Case INSTRUMENT.gen_current
                send_message("IF")
        End Select
    End Sub

    Public Sub set_gen_voltage(ByVal vset As Double)
        send_message("SOV " & vset)
    End Sub

    Public Sub set_gen_current(ByVal iset As Double)
        send_message("SOI " & iset)
    End Sub

    Public Sub output(ByVal SW As Integer)
        Select Case SW
            Case INSTRUMENT.out_standby
                send_message("SBY")
            Case INSTRUMENT.out_operate
                send_message("OPR")
            Case INSTRUMENT.out_suspend
                send_message("SUS")
        End Select
    End Sub

    Public Sub mea_function(ByVal SW As Integer)
        Select Case SW
            Case INSTRUMENT.mea_off
                send_message("F0")
            Case INSTRUMENT.mea_voltage
                send_message("F1")
            Case INSTRUMENT.mea_current
                send_message("F2")
            Case INSTRUMENT.mea_resistar
                send_message("F3")
        End Select
    End Sub

    Public Function measure_voltage() As Double
        Dim value As String

        send_message("M1")
        value = ReadStringValue("*TRG")
        Return value.Substring(4, value.Length - 4)
    End Function

    Public Function measure_current() As Double
        Dim value As String

        send_message("M1")
        value = ReadStringValue("*TRG")
        Return value.Substring(4, value.Length - 4)
    End Function
End Class
