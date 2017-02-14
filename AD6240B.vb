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

    Public Function get_gen_voltage() As String
        Sleep(20)
        Dim res As String = dev.Query("SOV?")
        Return res.Substring(3, res.Length - 3)
    End Function

    Public Sub set_gen_current(ByVal iset As Double)
        send_message("SOI " & iset)
    End Sub

    Public Function get_gen_current() As String
        Sleep(20)
        Dim res As String = dev.Query("SOI?")
        Return res.Substring(3, res.Length - 3)
    End Function

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
            Case INSTRUMENT.mea_resistor
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

    Public Sub set_limit_voltage(ByVal data1 As Double, ByVal data2 As Double)
        Dim cmd As String = "LMV" & Format("0.000e0", data1) & "," & Format("0.000e0", data2)
        send_message(cmd)
    End Sub

    Public Sub get_limit_voltage(ByRef high_v As Double, ByRef low_v As Double)
        Dim res As String = read_string("LMV?")
        high_v = res.Substring(3, 9)
        low_v = res.Substring(13, 9)
    End Sub

    Public Sub set_limit_current(ByVal data1 As Double, ByVal data2 As Double)
        Dim cmd As String = "LMI" & Format("0.000e0", data1) & "," & Format("0.000e0", data2)
        send_message(cmd)
    End Sub

    Public Sub get_limit_current(ByRef high_i As Double, ByRef low_i As Double)
        Dim res As String = read_string("LMI?")
        high_i = res.Substring(3, 9)
        low_i = res.Substring(13, 9)
    End Sub

    Public Sub set_trigger_mode_to_auto()
        Dim cmd As String = "M0"
        send_message(cmd)
    End Sub

    Public Sub set_trigger_mode_to_hold()
        Dim cmd As String = "M1"
        send_message(cmd)
    End Sub

    Public Sub set_2wire()
        Dim cmd As String = "RS0"
        send_message(cmd)
    End Sub

    Public Sub set_4wire()
        Dim cmd As String = "RS1"
        send_message(cmd)
    End Sub

    Public Function get_wire() As String
        Dim res As String = read_string("RS?")
        Return res
    End Function
End Class
