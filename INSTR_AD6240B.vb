Imports System.Threading.Thread
Imports System.Windows.Forms.Application


Partial Public Class INSTRUMENT
    ' Add next two lines in initialize function 
    'Case "AD6240B"
    'er = ad6240b.OpenSession(parameter(2))

    Public Const gen_voltage As Boolean = 0
    Public Const gen_current As Boolean = 1

    Public Const out_standby As Integer = 0
    Public Const out_operate As Integer = 1
    Public Const out_suspend As Integer = 2

    Public Const mea_off As Integer = 0
    Public Const mea_voltage As Integer = 1
    Public Const mea_current As Integer = 2
    Public Const mea_resistor As Integer = 3

    Public cmd_gen_mode As String() = {"MD0", "MD1", "MD2", "MD3", "MD4"} 'DC, ﾊﾟﾙｽ, DCｽｲｰﾌﾟ, ﾊﾟﾙｽｽｲｰﾌﾟ, 低抵抗測定ﾊﾟﾙｽ
    Public cmd_gen_v_range As String() = {"SVRX", "SVR 3", "SVR 4", "SVR 5"} '最適, 300mV, 3V, 15V
    Public cmd_gen_i_range As String() = {"SIRX", "SIR -1", "SIR 0", "SIR 1", "SIR 2", "SIR 3", "SIR 4", "SIR 5"} '最適, 30μA, 300μA, 3mA, 30mA, 300mA, 1A, 4A
    Public cmd_remote_sensing_mode As String() = {"RS0", "RS1"} '2Wire, 4Wire

    Public ad6240b As New AD6240B

    Public Sub device_switch_on()
        ad6240b.output(out_operate)
        idform.change_device_swith_on_gui()
    End Sub

    Public Sub device_switch_off()
        ad6240b.output(out_standby)
        idform.change_device_switch_off_gui()
        idform.get_gen_setting()
    End Sub

    Public Sub set_voltage(ByVal voltage As Double)
        ad6240b.set_gen_function(gen_voltage)
        ad6240b.set_gen_voltage(voltage)
    End Sub

    Public Function get_voltage() As String
        Dim res As String = ad6240b.get_gen_voltage()
        Return res
    End Function

    Public Sub set_current(ByVal current As Double)
        ad6240b.set_gen_function(gen_current)
        ad6240b.set_gen_current(current)
    End Sub

    Public Function get_current() As String
        Dim res As String = ad6240b.get_gen_current()
        Return res
    End Function

    Public Function measure_voltage() As Double
        ad6240b.mea_function(mea_voltage)
        Return ad6240b.measure_voltage()
    End Function

    Public Function measure_current() As Double
        ad6240b.mea_function(mea_current)
        Return ad6240b.measure_current()
    End Function

    Public Sub set_gen_mode(ByVal cmd_mode As String)
        ad6240b.send_message(cmd_mode)
        idform.change_gen_mode_gui(cmd_mode)
    End Sub

    Public Sub set_gen_v_range(ByVal cmd_v_range As String)
        ad6240b.send_message(cmd_v_range)
        idform.change_gen_v_range_gui(cmd_v_range)
    End Sub

    Public Sub set_gen_i_range(ByVal cmd_i_range As String)
        ad6240b.send_message(cmd_i_range)
        idform.change_gen_i_range_gui(cmd_i_range)
    End Sub

    Public Sub set_limit_voltage(ByVal data1 As String, ByVal data2 As String)
        ad6240b.set_limit_voltage(data1, data2)
    End Sub

    Public Sub get_limit_voltage(ByRef high_v As String, ByRef low_v As String)
        ad6240b.get_limit_voltage(high_v, low_v)
    End Sub

    Public Sub set_limit_current(ByVal data1 As String, ByVal data2 As String)
        Dim high_i As Double = data1 / 1000
        Dim low_i As Double = data2 / 1000
        ad6240b.set_limit_current(high_i, low_i)
    End Sub

    Public Sub get_limit_current(ByRef high_i As String, ByRef low_i As String)
        ad6240b.get_limit_current(high_i, low_i)
        high_i *= 1000
        low_i *= 1000
    End Sub

    Public Sub send_message(ByVal message As String)
        ad6240b.send_message(message)
    End Sub

    Public Function read_string(ByVal message As String) As String
        Return ad6240b.read_string(message)
    End Function

    Public Sub set_trigger_mode_to_auto()
        ad6240b.set_trigger_mode_to_auto()
    End Sub

    Public Sub set_trigger_mode_to_hold()
        ad6240b.set_trigger_mode_to_hold()
    End Sub

    Public Sub set_2wire()
        ad6240b.set_2wire()
    End Sub

    Public Sub set_4wire()
        ad6240b.set_4wire()
    End Sub

    Public Sub get_wire()
        Dim res As String = ad6240b.get_wire().Trim()
        Select Case res
            Case cmd_remote_sensing_mode(0)
                idform.Radio_2wire.Checked = True
            Case cmd_remote_sensing_mode(1)
                idform.Radio_4wire.Checked = True
        End Select
    End Sub
End Class
