Imports System.Threading.Thread
Imports System.Windows.Forms.Application


Partial Public Class INSTRUMENT
    ' Add next lines in initialize function 
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
    Public Const mea_resistar As Integer = 3

    Public gen_mode As String() = {"MD0", "MD1", "MD2", "MD3", "MD4"} 'DC, ﾊﾟﾙｽ, DCｽｲｰﾌﾟ, ﾊﾟﾙｽｽｲｰﾌﾟ, 低抵抗測定ﾊﾟﾙｽ
    Public gen_v_ragne As String() = {"SVRX", "SVR3", "SVR4", "SVR5"} '最適, 300mV, 3V, 15V
    Public gen_i_range As String() = {"SIRX", "SVR-1", "SVR0", "SVR1", "SVR2", "SVR3", "SVR4", "SVR5"} '最適, 30μA, 300μA, 3mA, 30mA, 300mA, 1A, 4A

    Public ad6240b As New AD6240B

    Public Sub switch_on()
        ad6240b.output(out_operate)
        idform.Group_genarate.Enabled = False
        idform.Button_POWER.BackColor = Color.Red
    End Sub

    Public Sub switch_off()
        ad6240b.output(out_standby)
        idform.Group_genarate.Enabled = True
        idform.Button_POWER.BackColor = Color.Silver
    End Sub

    Public Sub set_voltage(ByVal voltage As Double)
        ad6240b.set_gen_function(gen_voltage)
        ad6240b.set_gen_voltage(voltage)
    End Sub

    Public Sub set_current(ByVal current As Double)
        ad6240b.set_gen_function(gen_current)
        ad6240b.set_gen_current(current)
    End Sub

    Public Function measure_voltage() As Double
        ad6240b.mea_function(mea_voltage)
        Return ad6240b.measure_voltage()
    End Function

    Public Function measure_current() As Double
        ad6240b.mea_function(mea_current)
        Return ad6240b.measure_current()
    End Function

    Public Function get_gen_mode() As Integer
        Return Array.IndexOf(gen_mode, ad6240b.read_string("MD?"))
    End Function

    Public Sub set_gen_mode(ByVal mode_index As String)
        ad6240b.send_message(gen_mode(mode_index))
        For Each radio As RadioButton In idform.Group_gen_mode.Controls
            If radio.TabIndex = mode_index Then
                radio.Checked = True
                Exit For
            End If
        Next
    End Sub

    Public Function get_gen_v_range() As Integer
        Return Array.IndexOf(gen_mode, ad6240b.read_string("MD?").Substring(4))
    End Function

    Public Sub set_gen_v_range(ByVal range_index As String)
        ad6240b.send_message(gen_v_ragne(range_index))
        For Each radio As RadioButton In idform.Group_gen_v_range.Controls
            If radio.TabIndex = range_index Then
                radio.Checked = True
                Exit For
            End If
        Next
    End Sub

    Public Function get_gen_i_range() As String
        Dim range As String = ad6240b.read_string("SIR?")
        Select Case range.Substring(3, 1)
            Case "X"
                Return 0
            Case "-"
                Return 1
            Case "0"
                Return 2
            Case "1"
                Return 3
            Case "2"
                Return 4
            Case "3"
                Return 5
            Case "4"
                Return 6
            Case "5"
                Return 7
        End Select
        Return 0
    End Function

    Public Sub set_gen_i_range(ByVal range_index As String)
        ad6240b.send_message(gen_i_range(range_index))
        For Each radio As RadioButton In idform.Group_gen_i_range.Controls
            If radio.TabIndex = range_index Then
                radio.Checked = True
                Exit For
            End If
        Next
    End Sub
End Class
