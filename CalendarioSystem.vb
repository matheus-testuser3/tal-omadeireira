Imports System.Windows.Forms
Imports System.Drawing

''' <summary>
''' Sistema de calendário integrado para o PDV
''' Formulário de calendário moderno para seleção de datas
''' </summary>
Public Class FormCalendario
    Inherits Form
    
    Private WithEvents monthCalendar As MonthCalendar
    Private WithEvents btnOK As Button
    Private WithEvents btnCancelar As Button
    Private WithEvents btnHoje As Button
    Private WithEvents lblDataSelecionada As Label
    
    Private _dataSelecionada As Date
    Private _targetControl As Control
    
    ''' <summary>
    ''' Data selecionada no calendário
    ''' </summary>
    Public Property DataSelecionada As Date
        Get
            Return _dataSelecionada
        End Get
        Set(value As Date)
            _dataSelecionada = value
            monthCalendar.SetDate(value)
            AtualizarLabel()
        End Set
    End Property
    
    ''' <summary>
    ''' Controle que receberá a data selecionada
    ''' </summary>
    Public Property TargetControl As Control
        Get
            Return _targetControl
        End Get
        Set(value As Control)
            _targetControl = value
        End Set
    End Property
    
    Public Sub New()
        InitializeComponent()
        _dataSelecionada = Date.Today
        ConfigurarInterface()
    End Sub
    
    Private Sub InitializeComponent()
        Me.Text = "Selecionar Data"
        Me.Size = New Size(300, 280)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.BackColor = Color.White
        Me.Icon = Nothing
        
        ' Calendário principal
        monthCalendar = New MonthCalendar() With {
            .Location = New Point(10, 10),
            .MaxSelectionCount = 1,
            .ShowToday = True,
            .ShowTodayCircle = True
        }
        
        ' Label com data selecionada
        lblDataSelecionada = New Label() With {
            .Location = New Point(10, 170),
            .Size = New Size(250, 20),
            .Text = "Data selecionada: " & Date.Today.ToString("dd/MM/yyyy"),
            .Font = New Font("Segoe UI", 9, FontStyle.Bold),
            .ForeColor = Color.DarkBlue
        }
        
        ' Botão Hoje
        btnHoje = New Button() With {
            .Text = "📅 Hoje",
            .Location = New Point(10, 200),
            .Size = New Size(70, 30),
            .BackColor = Color.LightBlue,
            .ForeColor = Color.Black,
            .FlatStyle = FlatStyle.Flat
        }
        
        ' Botão OK
        btnOK = New Button() With {
            .Text = "✅ OK",
            .Location = New Point(130, 200),
            .Size = New Size(70, 30),
            .BackColor = Color.Green,
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        
        ' Botão Cancelar
        btnCancelar = New Button() With {
            .Text = "❌ Cancelar",
            .Location = New Point(210, 200),
            .Size = New Size(70, 30),
            .BackColor = Color.Gray,
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        
        Me.Controls.AddRange({monthCalendar, lblDataSelecionada, btnHoje, btnOK, btnCancelar})
    End Sub
    
    Private Sub ConfigurarInterface()
        ' Configurar cores do calendário
        monthCalendar.TitleBackColor = Color.DarkBlue
        monthCalendar.TitleForeColor = Color.White
        monthCalendar.TrailingForeColor = Color.LightGray
        
        ' Definir data inicial
        monthCalendar.SetDate(_dataSelecionada)
        AtualizarLabel()
    End Sub
    
    Private Sub AtualizarLabel()
        lblDataSelecionada.Text = "Data selecionada: " & _dataSelecionada.ToString("dd/MM/yyyy (dddd)")
    End Sub
    
    Private Sub monthCalendar_DateSelected(sender As Object, e As DateRangeEventArgs) Handles monthCalendar.DateSelected
        _dataSelecionada = e.Start
        AtualizarLabel()
    End Sub
    
    Private Sub btnHoje_Click(sender As Object, e As EventArgs) Handles btnHoje.Click
        _dataSelecionada = Date.Today
        monthCalendar.SetDate(_dataSelecionada)
        AtualizarLabel()
    End Sub
    
    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        ' Atualizar controle de destino se especificado
        If _targetControl IsNot Nothing Then
            If TypeOf _targetControl Is TextBox Then
                CType(_targetControl, TextBox).Text = _dataSelecionada.ToString("dd/MM/yyyy")
            ElseIf TypeOf _targetControl Is DateTimePicker Then
                CType(_targetControl, DateTimePicker).Value = _dataSelecionada
            End If
            
            ' Disparar evento Change se disponível
            Try
                Dim changeMethod = _targetControl.GetType().GetMethod("OnTextChanged", 
                    Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Instance)
                If changeMethod IsNot Nothing Then
                    changeMethod.Invoke(_targetControl, {EventArgs.Empty})
                End If
            Catch
                ' Ignore se não conseguir disparar evento
            End Try
        End If
        
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub
    
    Private Sub btnCancelar_Click(sender As Object, e As EventArgs) Handles btnCancelar.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub
    
    ''' <summary>
    ''' Método estático para abrir calendário rapidamente
    ''' </summary>
    Public Shared Function SelecionarData(Optional dataInicial As Date = Nothing, 
                                         Optional targetControl As Control = Nothing) As Date?
        Dim form As New FormCalendario()
        
        If dataInicial <> Nothing Then
            form.DataSelecionada = dataInicial
        End If
        
        If targetControl IsNot Nothing Then
            form.TargetControl = targetControl
        End If
        
        If form.ShowDialog() = DialogResult.OK Then
            Return form.DataSelecionada
        Else
            Return Nothing
        End If
    End Function
End Class

''' <summary>
''' Extensões para integração automática do calendário
''' </summary>
Public Module CalendarioExtensions
    
    ''' <summary>
    ''' Adiciona botão de calendário ao lado de um TextBox
    ''' </summary>
    Public Sub AdicionarBotaoCalendario(textBox As TextBox, parent As Control)
        Dim btnCalendario As New Button() With {
            .Text = "📅",
            .Size = New Size(25, textBox.Height),
            .Location = New Point(textBox.Right + 2, textBox.Top),
            .FlatStyle = FlatStyle.Flat,
            .BackColor = Color.LightBlue,
            .ForeColor = Color.Black
        }
        
        AddHandler btnCalendario.Click, Sub(sender, e)
            Dim dataAtual As Date
            If Date.TryParse(textBox.Text, dataAtual) Then
                FormCalendario.SelecionarData(dataAtual, textBox)
            Else
                FormCalendario.SelecionarData(Date.Today, textBox)
            End If
        End Sub
        
        parent.Controls.Add(btnCalendario)
    End Sub
    
    ''' <summary>
    ''' Converte string para data com tratamento de erro
    ''' </summary>
    Public Function ConverterStringParaData(texto As String) As Date?
        Dim data As Date
        If Date.TryParse(texto, data) Then
            Return data
        End If
        Return Nothing
    End Function
    
    ''' <summary>
    ''' Formata data para exibição brasileira
    ''' </summary>
    Public Function FormatarDataBrasileira(data As Date) As String
        Return data.ToString("dd/MM/yyyy")
    End Function
    
    ''' <summary>
    ''' Formata data com dia da semana
    ''' </summary>
    Public Function FormatarDataCompleta(data As Date) As String
        Return data.ToString("dddd, dd 'de' MMMM 'de' yyyy", 
                            System.Globalization.CultureInfo.CreateSpecificCulture("pt-BR"))
    End Function
End Module

''' <summary>
''' Classe para eventos de calendário (equivalente a cCalendário.cls)
''' </summary>
Public Class CalendarioEventos
    Private _data As Date
    Private _descricao As String
    Private _tipo As String
    Private _importante As Boolean
    
    Public Property Data As Date
        Get
            Return _data
        End Get
        Set(value As Date)
            _data = value
        End Set
    End Property
    
    Public Property Descricao As String
        Get
            Return _descricao
        End Get
        Set(value As String)
            _descricao = value
        End Set
    End Property
    
    Public Property Tipo As String
        Get
            Return _tipo
        End Get
        Set(value As String)
            _tipo = value
        End Set
    End Property
    
    Public Property Importante As Boolean
        Get
            Return _importante
        End Get
        Set(value As Boolean)
            _importante = value
        End Set
    End Property
    
    Public Sub New()
        _data = Date.Today
        _descricao = ""
        _tipo = "Geral"
        _importante = False
    End Sub
    
    Public Sub New(data As Date, descricao As String, Optional tipo As String = "Geral", Optional importante As Boolean = False)
        _data = data
        _descricao = descricao
        _tipo = tipo
        _importante = importante
    End Sub
    
    Public Overrides Function ToString() As String
        Dim prefix = If(_importante, "⭐ ", "")
        Return $"{prefix}{_data:dd/MM/yyyy} - {_descricao} ({_tipo})"
    End Function
End Class

''' <summary>
''' Gerenciador de eventos de calendário
''' </summary>
Public Class CalendarioManager
    Private _eventos As List(Of CalendarioEventos)
    
    Public Sub New()
        _eventos = New List(Of CalendarioEventos)()
        CarregarEventosPadrao()
    End Sub
    
    ''' <summary>
    ''' Adiciona evento ao calendário
    ''' </summary>
    Public Sub AdicionarEvento(evento As CalendarioEventos)
        _eventos.Add(evento)
    End Sub
    
    ''' <summary>
    ''' Remove evento do calendário
    ''' </summary>
    Public Sub RemoverEvento(evento As CalendarioEventos)
        _eventos.Remove(evento)
    End Sub
    
    ''' <summary>
    ''' Obtém eventos de uma data específica
    ''' </summary>
    Public Function ObterEventosDaData(data As Date) As List(Of CalendarioEventos)
        Return _eventos.Where(Function(e) e.Data.Date = data.Date).ToList()
    End Function
    
    ''' <summary>
    ''' Obtém eventos de um período
    ''' </summary>
    Public Function ObterEventosPeríodo(dataInicio As Date, dataFim As Date) As List(Of CalendarioEventos)
        Return _eventos.Where(Function(e) e.Data.Date >= dataInicio.Date AndAlso e.Data.Date <= dataFim.Date).ToList()
    End Function
    
    ''' <summary>
    ''' Verifica se uma data possui eventos
    ''' </summary>
    Public Function DataPossuiEventos(data As Date) As Boolean
        Return _eventos.Any(Function(e) e.Data.Date = data.Date)
    End Function
    
    ''' <summary>
    ''' Carrega eventos padrão do sistema
    ''' </summary>
    Private Sub CarregarEventosPadrao()
        ' Adicionar alguns eventos de exemplo
        _eventos.Add(New CalendarioEventos(Date.Today, "Sistema PDV iniciado", "Sistema", True))
        _eventos.Add(New CalendarioEventos(Date.Today.AddDays(1), "Verificar estoque", "Estoque", False))
        _eventos.Add(New CalendarioEventos(Date.Today.AddDays(7), "Backup semanal", "Sistema", True))
    End Sub
    
    ''' <summary>
    ''' Salva eventos (implementar conforme necessário)
    ''' </summary>
    Public Sub SalvarEventos()
        ' TODO: Implementar salvamento em arquivo ou banco
    End Sub
    
    ''' <summary>
    ''' Carrega eventos (implementar conforme necessário)
    ''' </summary>
    Public Sub CarregarEventos()
        ' TODO: Implementar carregamento de arquivo ou banco
    End Sub
    
    ''' <summary>
    ''' Propriedade para acessar todos os eventos
    ''' </summary>
    Public ReadOnly Property TodosEventos As List(Of CalendarioEventos)
        Get
            Return _eventos
        End Get
    End Property
End Class