﻿Imports System.Threading

Public Class Form1

    'Dim t As Thread


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        lblProcessando.Visible = False
        ProgressBar1.Visible = False
    End Sub

    Private Sub btnImportarExcel_Click(sender As Object, e As EventArgs) Handles btnImportarExcel.Click

        'ColarDoExcel()
        Dim vAreaTransferencia As String() = Clipboard.GetText.Split(vbNewLine)

        DataGridView1.Invoke(Sub()
                                 ColarDoExcel(vAreaTransferencia)
                             End Sub)

        'Control.CheckForIllegalCrossThreadCalls = False
        't = New Thread(Sub() ColarDoExcel(vAreaTransferencia))
        't.Start()





    End Sub



    ' No evento KeyDown da DataGridView
    Private Sub DataGridView1_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DataGridView1.KeyDown

        ' Caso as teclas pressinadas sejam CTRL+V
        If e.Control AndAlso e.KeyCode = Keys.V Then
            'ColarDoExcel(x)

        End If

    End Sub

    Private Sub ColarDoExcel(ByVal pAreaTransferencia() As String)
        Try

            'Dim dt As New DataTable
            'dt.Columns.Add("coluna1")
            'dt.Columns.Add("coluna2")
            'dt.Columns.Add("coluna3")
            'dt.Columns.Add("coluna4")
            'dt.Columns.Add("coluna5")

            DataGridView1.ScrollBars = ScrollBars.None
            DataGridView1.SuspendLayout() 'Acelerou absurdamente a importacao do CTRL+V
            DataGridView1.Visible = False

            lblProcessando.Visible = True
            ProgressBar1.Visible = True

            ProgressBar1.Value = 0
            ProgressBar1.Maximum = pAreaTransferencia.Length - 1

            ' Ciclo nas linhas copiadas
            'For Each line As String In Clipboard.GetText.Split(vbNewLine)
            For Each line As String In pAreaTransferencia.AsParallel().AsOrdered()

                If line.Length = 1 Then Exit For 'Para retirar a ultima linha (estava com bug)

                ProgressBar1.Value += 1
                ' Separa as colunas referentes à linha actual
                'Dim item() As String = line.Trim.Split(vbTab)
                Dim item() As String = line.Replace(vbTab, ";").Trim.Split(";") 'Para retirar o bug caso a celula do excel esteja vazia


                ' Se o número de colunas for diferente mostra uma mensagem de erro
                If item.Length <> Me.DataGridView1.ColumnCount Then
                    Dim str As String = "O número de colunas copiadas é diferente" &
                                                " do número de colunas da Grid"
                    Throw New Exception(str)
                End If

                'teste merge
                ' dt.Rows.Add(item)
                'System.Threading.Thread.Sleep(10)
                ' Adicionar a linha a DataGridView
                Me.DataGridView1.Rows.Add(item)


            Next

            'DataGridView1.DataSource = dt
            lblTotalLinhasGrid.Text = DataGridView1.Rows.Count()
            MessageBox.Show("Informações do Excel coladas com sucesso!", My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            ' Mensagem de erro caso exista
            MessageBox.Show(ex.Message, My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            lblProcessando.Visible = False
            ProgressBar1.Visible = False

            DataGridView1.Visible = True
            DataGridView1.ResumeLayout(True) 'Acelerou absurdamente a importacao do CTRL+V

            DataGridView1.ScrollBars = ScrollBars.Both

        End Try

    End Sub

    Private Sub AvisoNaTela(ByVal pTexto As String)
        Dim lbl As New Label
        lbl.Size = New System.Drawing.Size(159, 23) 'set your size (if required)
        lbl.Location = New System.Drawing.Point(12, 180) 'set your location
        lbl.Text = pTexto 'set the text for your label
        Me.Controls.Add(lbl)  'add your new control to your forms control collection
    End Sub

End Class
