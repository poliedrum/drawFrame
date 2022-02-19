Sub DrawFrame()
'
'  DrawFrame: desenha formato de frame conforme sequencia especificada
'      Autor: Alexandre Freitas  <alexandreueg@yahoo.com.br>
'   LinkedIn: www.linkedin.com/in/lexion
'     Github: https://github.com/poliedrum/drawFrame
'     Versão: 0.10 Últimos ajustes para publicação
'       Data: 19.2.2022 16h28

    ' Tipos:
    '  flag_nome = flag de 1 bit (ou fl)
    '  cb0 = const bit 0
    '  cb1 = const bit 1
    '  pad = bits 0 até completar byte
    '  nb = nibble 4 bits alinhados no bit 0 ou 4
    '  cnx = const nibble com x de 0 a F (nome= 0xF)
    '  ui2..ui64 = unsigned int de 2 a 64 bits
    '  byte = ui8
    '  byXX = const byte com valor hexa XX
    '  fp = floating point 32 bits
    '  fim = garante fim com borda de byte

    ' valor para teste: nb_nome+fl_FL1+cb0+pad+ui5_nome+ui4_nome+ui7_nome+byte_nome+cn7+nb_nome+pad+cb0+cb0+fl+cb1+fl_SYN+fl_PK+fl_URG+fl_ACK+cnb+nb+byaf+ui6+ui3+ui4+ui10+ui2+pad

    Dim areaTotal As Range
    Dim areaTemp As Range
    Dim planilha As Worksheet
    Dim framespec As String
    Dim nbits As Integer
    Dim nbytes As Integer

    ' Constantes
    cor_borda_frame = &H996655  ' o & no final forca para tipo long, evitando valor negativo. Aqui nem seria preciso pois nao cabe em 16 bits. tente com &h8000
    cor_borda_byte = &H887766
    cor_borda_campo = &H999999
    cor_borda_bit = &HBBBBBB
    cor_borda_titulo = &H965430

    cor_bg_total = &HFFFDFA
    cor_bg_campo_bit1 = &HE6FAFF   ' o 1 aqui eh para alternancia de cores, nao representa o bit 1
    cor_bg_campo_bit2 = &HCCF5FF
    cor_bg_campo_nib1 = &HFAFFE8
    cor_bg_campo_nib2 = &HF2FFDA
    cor_bg_campo_byte1 = &HD1EEFF
    cor_bg_campo_byte2 = &HD1EEFF
    cor_bg_campo_byte1 = &HDDEEEB
    cor_bg_campo_byte2 = &HCCE2D8
    cor_bg_campo_ui1 = &HFFEECC
    cor_bg_campo_ui2 = &HFFE9B8
    cor_bg_campo_pad = &HEEEEEE   ' o pad nao precisa alternancia
    cor_bg_campo_cb = &HEEEEEE
    cor_bg_campo_cn = &HEEEEEE

    cor_fonte_bitpos = &HBB9999
    cor_fonte_bit = &H2933&
    cor_fonte_nibble = &H394D00
    cor_fonte_byte = &H1A3327
    cor_fonte_flag = &H333333
    cor_fonte_ui = &H4D3300
    cor_fonte_pad = &HBBBBBB
    cor_fonte_titulo = &HB5752F
    cor_fonte_rodape = &H717175

    cor_fundo = RGB(255, 255, 255)
    cor_sheet = &HFFBB66

    cont_bg_bit = 0
    cont_bg_nib = 0
    cont_bg_byte = 0
    cont_bg_ui = 0

    row_ini = 1
    col_ini = 1

    ' le especificação do frame
    framespec = Range("C3").Value

    ' Cria nova planilha para frame
    nomeNovaPlan = Left(Sheets("FrameMaker").Cells(5, 3).Value, 31)
    If Not ExistePlan(nomeNovaPlan) Then
        Sheets.Add After:=ActiveSheet
        ActiveSheet.Name = nomeNovaPlan
        ActiveSheet.Tab.Color = cor_sheet
        LimpaPlanilha
    Else
        Sheets(nomeNovaPlan).Select
        If Sheets("FrameMaker").OLEObjects("chkSobrescrever").Object.Value Then
            LimpaPlanilha
        Else
            MsgBox "Já existe uma planilha com o título informado." & vbCrLf & _
                "Altere o título ou delete a planilha"
            Exit Sub
        End If
    End If

    ' inicia desenho do frame
    nbits = 0 ' numero de bits do frame
    campos = Split(framespec, "+")
    ' calcula tamanho total em bits e bytes, verifica alinhamento e desenha cada campo
    For Each campo In campos
        infos = Split(campo, "_", 2)
        Select Case LCase(Trim(infos(0)))
            Case "flag", "fl", "cb0", "cb1"  ' campos de 1 bit
                ' desenha borda de campo e BG
                Range(Cells(row_ini + 3 + (nbits \ 8) * 3, col_ini + 9 - (nbits Mod 8)), Cells(row_ini + 5 + (nbits \ 8) * 3, col_ini + 9 - (nbits Mod 8))).Select
                If cont_bg_bit = 0 Then
                    Selection.Interior.Color = cor_bg_campo_bit1
                    cont_bg_bit = 1
                Else
                    Selection.Interior.Color = cor_bg_campo_bit2
                    cont_bg_bit = 0
                End If
                If nbits Mod 8 <> 7 Then ' se não for ultimo bit do byte, desenha borda de campo
                    With Selection.Borders(7)
                        .LineStyle = xlContinuous
                        .Color = cor_borda_campo
                        .Weight = xlThin
                    End With
                End If
                ' escreve nome inclinado ou valor
                If Left(infos(0), 1) = "f" Then
                    Cells(row_ini + 4 + (nbits \ 8) * 3, col_ini + 9 - (nbits Mod 8)).Select
                    With Selection
                        .Font.Name = "Consolas"
                        .Font.Size = 7
                        .Font.Color = cor_fonte_flag
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .Orientation = 90
                    End With
                    If UBound(infos) = 1 Then
                        ActiveCell.FormulaR1C1 = Left(Trim(infos(1)), 3)
                    Else
                        ActiveCell.FormulaR1C1 = "flag"
                    End If
                Else
                    Cells(row_ini + 4 + (nbits \ 8) * 3, col_ini + 9 - (nbits Mod 8)).Select
                    ActiveCell.FormulaR1C1 = Right(infos(0), 1)
                    With Selection
                        .Font.Name = "Consolas"
                        .Font.Size = 11
                        .Font.Color = cor_fonte_bit
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .Orientation = 0
                    End With

                End If
                nbits = nbits + 1

            Case "pad"
                padbits = 8 - (nbits Mod 8) ' se já está alinhado, adiciona mais um byte
                For i = 0 To (padbits - 1)
                    Cells(row_ini + 4 + (nbits \ 8) * 3, col_ini + 9 - (nbits Mod 8) - i).Select
                    With Selection
                        .Font.Name = "Consolas"
                        .Font.Size = 11
                        .Font.Color = cor_fonte_pad
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .Orientation = 0
                    End With
                    ActiveCell.FormulaR1C1 = "0"
                    Range(Cells(row_ini + 3 + (nbits \ 8) * 3, col_ini + 9 - (nbits Mod 8) - i), Cells(row_ini + 5 + (nbits \ 8) * 3, col_ini + 9 - (nbits Mod 8) - i)).Select
                    Selection.Borders(7).LineStyle = xlContinuous
                    Selection.Borders(7).Color = cor_borda_campo
                    Selection.Borders(7).Weight = xlThin
                    Selection.Interior.Color = cor_bg_campo_pad
                Next i
                nbits = nbits + padbits

            Case "nb", "cn0" To "cn9", "cna" To "cnf"
                If nbits Mod 4 <> 0 Then
                    MsgBox "ERRO de alinhamento de nibble no campo " & campo & ". Bits acumulados = " & nbits
                    Exit Sub
                End If
                ' merge e nome
                Range(Cells(row_ini + 4 + (nbits \ 8) * 3, col_ini + 9 - (nbits Mod 8) - 3), Cells(row_ini + 4 + (nbits \ 8) * 3, col_ini + 9 - (nbits Mod 8))).Select
                Selection.Merge
                ActiveCell.Font.Color = cor_fonte_nibble
                If Left(infos(0), 1) = "n" Then
                    If UBound(infos) = 1 Then
                        ActiveCell.FormulaR1C1 = Trim(infos(1))
                    Else
                        ActiveCell.FormulaR1C1 = "Nibble"
                    End If
                Else
                    valor_nibble = CInt("&h" & Right(infos(0), 1))
                    ActiveCell.FormulaR1C1 = "0x" & Hex(valor_nibble)
                    ' preenche linha inferior com bits
                    For i = 0 To 3
                        Cells(row_ini + 5 + (nbits \ 8) * 3, col_ini + 9 - (nbits Mod 8) - i).Select
                        ' o vba nao possui operador nativo para deslocamento de bits mas o excel expoe as funcoes Bitand, Bitlshift, Bitor, Bitrshift Bitxor
                        ActiveCell.FormulaR1C1 = Application.WorksheetFunction.Bitrshift(valor_nibble, i) And 1
                        With ActiveCell.Font
                            .Size = 6
                            .Color = cor_fonte_bit
                        End With
                    Next i
                End If
                ' borda esquerda e BG
                Range(Cells(row_ini + 3 + (nbits \ 8) * 3, col_ini + 9 - (nbits Mod 8) - 3), Cells(row_ini + 5 + (nbits \ 8) * 3, col_ini + 9 - (nbits Mod 8))).Select
                If cont_bg_nib = 0 Then
                    Selection.Interior.Color = cor_bg_campo_nib1
                    cont_bg_nib = 1
                Else
                    Selection.Interior.Color = cor_bg_campo_nib2
                    cont_bg_nib = 0
                End If
                With Selection.Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .Color = cor_borda_campo
                    .Weight = xlThin
                End With
                
                ' borda de bits
                Range(Cells(row_ini + 5 + (nbits \ 8) * 3, col_ini + 9 - (nbits Mod 8) - 3), Cells(row_ini + 5 + (nbits \ 8) * 3, col_ini + 9 - (nbits Mod 8))).Select
                With Selection.Borders(xlInsideVertical)
                    .LineStyle = xlContinuous
                    .Color = cor_borda_bit
                    .Weight = xlThin
                End With
                ' LSB e MSB
                Cells(row_ini + 3 + (nbits \ 8) * 3, col_ini + 9 - (nbits Mod 8)).Select
                ActiveCell.FormulaR1C1 = "LSB"
                Selection.Offset(0, -3).Select
                ActiveCell.FormulaR1C1 = "MSB"
                Range(Cells(row_ini + 3 + (nbits \ 8) * 3, col_ini + 9 - (nbits Mod 8) - 3), Cells(row_ini + 3 + (nbits \ 8) * 3, col_ini + 9 - (nbits Mod 8))).Select
                Selection.Font.Name = "Consolas"
                Selection.Font.Size = 6
                nbits = nbits + 4

            Case "byte", "by00" To "byff"
                If nbits Mod 8 <> 0 Then
                    MsgBox "ERRO de alinhamento de byte no campo " & campo & ". Bits acumulados = " & nbits
                    Exit Sub
                End If
                ' merge e nome
                Range(Cells(row_ini + 4 + (nbits \ 8) * 3, col_ini + 2), Cells(row_ini + 4 + (nbits \ 8) * 3, col_ini + 9)).Select
                Selection.Merge
                If infos(0) = "byte" Then
                    If UBound(infos) = 1 Then
                        ActiveCell.FormulaR1C1 = Trim(infos(1))
                    Else
                        ActiveCell.FormulaR1C1 = "byte"
                    End If
                    ActiveCell.Font.Color = cor_fonte_byte
                Else
                    valor_byte = CInt("&h" & Right(infos(0), 2))
                    If UBound(infos) = 1 Then
                        ActiveCell.FormulaR1C1 = Trim(infos(1))
                    Else
                        ActiveCell.FormulaR1C1 = "0x" & Hex(valor_byte)
                    End If
                    ' preenche linha inferior com bits
                    For i = 0 To 7
                        Cells(row_ini + 5 + (nbits \ 8) * 3, col_ini + 9 - i).Select
                        ActiveCell.FormulaR1C1 = Application.WorksheetFunction.Bitrshift(valor_byte, i) And 1
                        ' Application.WorksheetFunction.Bitrshift(wordmsg, 3) ' o vba não possui operador nativo para deslocamento de bits
                        ' mas o excel expõem as funções Bitand, Bitlshift, Bitor, Bitrshift Bitxor
                        With ActiveCell.Font
                            .Size = 6
                            .Color = cor_fonte_bit
                        End With
                    Next i
                End If
                ' BG
                Range(Cells(row_ini + 3 + (nbits \ 8) * 3, col_ini + 2), Cells(row_ini + 5 + (nbits \ 8) * 3, col_ini + 9)).Select
                If cont_bg_byte = 0 Then
                    Selection.Interior.Color = cor_bg_campo_byte1
                    cont_bg_byte = 1
                Else
                    Selection.Interior.Color = cor_bg_campo_byte2
                    cont_bg_byte = 0
                End If
                ' borda de bits
                Range(Cells(row_ini + 5 + (nbits \ 8) * 3, col_ini + 2), Cells(row_ini + 5 + (nbits \ 8) * 3, col_ini + 9)).Select
                With Selection.Borders(xlInsideVertical)
                    .LineStyle = xlContinuous
                    .Color = cor_borda_bit
                    .Weight = xlThin
                End With
                ' LSB e MSB
                Cells(row_ini + 3 + (nbits \ 8) * 3, col_ini + 9).Select
                ActiveCell.FormulaR1C1 = "LSB"
                Selection.Offset(0, -7).Select
                ActiveCell.FormulaR1C1 = "MSB"
                Range(Cells(row_ini + 3 + (nbits \ 8) * 3, col_ini + 2), Cells(row_ini + 3 + (nbits \ 8) * 3, col_ini + 9)).Select
                Selection.Font.Name = "Consolas"
                Selection.Font.Size = 6
                nbits = nbits + 8

            Case "ui2" To "ui9", "ui10" To "ui64"
                tam = CInt(Mid(infos(0), 3, 2))
                ' escreve lsb e msb
                Cells(row_ini + 3 + (nbits \ 8) * 3, col_ini + 9 - (nbits Mod 8)).Select
                ActiveCell.FormulaR1C1 = "LSB"
                Selection.Font.Name = "Consolas"
                Selection.Font.Size = 6
                Cells(row_ini + 3 + ((nbits + tam - 1) \ 8) * 3, col_ini + 9 - ((nbits + tam - 1) Mod 8)).Select
                ActiveCell.FormulaR1C1 = "MSB"
                Selection.Font.Name = "Consolas"
                Selection.Font.Size = 6
                ' borda de campo
                Range(Cells(row_ini + 3 + ((nbits + tam - 1) \ 8) * 3, col_ini + 9 - ((nbits + tam - 1) Mod 8)), Cells(row_ini + 5 + ((nbits + tam - 1) \ 8) * 3, col_ini + 9 - ((nbits + tam - 1) Mod 8))).Select
                With Selection.Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .Color = cor_borda_campo
                    .Weight = xlThin
                End With
                ' BG e borda de bits
                bits_restantes = tam
                nbits_temp = nbits
                maior_pedaco = 0
                ' enquanto tem bits no campo
                While bits_restantes > 0
                    ' calcula proximo pedaco a processar
                    bits_no_pedaco = Application.WorksheetFunction.Min(bits_restantes, 8 - nbits_temp Mod 8)
                    ' se é maior pedaco, guarda suas informacoes
                    If bits_no_pedaco > maior_pedaco Then
                        maior_pedaco = bits_no_pedaco
                        maior_pedaco_inicio = nbits_temp
                    End If
                    ' desenha borda de bits
                    Range(Cells(row_ini + 5 + (nbits_temp \ 8) * 3, col_ini + 9 - (nbits_temp Mod 8)), _
                          Cells(row_ini + 5 + (nbits_temp \ 8) * 3, col_ini + 9 - (nbits_temp Mod 8) - bits_no_pedaco + 1)).Select
                    With Selection.Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .Color = cor_borda_bit
                        .Weight = xlThin
                    End With
                    ' desenha bg
                    Range(Cells(row_ini + 3 + (nbits_temp \ 8) * 3, col_ini + 9 - (nbits_temp Mod 8)), _
                          Cells(row_ini + 5 + (nbits_temp \ 8) * 3, col_ini + 9 - (nbits_temp Mod 8) - bits_no_pedaco + 1)).Select
                    If cont_bg_ui = 0 Then
                        Selection.Interior.Color = cor_bg_campo_ui1
                    Else
                        Selection.Interior.Color = cor_bg_campo_ui2
                    End If
                    ' atualiza nbits_temp e bits_restantes
                    nbits_temp = nbits_temp + bits_no_pedaco
                    bits_restantes = bits_restantes - bits_no_pedaco
                Wend
                If cont_bg_ui = 0 Then cont_bg_ui = 1 Else cont_bg_ui = 0
                ' no maior pedaco, mescla e grava nome
                Range(Cells(row_ini + 4 + (maior_pedaco_inicio \ 8) * 3, col_ini + 9 - (maior_pedaco_inicio Mod 8) - maior_pedaco + 1), _
                      Cells(row_ini + 4 + (maior_pedaco_inicio \ 8) * 3, col_ini + 9 - (maior_pedaco_inicio Mod 8))).Select
                Selection.Merge
                ActiveCell.Font.Color = cor_fonte_ui
                If UBound(infos) = 1 Then
                    ActiveCell.FormulaR1C1 = Trim(infos(1))
                Else
                    ActiveCell.FormulaR1C1 = infos(0)
                End If
                nbits = nbits + tam

            Case "fp" ' TODO verificar se permite desalinhado e implementar campo tipo ponto flutuante
                nbits = nbits + 32
            ' TODO especificar outros tipos de campos cuja representacao eh necessaria
            Case "fim"
                If nbits Mod 8 <> 0 Then
                    MsgBox "ERRO de alinhamento ao encontrar marca 'fim'. Total de bits = " & nbits
                    Sheets(nomeNovaPlan).Select
                    Application.DisplayAlerts = False
                    ActiveWindow.SelectedSheets.Delete
                    Application.DisplayAlerts = True
                    Sheets("FrameMaker").Select
                    Range("A1").Select
                    Exit Sub
                End If
            Case Else
                Debug.Print "Tipo de campo ainda não implementado: " & infos(0)
        End Select
    Next campo
    Debug.Print "tamanho total em bits=" & nbits & "     tamanho em bytes=" & (nbits / 8)
    If nbits Mod 8 <> 0 Then
        MsgBox "O ideal é que o frame esteja alinhado dentro de bytes completos. Total de bits = " & nbits
        nbits = nbits + (8 - (nbits Mod 8))
    End If
    nbytes = nbits / 8

    ' Borda do Frame
    Set areaTemp = Range(Cells(row_ini + 3, col_ini + 2), Cells(row_ini + 3 + (nbytes * 3) - 1, col_ini + 2 + 8 - 1))
    areaTemp.Font.Name = "Consolas"
    For i = 7 To 10 ' Left,Top,Bottom,Right = 7, 8, 9, 10    InsideVertical,InsideHorizontal = 11, 12   DiagUpDown, DiagBottomUp = 5, 6
        With areaTemp.Borders(i)
            .LineStyle = xlContinuous
            .Color = cor_borda_frame
            .Weight = xlMedium
        End With
    Next i

    ' Bits position
    Cells(row_ini + 2, col_ini + 2).Select
    For i = 7 To 0 Step -1
        ActiveCell.Value = i
        Selection.Offset(0, 1).Select
    Next i
    Range(Cells(row_ini + 2, col_ini + 2), Cells(row_ini + 2, col_ini + 9)).Select
    With Selection
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Font.Color = cor_fonte_bitpos
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Borda de bytes
    For i = 1 To nbytes - 1
        Set areaTemp = Range(Cells(row_ini + 2 + i * 3, col_ini + 2), Cells(row_ini + 2 + i * 3, col_ini + 9))
        With areaTemp.Borders(9)
            .LineStyle = xlContinuous
            .Color = cor_borda_byte
            .Weight = xlThin
        End With
    Next i

    ' Altura linhas
    Dim alturas() As Integer, tamfontes() As Single, cores() As Long
    ReDim alturas(0 To 2): ReDim tamfontes(0 To 2): ReDim cores(0 To 2)
    alturas(0) = 9: tamfontes(0) = 7: cores(0) = &H225F5F5F    ' o 22 é o alpha channel mas é ignorado. Mesmo que FF
    alturas(1) = 18: tamfontes(1) = 11: cores(1) = &H5F5F5F
    alturas(2) = 8: tamfontes(2) = 6: cores(2) = &HFF0000
    For i = 0 To 2
        For bytes = 0 To (nbytes - 1)
            Rows(row_ini + 3 + bytes * 3 + i).Select
            With Selection
                .RowHeight = alturas(i)
                '.font.size = tamfontes(i)
                '.font.Color = cores(i)
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
        Next bytes
    Next i

    ' Titulo do Frame
    Range(Cells(row_ini, col_ini + 1), Cells(row_ini, col_ini + 10)).Select
    Selection.Merge
    Selection.Value = nomeNovaPlan
    Selection.Font.Name = "Calibri"
    Selection.Font.Size = 16
    Selection.Font.Color = cor_fonte_titulo
    Selection.RowHeight = 35
    Selection.HorizontalAlignment = xlCenter
    Selection.VerticalAlignment = xlBottom
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = cor_borda_titulo
        .Weight = xlMedium
    End With

    ' Rodapé do Frame
    Range(Cells(row_ini + 3 + nbytes * 3, col_ini + 2), Cells(row_ini + 3 + nbytes * 3 + 2, col_ini + 9)).Select
    With Selection
        .Merge
        .Value = Sheets("FrameMaker").Range("N5").Value
        .Font.Name = "Book Antiqua"
        .Font.Size = 11
        .Font.Color = cor_fonte_rodape
        .Font.Italic = True
        .RowHeight = 19
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
    End With

    ' Numera Byte offset
    For b = 0 To nbytes - 1
        Cells(row_ini + 4 + b * 3, col_ini + 1).Select
        Selection.FormulaR1C1 = b
        Selection.Font.Name = "Consolas"
        Selection.Font.Size = 9
        Selection.Font.Color = &HAA8888
    Next b

    ' acrescenta expressao que gerou frame, de forma invisivel
    Cells(row_ini + 7 + nbytes * 3, col_ini + 2).Select
    Selection.FormulaR1C1 = framespec
    Cells(row_ini + 8 + nbytes * 3, col_ini + 2).Select
    Selection.FormulaR1C1 = nomeNovaPlan
    Cells(row_ini + 9 + nbytes * 3, col_ini + 2).Select
    Selection.FormulaR1C1 = Sheets("FrameMaker").Range("N5").Value

    Range(Cells(row_ini + 7 + nbytes * 3, col_ini + 2), Cells(row_ini + 9 + nbytes * 3, col_ini + 2)).Select
    Selection.Font.Name = "Consolas"
    Selection.Font.Size = 8
    Selection.Font.Color = &HFFFFFF

End Sub


Sub LimpaPlanilha()
    If ActiveSheet.Name = "FrameMaker" Then
        Exit Sub
    End If
    ' Largura e altura padrão celulas
    Columns("A:LX").ColumnWidth = 3.14
    Rows("1:10000").RowHeight = 18

    Range(Range("A1"), Selection.End(xlDown).End(xlToRight)).Select
    Selection.ClearContents
    With Selection
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        .Font.Name = "Calibri"
        .Font.FontStyle = "Regular"
        .Font.Size = 11
        .Font.Strikethrough = False
        .Font.Superscript = False
        .Font.Subscript = False
        .Font.OutlineFont = False
        .Font.Shadow = False
        .Font.Underline = xlUnderlineStyleNone
        .Font.ThemeColor = xlThemeColorLight1
        .Font.TintAndShade = 0
        .Font.ThemeFont = xlThemeFontMinor
        .Interior.Pattern = xlNone
        .Interior.TintAndShade = 0
        .Interior.PatternTintAndShade = 0
        .Interior.Color = &HFFFFFF
        .NumberFormat = "General"
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub

Sub Ex(n As Integer)
    titulo = Sheets("FrameMaker").Cells(26 + n, 2).Value
    expressao = Sheets("FrameMaker").Cells(26 + n, 3).Value
    rodape = Sheets("FrameMaker").Cells(26 + n, 35).Value
    
    Sheets("FrameMaker").Cells(3, 3).Value = expressao
    Sheets("FrameMaker").Cells(5, 3).Value = titulo
    Sheets("FrameMaker").Cells(5, 14).Value = rodape
        
End Sub

Public Function ExistePlan(ByVal nome As String) As Boolean
    For i = 1 To Worksheets.Count
        If Worksheets(i).Name = nome Then
            ExistePlan = True
            Exit Function
        End If
    Next i
    ExistePlan = False
End Function
