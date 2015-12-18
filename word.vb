Sub create_EQ_SWRESET_EN_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_SWRESET_EN: 0x0"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="SWRESET"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="wo"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+"Software reset.write to this register will launch soft reset to be asserted."+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CTRL_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CTRL: 0x2"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=3, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_ON"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+"This bit is for ON and OFF the equalizer when rEQ_DIS=0."+vbCr
    End With
    With tblNew.Cell(Row:=3, Column:=1).Range
        .Delete
        .InsertAfter Text:="1"
    End With
    With tblNew.Cell(Row:=3, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_DIS"
    End With
    With tblNew.Cell(Row:=3, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=3, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=3, Column:=5).Range
        .Delete
        .InsertAfter Text:=+"This register decides the use of equalizer. In order to use the equalizer after not using the equalizer,"+vbCr+""+vbCr+"normal operation is possible by setting this bit to '0' and then by setting Searcher Epoch. In case of"+vbCr+""+vbCr+"using equalizer timing in Combiner, this bit needs to be set to 0."+vbCr+""+vbCr+"When read, the value is what we have written to this bit, it changes immediately  after been written."+vbCr+""+vbCr+"1 : Equalizer is not used;0 : Equalizer is used"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_MODE_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_MODE: 0x4"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=4, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="INIT_RESET"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+"In case that rINIT_RESET is 1, algorithm for updating equalizer coefficients is initialized at every 512 chip unit (instead of using previous information)."+vbCr
    End With
    With tblNew.Cell(Row:=3, Column:=1).Range
        .Delete
        .InsertAfter Text:="1"
    End With
    With tblNew.Cell(Row:=3, Column:=2).Range
        .Delete
        .InsertAfter Text:="CHIP_SEL"
    End With
    With tblNew.Cell(Row:=3, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=3, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=3, Column:=5).Range
        .Delete
        .InsertAfter Text:=+"0 : 16 tap equalizer of one chip resolution;1 : 32 tap equalizer of half chip resolution"+vbCr
    End With
    With tblNew.Cell(Row:=4, Column:=1).Range
        .Delete
        .InsertAfter Text:="2"
    End With
    With tblNew.Cell(Row:=4, Column:=2).Range
        .Delete
        .InsertAfter Text:="RX_DIV_ON"
    End With
    With tblNew.Cell(Row:=4, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=4, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=4, Column:=5).Range
        .Delete
        .InsertAfter Text:=+"0 : Diversity OFF;1 : Diversity ON"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_TC_CON_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "TC_CON: 0x6"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="TC_CON"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+"1 : 1/8 chip delay; 0 : 1/8 chip advance"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_SLEW_OFFSET_L_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_SLEW_OFFSET_L: 0x8"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_SLEW_OFFSET_L"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+"This register defines the direction and relative-offset that equalizer needs to slew from its current position."+vbCr+""+vbCr+"A positive slew retards equalizer timing and a negative slew advances equalizer timing. The value must be"+vbCr+""+vbCr+" written as 2's complement number and is given in unit of 1/8 PN chip."+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_SLEW_OFFSET_M_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_SLEW_OFFSET_M: 0xa"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_SLEW_OFFSET_M"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+"rEQ_SLEW_OFFSET_M[2] is MSB and rEQ_SLEW_OFFSET_L[0] is LSB. Equalizer starts CPU-directed slew operation by"+vbCr+""+vbCr+" rEQ_SLEW_OFFSET automatically after rEQ_SLEW_ OFFSET_M is written."+vbCr+""+vbCr+"Accordingly, rEQ_SLEW_OFFSET_L must be written before rEQ_SLEW_OFFSET_M is written. This register is also used "+vbCr+""+vbCr+" to initialize equalizer. Setting rEQ_SLEW_OFFSET to '0' initializes equalizer without slew operation."+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_MASK_I_L_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_MASK_I_L: 0xc"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_MASK_I_L"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="1"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+"This register defines I channel mask value for scrambling code generation. "+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_MASK_I_M_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_MASK_I_M: 0xe"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_MASK_I_M"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+"rEQ_MASK_I_M [1] is MSB and rEQ_MASK_I_L[0] is LSB. This value applies to equalizer right after setting."+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_MASK_Q_L_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_MASK_Q_L: 0x10"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_MASK_Q_L"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="32848"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+"This register defines Q channel mask value for scrambling code generation."+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_MASK_Q_M_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_MASK_Q_M: 0x12"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_MASK_Q_M"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+"rEQ_MASK_Q_M[1] is MSB and rEQ_MASK_Q_L[0] is LSB. This value applies to equalizer right after setting."+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_NUM_CC_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_NUM_CC: 0x14"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_NUM_CC"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+"These bits define the OVSF code number for CPICH. This value applies to equalizer at the zero-offset frame boundary"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_DUMP_EN_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_DUMP_EN: 0x16"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="WD_Address"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" The write-only register values (rEQ_POSITION_L and rEQ_POSITION_M) are latched automatically after rEQ_DUMP_EN is written."+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_POSITION_L_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_POSITION_L: 0x18"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_POSITION_L"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" This register shows the position of equalizer, at the most recent status dump, relative to the most recent EPOCH time. "+vbCr+""+vbCr+" After EPOCH signal is received, equalizer position is reset. "+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_POSITION_M_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_POSITION_M: 0x1a"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_POSITION_M"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" The position of equalizer is then adjusted by CPU-directed slewing and time tracking. rEQ POSITION_M[2] is MSB and rEQ_POSITION_L[0] is LSB. "+vbCr+""+vbCr+"  The position value is 0 to 307199 and it is given in units of 1/8 PN chip "+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CH_SET_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CH_SET: 0x1c"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=3, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CH_TD"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" These bits select the transmit diversity mode of HS-PDSCH channel. "+vbCr+""+vbCr+" 00:non-TD mode; 01:STTD mode; 10:	CL1 mode"+vbCr
    End With
    With tblNew.Cell(Row:=3, Column:=1).Range
        .Delete
        .InsertAfter Text:="2"
    End With
    With tblNew.Cell(Row:=3, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CH_EN"
    End With
    With tblNew.Cell(Row:=3, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=3, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=3, Column:=5).Range
        .Delete
        .InsertAfter Text:=na
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE_SET_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE_SET: 0x1e"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=4, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_COVA_GAIN"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" This is the gain used for acquiring Covariance"+vbCr
    End With
    With tblNew.Cell(Row:=3, Column:=1).Range
        .Delete
        .InsertAfter Text:="1"
    End With
    With tblNew.Cell(Row:=3, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE_GAIN"
    End With
    With tblNew.Cell(Row:=3, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=3, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=3, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" This is the gain used for channel estimation. "+vbCr
    End With
    With tblNew.Cell(Row:=4, Column:=1).Range
        .Delete
        .InsertAfter Text:="2"
    End With
    With tblNew.Cell(Row:=4, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_COVA_SEL"
    End With
    With tblNew.Cell(Row:=4, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=4, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=4, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" This bit elects whether to replace the diagonal term of covariance matrix with sample covariance value."+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_RESI_GAIN_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "RESI_GAIN: 0x20"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="RESI_GAIN"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+"These bits are the gain used when the initial residual is acquired"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_OGRS_GAIN_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "OGRS_GAIN: 0x22"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="OGRS_GAIN"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+"These bits are the gain used when Optimal Gradient Step is acquired"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_OLES1_GAIN_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "OLES1_GAIN: 0x24"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="OLES1_GAIN"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+"These bits are the gain for acquiring the 1st variable of Optimal Learning Step of CG algorithm."+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_OLES2_GAIN_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "OLES2_GAIN: 0x26"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="OLES2_GAIN"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+"These bits are the gain for acquiring the 2nd variable of Optimal Learning Step of CG algorithm."+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_COEF_GAIN_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "COEF_GAIN: 0x28"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="COEF_GAIN"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+"These bits are the gain used when Filter Coefficient of CG algorithm is acquired."+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_GOPS_GAIN_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "GOPS_GAIN: 0x2a"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="GOPS_GAIN"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+"These bits are the gain used when Gradient Optimal Step of CG algorithm is acquired."+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_OLES3_GAIN_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "OLES3_GAIN: 0x2c"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="OLES3_GAIN"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+"These bits are the gain used when optimal learning step of CG algorithm is acquired."+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_RESI2_GAIN_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "RESI2_GAIN: 0x30"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="RESI2_GAIN"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+"These bits are the gain used when residual of CG algorithm is acquired."+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_GRAD_GAIN_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "GRAD_GAIN: 0x32"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="GRAD_GAIN"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+"These bits are the gain used when Gradient of CG algorithm is acquired."+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_ITER_THRE1_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "ITER_THRE1: 0x36"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="ITER_THRE1"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+"These bits are the value compared with Optimal Gradient Step in 1st condition that stops the Iteration of CG algorithm. "+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_ITER_THRE2_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "ITER_THRE2: 0x38"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="ITER_THRE2"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+"These bits are the value compared with the Divider output in 2nd condition that stops the Iteration of CG algorithm. "+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_FIR_GAIN_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_FIR_GAIN: 0x3a"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_FIR_GAIN"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="2"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+"These bits are the gain applied after FIR filtering by coefficient.  "+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_DEM_GAIN_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_DEM_GAIN: 0x3c"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_DEM_GAIN"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="1"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+"These bits are the gain applied after data correlation.  "+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_COVA_LOAD_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "COVA_LOAD: 0x3e"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="COVA_LOAD"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+"These bits set the value used for diagonal loading of covariance matrix.   "+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_W_SLEEP_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_W_SLEEP: 0x40"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_W_SLEEP"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="wo"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+"When rEQ_W_SLEEP is accessed, the start values needed for wake-up are loaded. The values have to be written at 0x42 ~ 0x5a before rEQ_W_SLEEP is accessed.  "+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_SMP_CNT_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_SMP_CNT: 0x42"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_SMP_CNT"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+"For sleep operation,Sample count value within one chip.Range : 0~7  (in units of 1/8 chip)  "+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_PN_CNT_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_PN_CNT: 0x44"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_PN_CNT"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+"Chip count value within one slot  "+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_SLOT_CNT_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_SLOT_CNT: 0x46"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=3, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_SLOT_CNT"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+"Sub-frame count value within one frame (10ms)  "+vbCr
    End With
    With tblNew.Cell(Row:=3, Column:=1).Range
        .Delete
        .InsertAfter Text:="4"
    End With
    With tblNew.Cell(Row:=3, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_SUBF_CNT"
    End With
    With tblNew.Cell(Row:=3, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=3, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=3, Column:=5).Range
        .Delete
        .InsertAfter Text:=+"Slot count value within one sub-frame (2ms)  "+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_POSITION_S_L_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_POSITION_S_L: 0x48"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_POSITION_S_L"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+"Current position of the equalizer,Range : 0 ~ 307199  (in units of 1/8 chip)  "+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_POSITION_S_M_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_POSITION_S_M: 0x4a"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_POSITION_S_M"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+"rEQ_POSITION_S_M[2] : MSB; rEQ_ POSITION_S_L[0] : LSB  "+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_SR1_CE_L_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_SR1_CE_L: 0x4c"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_SR1_CE_L"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+"State of shift register X for scrambling code of channel estimator  "+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_SR1_CE_M_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_SR1_CE_M: 0x4e"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_SR1_CE_M"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+"rEQ_SR1_CE_M[1] : MSB; rEQ_ SR1_CE_L[0] : LSB  "+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_SR2_CE_L_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_SR2_CE_L: 0x50"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_SR2_CE_L"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+"State of shift register Y for scrambling code of channel estimator  "+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_SR2_CE_M_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_SR2_CE_M: 0x52"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_SR2_CE_M"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+"rEQ_SR2_CE_M[1] : MSB; rEQ_ SR2_CE_L[0] : LSB  "+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_SR1_D_L_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_SR1_D_L: 0x54"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_SR1_D_L"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+"State of shift register X for scrambling code of data correlator "+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_SR1_D_M_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_SR1_D_M: 0x56"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_SR1_D_M"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+"rEQ_SR1_D_M[1] : MSB; rEQ_ SR1_D_L[0] : LSB  "+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_SR2_D_L_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_SR2_D_L: 0x58"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_SR2_D_L"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+"State of shift register Y for scrambling code of data correlator "+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_SR2_D_M_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_SR2_D_M: 0x5a"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_SR2_D_M"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="rw"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+"rEQ_SR2_D_M[1] : MSB; rEQ_ SR2_D_L[0] : LSB  "+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_eRSCP00_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_eRSCP00: 0x82"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_eRSCP00"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Measured RSCP(Received Signal Code Power) of CPICH, total with 30 16bit value, 1st 16 bit measured RSCP value"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_eRSCP01_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_eRSCP01: 0x84"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_eRSCP01"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Measured RSCP(Received Signal Code Power) of CPICH, total with 30 16bit value, 2nd 16 bit measured RSCP value"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_eRSCP02_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_eRSCP02: 0x86"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_eRSCP02"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Measured RSCP(Received Signal Code Power) of CPICH, total with 30 16bit value, 3rd 16 bit measured RSCP value"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_eRSCP03_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_eRSCP03: 0x88"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_eRSCP03"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Measured RSCP(Received Signal Code Power) of CPICH, total with 30 16bit value, 4th 16 bit measured RSCP value"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_eRSCP04_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_eRSCP04: 0x8a"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_eRSCP04"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Measured RSCP(Received Signal Code Power) of CPICH, total with 30 16bit value, 5th 16 bit measured RSCP value"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_eRSCP05_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_eRSCP05: 0x8c"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_eRSCP05"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Measured RSCP(Received Signal Code Power) of CPICH, total with 30 16bit value, 6th 16 bit measured RSCP value"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_eRSCP06_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_eRSCP06: 0x8e"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_eRSCP06"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Measured RSCP(Received Signal Code Power) of CPICH, total with 30 16bit value, 7th 16 bit measured RSCP value"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_eRSCP07_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_eRSCP07: 0x90"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_eRSCP07"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Measured RSCP(Received Signal Code Power) of CPICH, total with 30 16bit value, 8th 16 bit measured RSCP value"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_eRSCP08_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_eRSCP08: 0x92"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_eRSCP08"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Measured RSCP(Received Signal Code Power) of CPICH, total with 30 16bit value, 9th 16 bit measured RSCP value"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_eRSCP09_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_eRSCP09: 0x94"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_eRSCP09"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Measured RSCP(Received Signal Code Power) of CPICH, total with 30 16bit value, 10th 16 bit measured RSCP value"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_eRSCP10_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_eRSCP10: 0x96"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_eRSCP10"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Measured RSCP(Received Signal Code Power) of CPICH, total with 30 16bit value, 11th 16 bit measured RSCP value"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_eRSCP11_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_eRSCP11: 0x98"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_eRSCP11"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Measured RSCP(Received Signal Code Power) of CPICH, total with 30 16bit value, 12th 16 bit measured RSCP value"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_eRSCP12_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_eRSCP12: 0x9a"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_eRSCP12"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Measured RSCP(Received Signal Code Power) of CPICH, total with 30 16bit value, 13tht 16 bit measured RSCP value"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_eRSCP13_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_eRSCP13: 0x9c"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_eRSCP13"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Measured RSCP(Received Signal Code Power) of CPICH, total with 30 16bit value, 14th 16 bit measured RSCP value"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_eRSCP14_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_eRSCP14: 0x9e"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_eRSCP14"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Measured RSCP(Received Signal Code Power) of CPICH, total with 30 16bit value, 15th 16 bit measured RSCP value"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_eRSCP15_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_eRSCP15: 0xa0"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_eRSCP15"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Measured RSCP(Received Signal Code Power) of CPICH, total with 30 16bit value, 16th 16 bit measured RSCP value"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_eRSCP16_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_eRSCP16: 0xa2"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_eRSCP16"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Measured RSCP(Received Signal Code Power) of CPICH, total with 30 16bit value, 17th 16 bit measured RSCP value"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_eRSCP17_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_eRSCP17: 0xa4"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_eRSCP17"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Measured RSCP(Received Signal Code Power) of CPICH, total with 30 16bit value, 18th 16 bit measured RSCP value"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_eRSCP18_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_eRSCP18: 0xa6"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_eRSCP18"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Measured RSCP(Received Signal Code Power) of CPICH, total with 30 16bit value, 19th 16 bit measured RSCP value"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_eRSCP19_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_eRSCP19: 0xa8"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_eRSCP19"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Measured RSCP(Received Signal Code Power) of CPICH, total with 30 16bit value, 20th 16 bit measured RSCP value"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_eRSCP20_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_eRSCP20: 0xaa"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_eRSCP20"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Measured RSCP(Received Signal Code Power) of CPICH, total with 30 16bit value, 21th 16 bit measured RSCP value"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_eRSCP21_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_eRSCP21: 0xac"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_eRSCP21"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Measured RSCP(Received Signal Code Power) of CPICH, total with 30 16bit value, 22th 16 bit measured RSCP value"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_eRSCP22_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_eRSCP22: 0xae"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_eRSCP22"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Measured RSCP(Received Signal Code Power) of CPICH, total with 30 16bit value, 23th 16 bit measured RSCP value"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_eRSCP23_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_eRSCP23: 0xb0"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_eRSCP23"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Measured RSCP(Received Signal Code Power) of CPICH, total with 30 16bit value, 24th 16 bit measured RSCP value"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_eRSCP24_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_eRSCP24: 0xb2"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_eRSCP24"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Measured RSCP(Received Signal Code Power) of CPICH, total with 30 16bit value, 25th 16 bit measured RSCP value"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_eRSCP25_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_eRSCP25: 0xb4"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_eRSCP25"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Measured RSCP(Received Signal Code Power) of CPICH, total with 30 16bit value, 26th 16 bit measured RSCP value"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_eRSCP26_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_eRSCP26: 0xb6"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_eRSCP26"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Measured RSCP(Received Signal Code Power) of CPICH, total with 30 16bit value, 27th 16 bit measured RSCP value"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_eRSCP27_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_eRSCP27: 0xb8"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_eRSCP27"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Measured RSCP(Received Signal Code Power) of CPICH, total with 30 16bit value, 28th 16 bit measured RSCP value"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_eRSCP28_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_eRSCP28: 0xba"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_eRSCP28"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Measured RSCP(Received Signal Code Power) of CPICH, total with 30 16bit value, 29th 16 bit measured RSCP value"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_eRSCP29_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_eRSCP29: 0xbc"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_eRSCP29"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Measured RSCP(Received Signal Code Power) of CPICH, total with 30 16bit value, 30th 16 bit measured RSCP value"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE11_I0_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE11_I0: 0xc2"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE11_I0"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX1, 1st value I"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE11_Q0_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE11_Q0: 0xc4"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE11_Q0"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX1, 1st value Q"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE11_I1_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE11_I1: 0xc6"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE11_I1"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX1, 2nd value I"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE11_Q1_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE11_Q1: 0xc8"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE11_Q1"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX1, 2nd value Q"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE11_I2_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE11_I2: 0xca"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE11_I2"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX1, 3rd value I"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE11_Q2_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE11_Q2: 0xcc"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE11_Q2"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX1, 3rd value Q"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE11_I3_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE11_I3: 0xce"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE11_I3"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX1, 4th value I"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE11_Q3_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE11_Q3: 0xd0"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE11_Q3"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX1, 4th value Q"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE11_I4_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE11_I4: 0xd2"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE11_I4"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX1, 5th value I"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE11_Q4_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE11_Q4: 0xd4"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE11_Q4"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX1, 5th value Q"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE11_I5_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE11_I5: 0xd6"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE11_I5"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX1, 6th value I"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE11_Q5_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE11_Q5: 0xd8"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE11_Q5"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX1, 6th value Q"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE11_I6_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE11_I6: 0xda"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE11_I6"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX1, 7th value I"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE11_Q6_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE11_Q6: 0xdc"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE11_Q6"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX1, 7th value Q"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE11_I7_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE11_I7: 0xde"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE11_I7"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX1, 8th value I"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE11_Q7_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE11_Q7: 0xe0"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE11_Q7"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX1, 8th value Q"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE11_I8_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE11_I8: 0xe2"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE11_I8"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX1, 9th value I"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE11_Q8_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE11_Q8: 0xe4"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE11_Q8"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX1, 9th value Q"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE11_I9_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE11_I9: 0xe6"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE11_I3"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX1, 10th value I"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE11_Q9_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE11_Q9: 0xe8"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE11_Q3"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX1, 10th value Q"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE11_I10_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE11_I10: 0xea"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE11_I10"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX1, 11th value I"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE11_Q10_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE11_Q10: 0xec"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE11_Q10"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX1, 11th value Q"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE11_I11_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE11_I11: 0xee"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE11_I11"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX1, 12th value I"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE11_Q11_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE11_Q11: 0xf0"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE11_Q11"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX1, 12th value Q"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE21_I0_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE21_I0: 0xf2"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE21_I0"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX2, 1st value I"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE21_Q0_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE21_Q0: 0xf4"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE21_Q0"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX2, 1st value Q"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE21_I1_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE21_I1: 0xf6"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE21_I1"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX2, 2nd value I"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE21_Q1_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE21_Q1: 0xf8"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE21_Q1"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX2, 2nd value Q"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE21_I2_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE21_I2: 0xfa"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE21_I2"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX2, 3rd value I"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE21_Q2_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE21_Q2: 0xfc"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE21_Q2"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX2, 3rd value Q"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE21_I3_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE21_I3: 0xfe"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE21_I3"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX2, 4th value I"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE21_Q3_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE21_Q3: 0x100"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE21_Q3"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX2, 4th value Q"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE21_I4_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE21_I4: 0x102"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE21_I4"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX2, 5th value I"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE21_Q4_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE21_Q4: 0x104"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE21_Q4"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX2, 5th value Q"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE21_I5_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE21_I5: 0x106"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE21_I5"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX2, 6th value I"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE21_Q5_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE21_Q5: 0x108"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE21_Q5"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX2, 6th value Q"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE21_I6_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE21_I6: 0x10a"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE21_I6"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX2, 7th value I"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE21_Q6_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE21_Q6: 0x10c"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE21_Q6"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX2, 7th value Q"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE21_I7_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE21_I7: 0x10e"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE21_I7"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX2, 8th value I"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE21_Q7_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE21_Q7: 0x110"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE21_Q7"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX2, 8th value Q"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE21_I8_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE21_I8: 0x112"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE21_I8"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX2, 9th value I"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE21_Q8_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE21_Q8: 0x114"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE21_Q8"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX2, 9th value Q"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE21_I9_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE21_I9: 0x116"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE21_I3"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX2, 10th value I"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE21_Q9_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE21_Q9: 0x118"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE21_Q3"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX2, 10th value Q"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE21_I10_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE21_I10: 0x11a"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE21_I10"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX2, 11th value I"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE21_Q10_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE21_Q10: 0x11c"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE21_Q10"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX2, 11th value Q"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE21_I11_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE21_I11: 0x11e"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE21_I11"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX2, 12th value I"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub create_EQ_CE21_Q11_table() 
    Dim docActive As Document
    Dim tblNew As Table
    Dim celTable As Cell
    Dim intCount As Integer
    Set docActive = ActiveDocument
    Selection.Text = "EQ_CE21_Q11: 0x120"
    Selection.EndKey unit:=wdStory
    Set tblNew = docActive.Tables.Add( _
        Range:=docActive.Range(Start:=Selection.End, End:=Selection.End), NumRows:=2, _
        NumColumns:=5)
    tblNew.AutoFormat Format:=wdTableFormatProfessional, ApplyShading:=False, _
        ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
    With tblNew.Cell(Row:=1, Column:=1).Range
        .Delete
        .InsertAfter Text:="Bits"
    End With
    With tblNew.Cell(Row:=1, Column:=2).Range
        .Delete
        .InsertAfter Text:="Field name"
    End With
    With tblNew.Cell(Row:=1, Column:=3).Range
        .Delete
        .InsertAfter Text:="access"
    End With
    With tblNew.Cell(Row:=1, Column:=4).Range
        .Delete
        .InsertAfter Text:="default"
    End With
    With tblNew.Cell(Row:=1, Column:=5).Range
        .Delete
        .InsertAfter Text:="Description"
    End With
    With tblNew.Cell(Row:=2, Column:=1).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=2).Range
        .Delete
        .InsertAfter Text:="EQ_CE21_Q11"
    End With
    With tblNew.Cell(Row:=2, Column:=3).Range
        .Delete
        .InsertAfter Text:="ro"
    End With
    With tblNew.Cell(Row:=2, Column:=4).Range
        .Delete
        .InsertAfter Text:="0"
    End With
    With tblNew.Cell(Row:=2, Column:=5).Range
        .Delete
        .InsertAfter Text:=+" Channel Estimation of TX1 to RX2, 12th value Q"+vbCr
    End With
    With tblNew
        .Columns.AutoFit
        .Rows(1).Shading.BackgroundPatternColor = wdColorGray15
    End With
    Selection.EndKey unit:=wdStory
    End Sub
Sub all_module_regs()
    Call create_EQ_SWRESET_EN_table()
    Call create_EQ_CTRL_table()
    Call create_EQ_MODE_table()
    Call create_TC_CON_table()
    Call create_EQ_SLEW_OFFSET_L_table()
    Call create_EQ_SLEW_OFFSET_M_table()
    Call create_EQ_MASK_I_L_table()
    Call create_EQ_MASK_I_M_table()
    Call create_EQ_MASK_Q_L_table()
    Call create_EQ_MASK_Q_M_table()
    Call create_EQ_NUM_CC_table()
    Call create_EQ_DUMP_EN_table()
    Call create_EQ_POSITION_L_table()
    Call create_EQ_POSITION_M_table()
    Call create_EQ_CH_SET_table()
    Call create_EQ_CE_SET_table()
    Call create_RESI_GAIN_table()
    Call create_OGRS_GAIN_table()
    Call create_OLES1_GAIN_table()
    Call create_OLES2_GAIN_table()
    Call create_COEF_GAIN_table()
    Call create_GOPS_GAIN_table()
    Call create_OLES3_GAIN_table()
    Call create_RESI2_GAIN_table()
    Call create_GRAD_GAIN_table()
    Call create_ITER_THRE1_table()
    Call create_ITER_THRE2_table()
    Call create_EQ_FIR_GAIN_table()
    Call create_EQ_DEM_GAIN_table()
    Call create_COVA_LOAD_table()
    Call create_EQ_W_SLEEP_table()
    Call create_EQ_SMP_CNT_table()
    Call create_EQ_PN_CNT_table()
    Call create_EQ_SLOT_CNT_table()
    Call create_EQ_POSITION_S_L_table()
    Call create_EQ_POSITION_S_M_table()
    Call create_EQ_SR1_CE_L_table()
    Call create_EQ_SR1_CE_M_table()
    Call create_EQ_SR2_CE_L_table()
    Call create_EQ_SR2_CE_M_table()
    Call create_EQ_SR1_D_L_table()
    Call create_EQ_SR1_D_M_table()
    Call create_EQ_SR2_D_L_table()
    Call create_EQ_SR2_D_M_table()
    Call create_EQ_eRSCP00_table()
    Call create_EQ_eRSCP01_table()
    Call create_EQ_eRSCP02_table()
    Call create_EQ_eRSCP03_table()
    Call create_EQ_eRSCP04_table()
    Call create_EQ_eRSCP05_table()
    Call create_EQ_eRSCP06_table()
    Call create_EQ_eRSCP07_table()
    Call create_EQ_eRSCP08_table()
    Call create_EQ_eRSCP09_table()
    Call create_EQ_eRSCP10_table()
    Call create_EQ_eRSCP11_table()
    Call create_EQ_eRSCP12_table()
    Call create_EQ_eRSCP13_table()
    Call create_EQ_eRSCP14_table()
    Call create_EQ_eRSCP15_table()
    Call create_EQ_eRSCP16_table()
    Call create_EQ_eRSCP17_table()
    Call create_EQ_eRSCP18_table()
    Call create_EQ_eRSCP19_table()
    Call create_EQ_eRSCP20_table()
    Call create_EQ_eRSCP21_table()
    Call create_EQ_eRSCP22_table()
    Call create_EQ_eRSCP23_table()
    Call create_EQ_eRSCP24_table()
    Call create_EQ_eRSCP25_table()
    Call create_EQ_eRSCP26_table()
    Call create_EQ_eRSCP27_table()
    Call create_EQ_eRSCP28_table()
    Call create_EQ_eRSCP29_table()
    Call create_EQ_CE11_I0_table()
    Call create_EQ_CE11_Q0_table()
    Call create_EQ_CE11_I1_table()
    Call create_EQ_CE11_Q1_table()
    Call create_EQ_CE11_I2_table()
    Call create_EQ_CE11_Q2_table()
    Call create_EQ_CE11_I3_table()
    Call create_EQ_CE11_Q3_table()
    Call create_EQ_CE11_I4_table()
    Call create_EQ_CE11_Q4_table()
    Call create_EQ_CE11_I5_table()
    Call create_EQ_CE11_Q5_table()
    Call create_EQ_CE11_I6_table()
    Call create_EQ_CE11_Q6_table()
    Call create_EQ_CE11_I7_table()
    Call create_EQ_CE11_Q7_table()
    Call create_EQ_CE11_I8_table()
    Call create_EQ_CE11_Q8_table()
    Call create_EQ_CE11_I9_table()
    Call create_EQ_CE11_Q9_table()
    Call create_EQ_CE11_I10_table()
    Call create_EQ_CE11_Q10_table()
    Call create_EQ_CE11_I11_table()
    Call create_EQ_CE11_Q11_table()
    Call create_EQ_CE21_I0_table()
    Call create_EQ_CE21_Q0_table()
    Call create_EQ_CE21_I1_table()
    Call create_EQ_CE21_Q1_table()
    Call create_EQ_CE21_I2_table()
    Call create_EQ_CE21_Q2_table()
    Call create_EQ_CE21_I3_table()
    Call create_EQ_CE21_Q3_table()
    Call create_EQ_CE21_I4_table()
    Call create_EQ_CE21_Q4_table()
    Call create_EQ_CE21_I5_table()
    Call create_EQ_CE21_Q5_table()
    Call create_EQ_CE21_I6_table()
    Call create_EQ_CE21_Q6_table()
    Call create_EQ_CE21_I7_table()
    Call create_EQ_CE21_Q7_table()
    Call create_EQ_CE21_I8_table()
    Call create_EQ_CE21_Q8_table()
    Call create_EQ_CE21_I9_table()
    Call create_EQ_CE21_Q9_table()
    Call create_EQ_CE21_I10_table()
    Call create_EQ_CE21_Q10_table()
    Call create_EQ_CE21_I11_table()
    Call create_EQ_CE21_Q11_table()
End Sub
