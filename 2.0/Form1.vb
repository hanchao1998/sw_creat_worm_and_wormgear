Imports System.Math
Public Class Form1
    '初始化
    Public Sub New() '此调用是设计器所必需的。
        InitializeComponent()
        ComboBoxWormType.Text = "阿基米德圆柱蜗杆(ZA型)"
        ComboBoxWormTypeAssem.Text = "阿基米德圆柱蜗杆(ZA型)"
        ComboBoxWormGearTypeAssem.Text = "整体式"
        ComboBoxWormGearType.Text = "整体式"
        ComboboxEdition.Text = 2016
        ComboBoxalpha.Text = 20
        ComboBoxDirection.Text = "右旋"
        ComboBoxDirectionAssem.Text = "右旋"
        ComboBoxhax.Text = 1
        ComboBoxcx.Text = 0.2
        ComboBoxm.Text = 2.5
        ComboBoxz2.Text = 27
        ComboBoxKeyWayNum.Text = 1
        Initialization = False
        FilePath = Application.StartupPath
        TextBoxFilePath.Text = FilePath
        DataUpdating()
        '在 InitializeComponent() 调用之后添加任何初始化。
    End Sub
    Dim objConn As New OleDb.OleDbConnection '声明连接
    Dim objAdp As New OleDb.OleDbDataAdapter '声明数据适配器
    Dim objDataSetmd, objDataSetmdz, objDataSetNutBolt, objDataSetBoltNutWasher, objDataSetKeyway As New DataSet '声明数据集
    Dim objDataSetm2d1MIN1, objDataSetm2d1MIN2 As New DataSet '声明数据集
    Dim conStr$, sqlStr$ '声明数组
    Dim Swapp As SldWorks.SldWorks
    Dim Part As SldWorks.ModelDoc2
    Dim Assem As SldWorks.AssemblyDoc
    Dim Sketchmer As SldWorks.SketchManager
    Dim StartPoint As SldWorks.SketchPoint
    Dim EndPoint As SldWorks.SketchPoint
    Dim Sketcharc As SldWorks.SketchArc
    Dim Featmgr As SldWorks.FeatureManager
    Dim Skeline As SldWorks.SketchLine
    Dim m#, d1#, d2#, z1%, z2%, Alpha#, Direction$, hax#, cx#, x2#, i#, q#, a#, RTooth#, Gamma# '定义基本参数
    Dim da1#, da2#, df1#, df2#, ha1#, hf1#, h1#, ha2#, hf2#, h2#, b1#, b2# '定义齿顶圆、齿底圆相关参数.
    Dim l1#, d11# '定义蜗杆自定义参数
    Dim FilePath$ '文件保存路径
    Dim AssemWormTitle$, AssemWormGearTitle$, AssemWormPath$, AssemWormGearPath$ '装配的文件名称
    Dim AssemWormGearType% '装配时蜗轮的类型
    Dim D2Spoke#, D2Hub#, D2Axis#, D2SpokeHole#, L2Hub#, B2Spoke#, D2Rim#, RCast# '定义蜗轮结构自定义参数；Hub轮毂，Axis轴，Rim轮缘
    Dim dj1#, dj2# '定义两齿轮节圆
    Dim dm2#, rg2#, rf2#, k#, db2# '蜗轮结构其他参数
    '键槽参数
    Dim Keywayb#, Keywayh#
    '螺栓参数
    Dim BoltL#, BoltD#, BoltNuts#, Boltk#, Boltdp#, Boltl2#, Nutm#, Washerd1#, Washerd2#, Washerh#, BoltNute#
    '螺钉参数
    Dim NutBoltD#, NutBoltHoleL#, NutBoltNum%, NutBoltn#, NutBoltt#, NutBoltc#, NutBoltL#, NutBoltdt#
    Dim Assema# '装配时中心距
    Dim InlayTypeWormGearKeyNum% '镶铸式键的数量
    Dim SpokeHoleNum% '轮辐孔的数量
    Dim Inputsize As Boolean  '定义布尔型以还原输入尺寸值
    Dim Initialization As Boolean = True '定义布尔型确定是否在初始化
    Dim UseDatebase As Boolean = True '定义布尔型确定是否使用数据库内容
    Dim SelfCirculation As Boolean = False '定义布尔型确定是否自我循环
    Dim ParameterError As Boolean  '定义系数确定是否出现错误
    Dim SolidworksEdition% '定义整形以选择软件版本
    Dim iDesign#, P1Design#, n1Design#, TDesign#, z1Design%, z2Design%, n2Design#, EfficiencyDesign#, T2Design#, vsDesign#, vsDesign500#, vsDesign750#, vsDesign1000#, vsDesign1500#
    Dim WormGearMaterialDesign$, WormMaterialDesign$, LubricationModeDesign$, SigmaHPDesign#, SigmaHPSkimDesign#, ZVSDesign#, ZNDesign#, NDesign#， LoadDriectionDesign$
    Dim KDesign#, m2d1MIN#, mDesign#, d1Design#
    Dim SigmaHDesign#, ZEDesign#, KADesign#, KVDesign#, KBetaDesign#, d2Design#, v2Design#
    Dim SigmaFDesign#, YFSDesign#, YBetaDesign#, SigmaFPSkimDesign#, SigmaFPDesign#, YNDesign#, GammaDesign#, zvDesign#, YSADesign#, YFADesign#
    '选择文件保存的路径
    Private Sub ButtonSelectFilePath_Click(sender As Object, e As EventArgs) Handles ButtonSelectFilePath.Click
        If FolderBrowserDialogFilePath.ShowDialog = DialogResult.OK Then TextBoxFilePath.Text = FolderBrowserDialogFilePath.SelectedPath
    End Sub

    'md1数据库
    Private Sub ComboBoxm_TextChanged(sender As Object, e As EventArgs) Handles ComboBoxm.TextChanged 'Comboboxm的值改变确定d1
        If SelfCirculation = False Then
            If ComboBoxm.DropDownStyle = 1 Then ComboBoxm.DropDownStyle = 2
            If IsNumeric(ComboBoxm.Text) Then
                Select Case ComboBoxm.Text
                    Case 1, 1.25, 1.6, 2, 2.5, 3.15, 4, 5, 6.3, 8, 10, 12.5, 16, 20, 25
                        UseDatebase = True
                        ComboBoxd1.DropDownStyle = 1
                        ComboBoxz1.DropDownStyle = 1
                        ComboBoxz2.DropDownStyle = 2
                        Dim i%, j#
                        objDataSetmd.Clear() '清空数据集md
                        ComboBoxd1.Items.Clear() '清空Comboboxd1内容
                        conStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data source=Worm and Worm Gear.mdb" '字符串链接到数据库
                        sqlStr = "select * from md1 where m=" & ComboBoxm.Text '选择符合m的d1
                        objConn.ConnectionString = conStr '链接到数据库
                        objAdp = New OleDb.OleDbDataAdapter(sqlStr, objConn) '符合m的d1给数据适配器
                        objAdp.Fill(objDataSetmd, "md") '数据适配器的m和d1赋值给数据集md
                        For i = 0 To objDataSetmd.Tables("md").Rows.Count - 1 '数据集md内容给Comboboxd1
                            j = objDataSetmd.Tables("md").Rows(i).Item(1)
                            ComboBoxd1.Items.Add(j)
                        Next
                        ComboBoxd1.Text = objDataSetmd.Tables("md").Rows(0).Item(1) 'Comboboxd1默认为最小的d1
                        ComboBoxd1.DropDownStyle = 2
                        If Initialization = False Then DataUpdating()
                    Case 1.5, 3, 3.5, 4.5, 5.5, 6, 7, 12, 14, 31.5, 40
                        UseDatebase = False
                        ComboBoxd1.DropDownStyle = 0
                        ComboBoxz1.DropDownStyle = 0
                        ComboBoxz2.DropDownStyle = 0
                End Select
            Else
                SelfCirculation = True
                MessageBox.Show("模数m必须选择数字！", "格式不正确"， MessageBoxButtons.OK, MessageBoxIcon.Warning)
                ComboBoxm.DropDownStyle = 1
                ComboBoxm.Text = 2.5
            End If
        End If
        SelfCirculation = False
        If Initialization = False Then DataUpdating()
    End Sub
    'md1z1数据库
    Private Sub ComboBoxd1_TextChanged(sender As Object, e As EventArgs) Handles ComboBoxd1.TextChanged 'Comboboxd1的值改变确定z1
        If UseDatebase = True Then
            Dim i%, j%
            objDataSetmdz.Clear() '清空数据集mdz
            ComboBoxz1.Items.Clear() '清空Comboboxz1内容
            conStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data source=Worm and Worm Gear.mdb" '字符串链接到数据库
            sqlStr = "select * from md1z1 where m=" & ComboBoxm.Text & "and d1=" & ComboBoxd1.Text '字符串选择符合m和d1的z1
            objConn.ConnectionString = conStr '链接到数据库
            objAdp = New OleDb.OleDbDataAdapter(sqlStr, objConn) '符合m和d1的z1给数据适配器
            objAdp.Fill(objDataSetmdz, "mdz") '数据适配器的m,d1和z1赋值给数据集mdz
            For i = 0 To objDataSetmdz.Tables("mdz").Rows.Count - 1 '数据集mdz内容给Comboboxz1
                j = objDataSetmdz.Tables("mdz").Rows(i).Item(2)
                ComboBoxz1.Items.Add(j)
            Next
            ComboBoxz1.Text = objDataSetmdz.Tables("mdz").Rows(0).Item(2) 'Comboboxz1默认为最小的z1
            ComboBoxz1.DropDownStyle = 2
        End If
        If Initialization = False Then DataUpdating()
    End Sub
    '键槽数据库
    Public Sub Keyway()
        objDataSetKeyway.Clear() '清空数据集
        conStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data source=Worm and Worm Gear.mdb" '字符串链接到数据库
        sqlStr = "select * from 键槽 where dMIN<" & D2Axis & "and dMAX>=" & D2Axis  '字符串选择符合的参数
        objConn.ConnectionString = conStr '链接到数据库
        objAdp = New OleDb.OleDbDataAdapter(sqlStr, objConn) '符合的值给数据适配器
        objAdp.Fill(objDataSetBoltNutWasher, "KeywayDataSet") '数据适配器的值赋值给数据集
        Keywayb = objDataSetBoltNutWasher.Tables("KeywayDataSet").Rows(0).Item(2) / 1000
        Keywayh = objDataSetBoltNutWasher.Tables("KeywayDataSet").Rows(0).Item(3) / 1000
    End Sub
    '螺栓螺母垫圈数据库
    Public Sub BoltNutWasher()
        objDataSetBoltNutWasher.Clear() '清空数据集
        conStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data source=Worm and Worm Gear.mdb" '字符串链接到数据库
        sqlStr = "select * from 螺栓螺母垫圈 where 螺纹规格d<=" & BoltD  '字符串选择符合的参数
        objConn.ConnectionString = conStr '链接到数据库
        objAdp = New OleDb.OleDbDataAdapter(sqlStr, objConn) '符合的值给数据适配器
        objAdp.Fill(objDataSetBoltNutWasher, "BoltNutWasherDataSet") '数据适配器的值赋值给数据集
        'DataGridView1.DataSource = objDataSetmd.Tables("BoltNutWasherDataSet") '显示数据集内容
        BoltD = objDataSetBoltNutWasher.Tables("BoltNutWasherDataSet").Rows(objDataSetBoltNutWasher.Tables("BoltNutWasherDataSet").Rows.Count - 1).Item(0) / 1000
        BoltNuts = objDataSetBoltNutWasher.Tables("BoltNutWasherDataSet").Rows(objDataSetBoltNutWasher.Tables("BoltNutWasherDataSet").Rows.Count - 1).Item(1) / 1000
        Boltk = objDataSetBoltNutWasher.Tables("BoltNutWasherDataSet").Rows(objDataSetBoltNutWasher.Tables("BoltNutWasherDataSet").Rows.Count - 1).Item(2) / 1000
        Boltdp = objDataSetBoltNutWasher.Tables("BoltNutWasherDataSet").Rows(objDataSetBoltNutWasher.Tables("BoltNutWasherDataSet").Rows.Count - 1).Item(3) / 1000
        Boltl2 = objDataSetBoltNutWasher.Tables("BoltNutWasherDataSet").Rows(objDataSetBoltNutWasher.Tables("BoltNutWasherDataSet").Rows.Count - 1).Item(4) / 1000
        Nutm = objDataSetBoltNutWasher.Tables("BoltNutWasherDataSet").Rows(objDataSetBoltNutWasher.Tables("BoltNutWasherDataSet").Rows.Count - 1).Item(5) / 1000
        Washerd1 = objDataSetBoltNutWasher.Tables("BoltNutWasherDataSet").Rows(objDataSetBoltNutWasher.Tables("BoltNutWasherDataSet").Rows.Count - 1).Item(6) / 1000
        Washerd2 = objDataSetBoltNutWasher.Tables("BoltNutWasherDataSet").Rows(objDataSetBoltNutWasher.Tables("BoltNutWasherDataSet").Rows.Count - 1).Item(1) / 1000 '实际d2产生干涉
        Washerh = objDataSetBoltNutWasher.Tables("BoltNutWasherDataSet").Rows(objDataSetBoltNutWasher.Tables("BoltNutWasherDataSet").Rows.Count - 1).Item(8) / 1000
        BoltNute = objDataSetBoltNutWasher.Tables("BoltNutWasherDataSet").Rows(objDataSetBoltNutWasher.Tables("BoltNutWasherDataSet").Rows.Count - 1).Item(9) / 1000
        BoltL = B2Spoke + Nutm * 1.5 + Washerh
    End Sub
    '螺钉数据库
    Public Sub NutBolt()
        objDataSetNutBolt.Clear() '清空数据集
        conStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data source=Worm and Worm Gear.mdb" '字符串链接到数据库
        sqlStr = "select * from 紧定螺钉 where 螺纹规格d<=" & NutBoltD  '字符串选择符合的参数
        objConn.ConnectionString = conStr '链接到数据库
        objAdp = New OleDb.OleDbDataAdapter(sqlStr, objConn) '符合的值给数据适配器
        objAdp.Fill(objDataSetNutBolt, "NutBoltDataSet") '数据适配器的值赋值给数据集
        NutBoltD = objDataSetNutBolt.Tables("NutBoltDataSet").Rows(objDataSetNutBolt.Tables("NutBoltDataSet").Rows.Count - 1).Item(0) / 1000
        NutBoltdt = objDataSetNutBolt.Tables("NutBoltDataSet").Rows(objDataSetNutBolt.Tables("NutBoltDataSet").Rows.Count - 1).Item(1) / 1000
        NutBoltc = objDataSetNutBolt.Tables("NutBoltDataSet").Rows(objDataSetNutBolt.Tables("NutBoltDataSet").Rows.Count - 1).Item(2) / 2000
        NutBoltn = (objDataSetNutBolt.Tables("NutBoltDataSet").Rows(objDataSetNutBolt.Tables("NutBoltDataSet").Rows.Count - 1).Item(3) + objDataSetNutBolt.Tables("NutBoltDataSet").Rows(objDataSetNutBolt.Tables("NutBoltDataSet").Rows.Count - 1).Item(4)) / 2000
        NutBoltt = (objDataSetNutBolt.Tables("NutBoltDataSet").Rows(objDataSetNutBolt.Tables("NutBoltDataSet").Rows.Count - 1).Item(5) + objDataSetNutBolt.Tables("NutBoltDataSet").Rows(objDataSetNutBolt.Tables("NutBoltDataSet").Rows.Count - 1).Item(6)) / 2000
    End Sub


    '实时更新
    Private Sub ComboBoxz1orz2orx2_TextChanged(sender As Object, e As EventArgs) Handles ComboBoxz1.TextChanged, ComboBoxz2.TextChanged, TextBoxx2.TextChanged
        If Initialization = False Then DataUpdating()
    End Sub
    '数据更新
    Public Sub DataUpdating()
        If IsNumeric(ComboBoxm.Text) And IsNumeric(ComboBoxd1.Text) And IsNumeric(ComboBoxz1.Text) And IsNumeric(ComboBoxz2.Text) Then
            TextBoxi.Text = ComboBoxz2.Text / ComboBoxz1.Text
            TextBoxq.Text = ComboBoxd1.Text / ComboBoxm.Text
            TextBoxd2.Text = ComboBoxm.Text * ComboBoxz2.Text
            If IsNumeric(TextBoxx2.Text) Then
                TextBoxa.Text = 0.5 * ComboBoxd1.Text + 0.5 * TextBoxd2.Text + ComboBoxm.Text * TextBoxx2.Text
            Else
                TextBoxa.Text = 0.5 * ComboBoxd1.Text + 0.5 * TextBoxd2.Text
            End If
            TextBoxgamma.Text = FormatNumber(180 * Atan(ComboBoxz1.Text * ComboBoxm.Text / ComboBoxd1.Text) / PI, 2)
        End If
    End Sub
    '数值检查
    Public Sub NumericalCheck()
        ParameterError = False
        If UseDatebase = False Then '判断未调用数据库参数输入值时是否正确
            If IsNumeric(ComboBoxd1.Text) Then '判断d1输入的值是否正确
                If ComboBoxd1.Text <= 0 Then
                    ParameterError = True
                    MessageBox.Show("请输入正确的蜗杆分度圆直径！", "格式不正确"， MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If
            Else
                ParameterError = True
                MessageBox.Show("请输入正确的蜗杆分度圆直径！", "格式不正确"， MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If
            If IsNumeric(ComboBoxz1.Text) Then '判断z1输入的值是否正确
                If ComboBoxz1.Text > 0 And Int(ComboBoxz1.Text) = ComboBoxz1.Text Then
                Else
                    ParameterError = True
                    MessageBox.Show("请输入正确的蜗杆头数！", "格式不正确"， MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If
            Else
                ParameterError = True
                MessageBox.Show("请输入正确的蜗杆头数！", "格式不正确"， MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If
        End If
        If IsNumeric(ComboBoxz2.Text) Then '判断z2输入的值是否正确
            If ComboBoxz2.Text > 0 And Int(ComboBoxz2.Text) = ComboBoxz2.Text Then
            Else
                ParameterError = True
                MessageBox.Show("请输入正确的蜗轮齿数！", "格式不正确"， MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If
        Else
            ParameterError = True
            MessageBox.Show("请输入正确的蜗轮齿数！", "格式不正确"， MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        If IsNumeric(TextBoxx2.Text) Then '判断x2输入的值是否正确
            If TextBoxx2.Text <= -1 Or TextBoxx2.Text >= 1 Then
                ParameterError = True
                MessageBox.Show("请输入正确的变位系数（一般取-1<x2<1）！", "格式不正确"， MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If
        Else
            ParameterError = True
            MessageBox.Show("请输入正确的变位系数（一般取-1<x2<1）！", "格式不正确"， MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        If IsNumeric(ComboBoxRTooth.Text) Then '判断RTooth输入的值是否正确
            If ComboBoxRTooth.Text < 0 Then
                ParameterError = True
                MessageBox.Show("请输入正确的齿根圆角半径！", "格式不正确"， MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If
        Else
            ParameterError = True
            MessageBox.Show("请输入正确的齿根圆角半径！", "格式不正确"， MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        If IsNumeric(TextBoxSpokeHoleNum.Text) Then '判断SpokeHoleNum输入的值是否正确
            If TextBoxSpokeHoleNum.Text > 0 And Int(TextBoxSpokeHoleNum.Text) = TextBoxSpokeHoleNum.Text Then
            Else
                ParameterError = True
                MessageBox.Show("请输入正确的轮辐孔的数量！", "格式不正确"， MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If
        Else
            ParameterError = True
            MessageBox.Show("请输入正确的轮辐孔的数量！", "格式不正确"， MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        If IsNumeric(TextBoxD2Axis.Text) Then '判断D2Axis输入的值是否正确
            If TextBoxD2Axis.Text <= 0 Then
                ParameterError = True
                MessageBox.Show("请输入正确的蜗轮轴径！", "格式不正确"， MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If
        Else
            ParameterError = True
            MessageBox.Show("请输入正确的蜗轮轴径！", "格式不正确"， MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        If IsNumeric(TextBoxRCast.Text) Then '判断RCast输入的值是否正确
            If TextBoxRCast.Text < 0 Then
                ParameterError = True
                MessageBox.Show("请输入正确的铸造圆角半径！", "格式不正确"， MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If
        Else
            ParameterError = True
            MessageBox.Show("请输入正确的铸造圆角半径！", "格式不正确"， MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
    End Sub
    '参数计算
    Public Sub NumericalCalculation() '参数计算
        SolidworksEdition = ComboboxEdition.Text
        FilePath = TextBoxFilePath.Text
        Direction = ComboBoxDirection.Text
        m = ComboBoxm.Text
        z1 = ComboBoxz1.Text
        z2 = Int(ComboBoxz2.Text)
        i = z2 / z1
        d1 = ComboBoxd1.Text / 1000
        d2 = m * z2 / 1000
        q = d1 * 1000 / m
        x2 = TextBoxx2.Text
        dj1 = (q + 2 * x2) * m / 1000
        dj2 = d2
        a = (d1 + d2 + 2 * x2 * m / 1000) / 2
        hax = ComboBoxhax.Text
        cx = ComboBoxcx.Text
        RTooth = ComboBoxRTooth.Text / 1000
        da1 = d1 + 2 * m * hax / 1000
        da2 = d2 + 2 * m * hax / 1000
        df1 = d1 - 2 * m * (hax + cx) / 1000
        df2 = d2 - 2 * m * (hax + cx) / 1000
        ha1 = m * hax / 1000
        hf1 = m * (hax + cx) / 1000
        h1 = ha1 + hf1
        ha2 = m * hax / 1000
        hf2 = m * (hax + cx) / 1000
        h2 = ha2 + hf2
        Alpha = ComboBoxalpha.Text * PI / 180
        Gamma = Atan(z1 / q)
        db2 = d2 * Cos(Alpha)
        If z1 = 1 Or z1 = 2 Then '蜗杆宽度
            b1 = m * (12 + 0.1 * z2) / 1000
        Else
            b1 = m * (13 + 0.1 * z2) / 1000
        End If
        If z1 <= 3 Then '蜗轮宽度
            b2 = 0.75 * da1
        Else
            b2 = 0.67 * da1
        End If
        '蜗轮建模参数计算
        rg2 = a - 0.5 * da2
        rf2 = 0.5 * da1 + 0.2 * m / 1000
        k = 2 * m / 1000
        If z1 = 1 Then '外圆
            dm2 = da2 + 2 * m / 1000
        ElseIf z1 = 2 Or z1 = 3 Then
            dm2 = da2 + 1.5 * m / 1000
        Else
            dm2 = da2 + m / 1000
        End If
        d11 = df1
        l1 = b1 / 2
        D2Axis = TextBoxD2Axis.Text
        Call Keyway()
        D2Axis = TextBoxD2Axis.Text / 1000
        D2Hub = 1.6 * D2Axis
        L2Hub = 1.8 * D2Axis
        D2Rim = (a - rf2 - k) * 2
        B2Spoke = b2 / 2
        Select Case ComboBoxWormGearType.Text'蜗轮副板孔径
            Case "整体式"
                D2Spoke = (D2Rim + D2Hub) / 2
                D2SpokeHole = (D2Rim - D2Hub) / 4
            Case "螺栓连接式"
                D2Spoke = (D2Rim + D2Hub) / 2
                D2SpokeHole = (D2Rim - D2Hub) / 4
                BoltD = D2SpokeHole * 1000
                Call BoltNutWasher()
                D2SpokeHole = BoltD
            Case "镶铸式"
                InlayTypeWormGearKeyNum = Int(TextBoxInlayTypeWormGearKeyNum.Text)
                D2Spoke = (D2Rim + D2Hub - 2 * k) / 2
                D2SpokeHole = (D2Rim - 2 * k - D2Hub) / 4
            Case "轮箍式"
                InlayTypeWormGearKeyNum = Int(TextBoxInlayTypeWormGearKeyNum.Text)
                D2Spoke = (D2Rim + D2Hub - 2 * k) / 2
                D2SpokeHole = (D2Rim - 2 * k - D2Hub) / 4
                If TextBoxNutBoltNum.Enabled = True Then
                    NutBoltD = 1.2 * m
                    Call NutBolt()
                    NutBoltHoleL = 3.5 * NutBoltD
                    NutBoltL = 3 * NutBoltD
                    NutBoltNum = Int(TextBoxNutBoltNum.Text)
                End If
        End Select
        SpokeHoleNum = Int(TextBoxSpokeHoleNum.Text)
        RCast = TextBoxRCast.Text / 1000
    End Sub

    '单击创建蜗杆
    Private Sub ButttonCreateWorm_Click(sender As Object, e As EventArgs) Handles ButttonCreateWorm.Click
        Call NumericalCheck()
        If ParameterError = True Then Exit Sub
        Call NumericalCalculation()
        Me.Hide()
        Select Case ComboBoxWormType.Text
            Case "阿基米德圆柱蜗杆(ZA型)"
                Call CreateZATypeWorm()
            Case "法向直廓圆柱蜗杆(ZN型)"
                Call CreateZNTypeWorm()
        End Select
        Me.Show()
    End Sub
    '创建阿基米德圆柱ZA蜗杆
    Public Sub CreateZATypeWorm()
        Dim ZATypeWormPath$ = FilePath & "\蜗杆(ZA)" & z1 & "X" & m & ".SLDPRT"
        AssemWormTitle = "蜗杆(ZA)" & z1 & "X" & m & ".SLDPRT"
        AssemWormPath = ZATypeWormPath
        Swapp = CreateObject("Sldworks.application")
        Swapp.Visible = True
        Swapp.NewDocument("C:\ProgramData\SolidWorks\SOLIDWORKS " & SolidworksEdition & "\templates\gb_part.prtdot", 0, 0, 0) '新建零件
        Part = Swapp.ActiveDoc
        Sketchmer = Part.SketchManager
        Featmgr = Part.FeatureManager
        Inputsize = Swapp.GetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate) '输入尺寸值记录
        Swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '输入尺寸值关闭
        Part.Extension.SelectByID2("右视基准面", "PLANE", 0, 0, 0, False, 32, Nothing, 0) '选中右视基准面
        Sketchmer.InsertSketch(True) '新建草图1
        Sketchmer.CreateCircleByRadius(0, 0, 0, da1 / 2) '画圆
        Part.AddDimension2(0, 0, 0) '给圆标尺寸
        Sketchmer.InsertSketch(True) '完成草图1
        Featmgr.FeatureExtrusion2(False, False, False, 6, 0, b1, 0, False, False, False, False, 0, 0, False, False, False, False, True, False, True, 0, 0, False) '拉伸
        Dim Line0, Line1, Line2, Line3, Line4, Arc As SldWorks.SketchSegment
        Dim x1#, x2#, x3#, y1#, y2#
        Sketchmer.InsertSketch(True) '新建草图2
        If ComboBoxDirection.Text = "右旋" Then
            x1 = -(b1 / 2 + PI * m / 1000)
            x2 = -(b1 / 2 + PI * m / 1000 + PI * m / 4000 - hf1 * Tan(Alpha))
            x3 = -(b1 / 2 + PI * m / 1000 + PI * m / 4000 + h1 * 0.5 * Tan(Alpha))
            y1 = df1 / 2
            y2 = d1 / 2 + h1 * 0.5
            Line0 = Sketchmer.CreateCenterLine(x1, y2, 0, x1, y1, 0)
            Line1 = Sketchmer.CreateLine(x1, y1, 0, x2, y1, 0)
            Line2 = Sketchmer.CreateLine(x2, y1, 0, x3, y2, 0)
            Line3 = Sketchmer.CreateLine(x3, y2, 0, x1, y2, 0)
            If RTooth <> 0 Then
                Line1.Select4(False, Nothing)
                Line2.Select4(True, Nothing)
                Arc = Sketchmer.CreateFillet(RTooth, 2) '倒圆角
            End If
            Line1.SelectChain(False, Nothing)
            Line0.Select4(True, Nothing)
            Part.SketchMirror() '镜像
            Line4 = Sketchmer.CreateCenterLine(-(b1 / 2 + PI * m / 1000 - PI * m / 4000 - h1 * Tan(Alpha)), d1 / 2, 0,
                                               -(b1 / 2 + PI * m / 1000 + PI * m / 4000 + h1 * Tan(Alpha)), d1 / 2, 0) '画分度圆
            Sketchmer.SketchTrim(0, -(b1 / 2 + PI * m / 1000 - PI * m / 4000 - h1 * Tan(Alpha)), d1 / 2, 0) '修剪分度圆
            Line4.Select4(False, Nothing)
            Sketchmer.SketchTrim(0, -(b1 / 2 + PI * m / 1000 + PI * m / 4000 + h1 * Tan(Alpha)), d1 / 2, 0) '修剪分度圆
            Line0.Select4(False, Nothing)
            Line2.Select4(True, Nothing)
            Part.AddDimension2(-(b1 / 2 + PI * m / 1000 + PI * m / 4000), da1 / 2 + 0.001, 0) '标注压力角
            Line4.Select4(False, Nothing)
            Part.AddDimension2(-(b1 / 2 + PI * m / 1000 + PI * m / 8000), d1 / 2 + 0.001, 0) '标注齿槽宽
            Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
            Line3.Select4(True, Nothing)
            Part.AddDimension2(-(b1 / 2 + PI * m / 1000 + PI * m / 4000 + ha1 * Tan(Alpha) + 0.002), da1 / 4, 0) '标注齿顶圆半径
            Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
            Line0.Select4(True, Nothing)
            Part.AddDimension2(-((b1 / 2 + PI * m / 1000) / 2), da1 / 2 + 0.001, 0) '标注左右距离
            Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
            Line4.Select4(True, Nothing)
            Part.AddDimension2(-(b1 / 2 + PI * m / 1000 + PI * m / 4000 + 0.001), d1 / 4, 0) '标注分度圆半径
            Line1.Select4(False, Nothing)
            Line3.Select4(True, Nothing)
            Part.AddDimension2(-(b1 / 2 + PI * m / 2000), d1 / 2, 0) '标注齿高
            Sketchmer.InsertSketch(True) '完成草图2
            Part.Extension.SelectByID2("", "FACE", -b1 / 2, 0, 0, False, 0, Nothing, 0)
            Sketchmer.InsertSketch(True) '新建草图3
            Sketchmer.CreateCircleByRadius(0, 0, 0, d1 / 2) '画圆
            Part.InsertHelix(True, True, True, False, 2, b1 * 2, PI * m * z1 / 1000, 0, 0, 0) '生成螺旋线
            Part.Extension.SelectByID2("草图2", "SKETCH", 0, 0, 0, False, 1, Nothing, 0) '选中扫描轮廓
            Part.Extension.SelectByID2("螺旋线/涡状线1", "REFERENCECURVES", 0, 0, 0, True, 4, Nothing, 0) '选中引导线
            Part.FeatureManager.InsertCutSwept4(False, True, 0, False, False, 1, 0, False, 0, 0, 0, 0, True, True, 0, True, True, True, False) '扫描切除
        ElseIf ComboBoxDirection.Text = "左旋" Then
            x1 = (b1 / 2 + PI * m / 1000)
            x2 = (b1 / 2 + PI * m / 1000 + PI * m / 4000 - hf1 * Tan(Alpha))
            x3 = (b1 / 2 + PI * m / 1000 + PI * m / 4000 + h1 * 0.5 * Tan(Alpha))
            y1 = df1 / 2
            y2 = d1 / 2 + h1 * 0.5
            Line0 = Sketchmer.CreateCenterLine(x1, y2, 0, x1, y1, 0)
            Line1 = Sketchmer.CreateLine(x1, y1, 0, x2, y1, 0)
            Line2 = Sketchmer.CreateLine(x2, y1, 0, x3, y2, 0)
            Line3 = Sketchmer.CreateLine(x3, y2, 0, x1, y2, 0)
            If RTooth <> 0 Then
                Line1.Select4(False, Nothing)
                Line2.Select4(True, Nothing)
                Arc = Sketchmer.CreateFillet(RTooth, 2) '倒圆角
            End If
            Line1.SelectChain(False, Nothing)
            Line0.Select4(True, Nothing)
            Part.SketchMirror() '镜像
            Line4 = Sketchmer.CreateCenterLine((b1 / 2 + PI * m / 1000 - PI * m / 4000 - h1 * Tan(Alpha)), d1 / 2, 0,
                                               (b1 / 2 + PI * m / 1000 + PI * m / 4000 + h1 * Tan(Alpha)), d1 / 2, 0) '画分度圆
            Sketchmer.SketchTrim(0, (b1 / 2 + PI * m / 1000 - PI * m / 4000 - h1 * Tan(Alpha)), d1 / 2, 0) '修剪分度圆
            Line4.Select4(False, Nothing)
            Sketchmer.SketchTrim(0, (b1 / 2 + PI * m / 1000 + PI * m / 4000 + h1 * Tan(Alpha)), d1 / 2, 0) '修剪分度圆
            Line0.Select4(False, Nothing)
            Line2.Select4(True, Nothing)
            Part.AddDimension2((b1 / 2 + PI * m / 1000 + PI * m / 4000), da1 / 2 + 0.001, 0) '标注压力角
            Line4.Select4(False, Nothing)
            Part.AddDimension2((b1 / 2 + PI * m / 1000 + PI * m / 8000), d1 / 2 + 0.001, 0) '标注齿槽宽
            Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
            Line3.Select4(True, Nothing)
            Part.AddDimension2((b1 / 2 + PI * m / 1000 + PI * m / 4000 + ha1 * Tan(Alpha) + 0.002), da1 / 4, 0) '标注齿顶圆半径
            Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
            Line0.Select4(True, Nothing)
            Part.AddDimension2(((b1 / 2 + PI * m / 1000) / 2), da1 / 2 + 0.001, 0) '标注左右距离
            Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
            Line4.Select4(True, Nothing)
            Part.AddDimension2((b1 / 2 + PI * m / 1000 + PI * m / 4000 + 0.001), d1 / 4, 0) '标注分度圆半径
            Line1.Select4(False, Nothing)
            Line3.Select4(True, Nothing)
            Part.AddDimension2((b1 / 2 + PI * m / 2000), d1 / 2, 0) '标注齿高
            Sketchmer.InsertSketch(True) '完成草图2
            Part.Extension.SelectByID2("", "FACE", -b1 / 2, 0, 0, False, 0, Nothing, 0)
            Sketchmer.InsertSketch(True) '新建草图3
            Sketchmer.CreateCircleByRadius(0, 0, 0, d1 / 2) '画圆
            Part.InsertHelix(True, False, True, False, 2, b1 * 2, PI * m * z1 / 1000, 0, 0, Gamma) '生成螺旋线
            Part.Extension.SelectByID2("草图2", "SKETCH", 0, 0, 0, False, 1, Nothing, 0) '选中扫描轮廓
            Part.Extension.SelectByID2("螺旋线/涡状线1", "REFERENCECURVES", 0, 0, 0, True, 4, Nothing, 0) '选中引导线
            Part.FeatureManager.InsertCutSwept4(False, True, 0, False, False, 1, 0, False, 0, 0, 0, 10, True, True, 0, True, True, True, False) '扫描切除
        End If
        Part.Extension.SelectByID2("螺旋线/涡状线1", "REFERENCECURVES", 0, 0, 0, False, 0, Nothing, 0) '选中引导线
        Part.BlankRefGeom() '隐藏引导线
        Part.Extension.SelectByID2("草图3", "SKETCH", 0, 0, 0, False, 0, Nothing, 0) '选中草图3
        Part.BlankSketch() '隐藏草图
        Part.Extension.SelectByID2("", "FACE", -b1 / 2, 0, 0, False, 0, Nothing, 0) '选中蜗杆左端面
        Sketchmer.InsertSketch(True) '新建草图4
        Sketchmer.CreateCircleByRadius(0, 0, 0, d11 / 2) '画圆
        Part.AddDimension2(0, 0, 0) '给圆标尺寸
        Sketchmer.InsertSketch(True) '完成草图4
        Featmgr.FeatureExtrusion2(True, False, False, 0, 0, l1, 0, False, False, False, False, 0, 0, False, False, False, False, True, False, True, 0, 0, False) '拉伸
        Part.Extension.SelectByID2("", "FACE", b1 / 2, 0, 0, False, 0, Nothing, 0) '选中蜗杆右端面
        Sketchmer.InsertSketch(True) '新建草图5
        Sketchmer.CreateCircleByRadius(0, 0, 0, d11 / 2) '画圆
        Part.AddDimension2(0, 0, 0) '给圆标尺寸
        Sketchmer.InsertSketch(True) '完成草图5
        Featmgr.FeatureExtrusion2(True, False, False, 0, 0, l1, 0, False, False, False, False, 0, 0, False, False, False, False, True, False, True, 0, 0, False) '拉伸
        If z1 <> 1 Then
            Part.Extension.SelectByID2("切除-扫描1", "BODYFEATURE", 0, 0, 0, False, 4, Nothing, 0)
            Part.Extension.SelectByID2("", "EDGE", b1 / 2 + l1, d11 / 2, 0, True, 1, Nothing, 0)
            Part.FeatureManager.FeatureCircularPattern4(z1, 6.2831853071796, False, "NULL", False, True, False)
        End If
        Part.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中上视基准面
        Part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, True, 0, Nothing, 0) '选中前视基准面
        Part.InsertAxis2(True) '新建基准轴1
        Part.Extension.SelectByID2("基准轴1", "AXIS", 0, 0, 0, False, 0, Nothing, 0)
        Part.BlankRefGeom() '隐藏基准轴1
        Part.AddConfiguration2("工程图", "", "", True, False, False, True, 256)
        Part.Extension.SelectByID2("切除-扫描1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        Part.EditSuppress2()
        Part.ShowConfiguration2("默认")
        Part.EditRebuild3() '重建模型
        Part.ClearSelection2（True）
        Part.ShowNamedView2("*等轴测", 7) '正等测视角
        Part.ViewZoomtofit2() '整屏显示
        Swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, Inputsize) '输入尺寸值还原
        Part.SaveAs3(ZATypeWormPath, 0, 2) '文件保存
        If CheckBoxCloseFile.Checked = True Then Swapp.CloseAllDocuments(True)
    End Sub
    '创建法向直廓圆柱ZN蜗杆
    Public Sub CreateZNTypeWorm()
        Dim ZNTypeWormPath$ = FilePath & "\蜗杆(ZN)" & z1 & "X" & m & ".SLDPRT"
        AssemWormTitle = "蜗杆(ZN)" & z1 & "X" & m & ".SLDPRT"
        AssemWormPath = ZNTypeWormPath
        Swapp = CreateObject("Sldworks.application")
        Swapp.Visible = True
        Swapp.NewDocument("C:\ProgramData\SolidWorks\SOLIDWORKS " & SolidworksEdition & "\templates\gb_part.prtdot", 0, 0, 0) '新建零件
        Part = Swapp.ActiveDoc
        Sketchmer = Part.SketchManager
        Featmgr = Part.FeatureManager
        Inputsize = Swapp.GetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate) '输入尺寸值记录
        Swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '输入尺寸值关闭
        Part.Extension.SelectByID2("右视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        Part.FeatureManager.InsertRefPlane(8, b1, 0, 0, 0, 0) '新建基准面1
        Part.Extension.SelectByID2("基准面1", "PLANE", 0, 0, 0, False, 32, Nothing, 0)
        Sketchmer.InsertSketch(True) '新建草图1
        Sketchmer.CreateCircleByRadius(0, 0, 0, da1 / 2) '画圆
        Part.AddDimension2(0, 0, 0) '给圆标尺寸
        Sketchmer.InsertSketch(True) '完成草图1
        Featmgr.FeatureExtrusion2(False, False, False, 6, 0, b1, 0, False, False, False, False, 0, 0, False, False, False, False, True, False, True, 0, 0, False) '拉伸
        Part.ShowNamedView2("*等轴测", 7) '正等测视角
        Part.ViewZoomtofit2() '整屏显示
        Dim Line0, Line1, Line2, Line3, Line4, Arc As SldWorks.SketchSegment
        Dim x1#, x2#, x3#, y1#, y2#
        Part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 32, Nothing, 0)
        Sketchmer.InsertSketch(True) '新建草图2
        x1 = 0
        x2 = -(PI * m / 4000 - hf1 * Tan(Alpha))
        x3 = -(PI * m / 4000 + h1 * 0.5 * Tan(Alpha))
        y1 = df1 / 2
        y2 = d1 / 2 + h1 * 0.5
        Line0 = Sketchmer.CreateCenterLine(x1, y2, 0, x1, y1, 0)
        Line1 = Sketchmer.CreateLine(x1, y1, 0, x2, y1, 0)
        Line2 = Sketchmer.CreateLine(x2, y1, 0, x3, y2, 0)
        Line3 = Sketchmer.CreateLine(x3, y2, 0, x1, y2, 0)
        If RTooth <> 0 Then
            Line1.Select4(False, Nothing)
            Line2.Select4(True, Nothing)
            Arc = Sketchmer.CreateFillet(RTooth, 2) '倒圆角
        End If
        Line1.SelectChain(False, Nothing)
        Line0.Select4(True, Nothing)
        Part.SketchMirror() '镜像
        Line4 = Sketchmer.CreateCenterLine(-(-PI * m / 4000 - h1 * Tan(Alpha)), d1 / 2, 0, -(PI * m / 4000 + h1 * Tan(Alpha)), d1 / 2, 0) '画分度圆
        Sketchmer.SketchTrim(0, -(-PI * m / 4000 - h1 * Tan(Alpha)), d1 / 2, 0) '修剪分度圆
        Line4.Select4(False, Nothing)
        Sketchmer.SketchTrim(0, -(+PI * m / 4000 + h1 * Tan(Alpha)), d1 / 2, 0) '修剪分度圆
        Line0.Select4(False, Nothing)
        Line2.Select4(True, Nothing)
        Part.AddDimension2(-(PI * m / 1000 + PI * m / 4000), da1 / 2 + 0.001, 0) '标注压力角
        Line4.Select4(False, Nothing)
        Part.AddDimension2(-(PI * m / 1000 + PI * m / 8000), d1 / 2 + 0.001, 0) '标注齿槽宽
        Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
        Line3.Select4(True, Nothing)
        Part.AddDimension2(-(PI * m / 1000 + PI * m / 4000 + ha1 * Tan(Alpha) + 0.002), da1 / 4, 0) '标注齿顶圆半径
        Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
        Line4.Select4(True, Nothing)
        Part.AddDimension2(-(PI * m / 1000 + PI * m / 4000 + 0.001), d1 / 4, 0) '标注分度圆半径
        Line1.Select4(False, Nothing)
        Line3.Select4(True, Nothing)
        Part.AddDimension2(-(PI * m / 2000), d1 / 2, 0) '标注齿高
        Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
        Line0.Select4(True, Nothing)
        Part.SketchAddConstraints("sgCOINCIDENT")
        Sketchmer.InsertSketch(True) '完成草图2
        Part.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, False, 32, Nothing, 0)
        Sketchmer.InsertSketch(True) '新建草图3
        Part.ShowNamedView2("*上视", 5)
        Part.ViewZoomtofit2() '整屏显示
        If ComboBoxDirection.Text = "右旋" Then
            Sketchmer.CreateLine(0, 0, 0, d1 / 2 * Cos(Gamma), d1 / 2 * Sin(Gamma), 0) '斜线
        ElseIf ComboBoxDirection.Text = "左旋" Then
            Sketchmer.CreateLine(0, 0, 0, d1 / 2 * Cos(Gamma), -d1 / 2 * Sin(Gamma), 0) '斜线
        End If
        Sketchmer.InsertSketch(True) '完成草图3
        Part.ShowNamedView2("*等轴测", 7) '正等测视角
        Part.ViewZoomtofit2() '整屏显示
        Part.Extension.SelectByID2("Line1@草图3", "EXTSKETCHSEGMENT", 0, 0, 0, True, 0, Nothing, 0)
        Part.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
        Part.FeatureManager.InsertRefPlane(4, 0, 2, 0, 0, 0) '新建基准面2
        Part.Extension.SelectByID2("基准面2", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        Sketchmer.InsertSketch(True) '新建草图4
        Part.Extension.SelectByID2("草图2", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
        Part.SketchManager.SketchUseEdge3(False, False)
        Sketchmer.InsertSketch(True) '完成草图4
        Part.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, False, 32, Nothing, 0)
        Sketchmer.InsertSketch(True) '新建草图5
        Sketchmer.CreateLine(0, 0, 0, 2 * b1, 0, 0) '直线
        Sketchmer.InsertSketch(True) '完成草图5
        Part.Extension.SelectByID2("草图4", "SKETCH", 0, 0, 0, False, 1, Nothing, 0)
        Part.Extension.SelectByID2("草图5", "SKETCH", 0, 0, 0, True, 4, Nothing, 0)
        Part.FeatureManager.InsertCutSwept4(False, False, 8, False, False, 1, 0, False, 0, 0, 0, 0, True, True, 2 * 2000 * b1 / m, True, True, True, False)
        Part.Extension.SelectByID2("右视基准面", "PLANE", 0, 0, 0, False, 32, Nothing, 0)
        Sketchmer.InsertSketch(True) '新建草图6
        Sketchmer.CreateCircleByRadius(0, 0, 0, df1 / 2) '画圆
        Part.FeatureManager.FeatureExtrusion2(True, False, False, 0, 0, b1 * 2, b1 * 2, False, False, False, False, 0, 0, False, False, False, False, True, True, True, 0, 0, False)
        Part.Extension.SelectByID2("基准面1", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        Part.Extension.SelectByID2("基准面2", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
        Part.BlankRefGeom()
        Part.Extension.SelectByID2("草图2", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
        Part.Extension.SelectByID2("草图3", "SKETCH", 0, 0, 0, True, 0, Nothing, 0)
        Part.BlankSketch()
        Part.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中上视基准面
        Part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, True, 0, Nothing, 0) '选中前视基准面
        Part.InsertAxis2(True) '新建基准轴1
        Part.Extension.SelectByID2("基准轴1", "AXIS", 0, 0, 0, False, 0, Nothing, 0)
        Part.BlankRefGeom() '隐藏基准轴1
        Part.AddConfiguration2("工程图", "", "", True, False, False, True, 256)
        Part.Extension.SelectByID2("切除-扫描1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        Part.EditSuppress2()
        Part.ShowConfiguration2("默认")
        Part.EditRebuild3() '重建模型
        Part.ClearSelection2（True）
        Part.ShowNamedView2("*等轴测", 7) '正等测视角
        Part.ViewZoomtofit2() '整屏显示
        Swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, Inputsize) '输入尺寸值还原
        Part.SaveAs3(ZNTypeWormPath, 0, 2) '文件保存
        If CheckBoxCloseFile.Checked = True And Direction = "右旋" Then Swapp.CloseAllDocuments(True)
        If Direction = "左旋" Then MsgBox("请自行修改切除-扫描1，保证旋向为左旋。", 0 + 64, "提示")
    End Sub

    '选择蜗轮类型
    Private Sub ComboBoxWormGearType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxWormGearType.SelectedIndexChanged, ComboBoxWormGearTypeAssem.SelectedIndexChanged
        Select Case ComboBoxWormGearType.Text
            Case "整体式"
                GropeBoxInlayType.Enabled = False
                GroupBoxTyreType.Enabled = False
            Case "镶铸式"
                GropeBoxInlayType.Enabled = True
                GroupBoxTyreType.Enabled = False
            Case "螺栓连接式"
                GropeBoxInlayType.Enabled = False
                GroupBoxTyreType.Enabled = False
            Case "轮箍式"
                GropeBoxInlayType.Enabled = False
                GroupBoxTyreType.Enabled = True
        End Select
    End Sub
    '单击创建蜗轮
    Private Sub ButtonCreateWormGear_Click(sender As Object, e As EventArgs) Handles ButtonCreateWormGear.Click
        Call NumericalCheck()
        If ParameterError = True Then Exit Sub
        Call NumericalCalculation()
        Me.Hide()
        Select Case ComboBoxWormGearType.Text
            Case "整体式"
                Call CreateIntegralTypeWormGear()
            Case "螺栓连接式"
                Call CreateBoltedTypeWormGear()
            Case "镶铸式"
                Call CreateInlayTypeWormGear()
            Case "轮箍式"
                Call CreateTyreTypeWormGear()
        End Select
        Me.Show()
    End Sub
    '创建整体式蜗轮
    Public Sub CreateIntegralTypeWormGear() '创建整体式蜗轮
        Dim IntegralTypWormGearRimPath$ = FilePath & "\蜗轮(整体式)" & z2 & "X" & m & ".SLDPRT"
        AssemWormGearTitle = "蜗轮(整体式)" & z2 & "X" & m & ".SLDPRT"
        AssemWormGearPath = IntegralTypWormGearRimPath
        Swapp = CreateObject("Sldworks.application")
        Swapp.Visible = True
        Swapp.NewDocument("C:\ProgramData\SolidWorks\SOLIDWORKS " & SolidworksEdition & "\templates\gb_part.prtdot", 0, 0, 0) '新建零件
        Part = Swapp.ActiveDoc
        Sketchmer = Part.SketchManager
        Featmgr = Part.FeatureManager
        Inputsize = Swapp.GetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate) '输入尺寸值记录
        Swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '输入尺寸值关闭
        Part.Extension.SelectByID2("右视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中右视基准面
        Part.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, True, 0, Nothing, 0) '选中上视基准面
        Part.InsertAxis2(True) '新建基准轴1
        Part.Extension.SelectByID2("右视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中右视基准面
        Dim Line0, Line1, Line2, Line3, Line4, Line5, Line6, Line7, Line8 As SldWorks.SketchSegment
        Dim Arc1, Arc2, Arc3, Arc4, Arc5 As SldWorks.SketchSegment
        Sketchmer.InsertSketch(True) '新建草图1
        Line1 = Sketchmer.CreateLine(0, D2Axis / 2, 0, L2Hub / 2, D2Axis / 2, 0) '孔径直线
        Line2 = Sketchmer.CreateLine(L2Hub / 2, D2Axis / 2, 0, L2Hub / 2, D2Hub / 2, 0) '轮毂侧面
        Line3 = Sketchmer.CreateLine(L2Hub / 2, D2Hub / 2, 0, B2Spoke / 2, D2Hub / 2, 0) '轮毂外圆
        Line4 = Sketchmer.CreateLine(B2Spoke / 2, D2Hub / 2, 0, B2Spoke / 2, dm2 / 2, 0) '凹槽面
        Arc2 = Sketchmer.CreateCircleByRadius(0, a, 0, rf2 + k) '画大圆
        Line4.Select4(False, Nothing) '选中凹槽面直线
        Sketchmer.SketchTrim(0, B2Spoke / 2, dm2 / 2, 0) '切割凹槽面直线
        Line5 = Sketchmer.CreateLine(0, dm2 / 2, 0, b2 / 2, dm2 / 2, 0) '最大圆直线
        Line6 = Sketchmer.CreateLine(b2 / 2, dm2 / 2, 0, b2 / 2, dm2 / 4, 0) '齿宽侧面
        Sketchmer.SketchTrim(0, b2 / 2, dm2 / 4, 0) '修剪齿宽侧面
        Arc2.Select4(False, Nothing) '选中大圆
        Sketchmer.SketchTrim(0, 0, a + rf2 + k, 0) '修剪大圆
        Arc1 = Sketchmer.CreateCircleByRadius(0, a, 0, rg2) '画小圆
        Line5.Select4(False, Nothing) '选中最大圆直线
        Sketchmer.SketchTrim(0, 0, dm2 / 2, 0) '修剪最大圆直线
        Line0 = Sketchmer.CreateCenterLine(0, 0, 0, 0, a - rg2, 0) '画中心线
        Arc1.Select4(False, Nothing) '选中大圆
        Sketchmer.SketchTrim(0, 0, a + rg2, 0) '修剪小圆
        If RCast <> 0 Then
            Line3.Select4(False, Nothing)
            Line4.Select4(True, Nothing)
            Arc3 = Sketchmer.CreateFillet(RCast, 2) '倒圆角
            Arc2.Select4(False, Nothing)
            Line4.Select4(True, Nothing)
            Arc4 = Sketchmer.CreateFillet(RCast, 2) '倒圆角
        End If
        Line5.Select4(False, Nothing)
        Line6.Select4(True, Nothing)
        Line7 = Sketchmer.CreateChamfer(1, m / 1000, m / 1000) '倒角
        Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
        Line1.Select4(True, Nothing)
        Part.AddDimension2(0, D2Axis / 4, L2Hub / 2 + m / 1000) '标注轴孔半径
        Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
        Line3.Select4(True, Nothing)
        Part.AddDimension2(0, D2Hub / 4, L2Hub / 2 + 2 * m / 1000) '标注轮毂外径
        Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
        Line5.Select4(True, Nothing)
        Part.AddDimension2(0, D2Hub / 4, L2Hub / 2 + 3 * m / 1000) '标注最大半径
        Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
        Line2.Select4(True, Nothing)
        Part.AddDimension2(0, D2Axis / 2 - 2 * m / 1000, L2Hub / 4) '标注轮廓宽度的一半
        Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
        Line4.Select4(True, Nothing)
        Part.AddDimension2(0, D2Axis / 2 - m / 1000, L2Hub / 4) '标注凹槽宽度的一半
        Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
        Line6.Select4(True, Nothing)
        Part.AddDimension2(0, D2Axis / 2 - 3 * m / 1000, L2Hub / 4) '标注齿宽
        Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
        Arc1.Select4(True, Nothing)
        Part.AddDimension2(0, D2Hub / 4, L2Hub / 2 + 4 * m / 1000) '标注中心距
        Arc2.Select4(True, Nothing)
        Part.AddDimension2(0, dm2 / 2, 0) '给大圆标尺寸
        Arc1.Select4(True, Nothing)
        Part.AddDimension2(-b1 / 2, dm2 / 2, 0) '给小圆标尺寸
        Line1.SelectChain(False, Nothing)
        Line0.Select4(True, Nothing)
        Part.SketchMirror() '镜像
        Part.Extension.SelectByID2("基准轴1", "AXIS", 0, 0, 0, False, 16, Nothing, 0)
        Part.FeatureManager.FeatureRevolve2(True, True, False, False, False, False, 0, 0, 2 * PI, 0, False, False, 0.01, 0.01, 0, 0, 0, True, True, True) '旋转
        Part.ShowNamedView2("*等轴测", 7) '正等测视角
        Part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 32, Nothing, 0) '选中前视基准面
        Sketchmer.InsertSketch(True) '新建草图2
        Sketchmer.CreateCircleByRadius(0, D2Spoke / 2, 0, D2SpokeHole / 2) '画圆
        Featmgr.FeatureCut3(False, False, False, 9, 1, 0.01, 0.01, False, False, False, False, 0, 0, False, False, False, False, False, True, True, True, True, False, 0, 0, False) '拉伸切除
        Part.Extension.SelectByID2("切除-拉伸1", "BODYFEATURE", 0, 0, 0, False, 4, Nothing, 0)
        Part.Extension.SelectByID2("基准轴1", "AXIS", 0, 0, 0, True, 1, Nothing, 0)
        Part.FeatureManager.FeatureCircularPattern4(SpokeHoleNum, 2 * PI, False, "NULL", False, True, False)
        Part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 32, Nothing, 0) '选中前视基准面
        Sketchmer.InsertSketch(True) '新建草图3
        Part.ShowNamedView2("*前视", 7) '前视角
        Part.ViewZoomtofit2() '整屏显示
        Dim t1#, t2#
        t1 = 0
        t2 = Tan(Acos(db2 / dm2))
        Arc1 = Sketchmer.CreateCircleByRadius(0, 0, 0, df2 / 2) '齿根圆
        Arc2 = Sketchmer.CreateCircleByRadius(0, 0, 0, dm2 / 2) '齿顶圆
        Line1 = Sketchmer.CreateEquationSpline2(db2 * 1000 / 2 & "*(sin(t)-t*cos(t))", db2 * 1000 / 2 & "*(cos(t)+t*sin(t))", "", t1, t2, False, 0, 0, 0, True, True) '绘制渐开线
        Line8 = Sketchmer.CreateCenterLine(0, 0, 0, 0.005, 0.005, 0) '绘制中心线
        Line8.Angle = PI / 2 + (Acos(db2 / d2) - Tan(Acos(db2 / d2)) + PI * 0.5 / z2) '中心线角度赋值
        If db2 > df2 Then
            Line2 = Sketchmer.CreateLine(0, db2 / 2, 0, 0, df2 / 2, 0) '绘制过渡曲线
            Line1.Select4(False, Nothing)
            Line2.Select4(True, Nothing)
            Line8.Select4(True, Nothing)
            Part.SketchMirror() '镜像
        Else
            Line1.Select4(False, Nothing)
            Sketchmer.SketchTrim(0, 0, db2 / 2, 0)
            Line1.Select4(False, Nothing)
            Line8.Select4(True, Nothing)
            Part.SketchMirror() '镜像
        End If
        Arc1.Select4(False, Nothing)
        Sketchmer.SketchTrim(0, 0, -df2 / 2, 0)
        Arc2.Select4(False, Nothing)
        Sketchmer.SketchTrim(0, 0, -dm2 / 2, 0)
        Line1.SelectChain(False, Nothing)
        Part.SketchAddConstraints("sgFIXED")
        Sketchmer.InsertSketch(True) '完成草图3
        Part.Extension.SelectByID2("Line1@草图3", "EXTSKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
        Part.Extension.SelectByID2("基准轴1", "AXIS", 0, 0, 0, True, 1, Nothing, 1)
        Featmgr.InsertRefPlane(4, 0, 4, 0, 0, 0) '新建基准面1
        Part.Extension.SelectByID2("基准面1", "PLANE", 0, 0, 0, False, 32, Nothing, 0) '选中基准面1
        Sketchmer.InsertSketch(True) '新建草图4
        Arc5 = Sketchmer.CreateCircleByRadius(0, a, 0, d1 / 2) '画圆
        Part.AddDimension2(0, a, 0) '给圆标直径
        Sketchmer.InsertSketch(True) '完成草图4
        Part.Extension.SelectByID2("草图3", "SKETCH", 0, 0, 0, False, 1, Nothing, 0) '选中扫描轮廓
        Part.Extension.SelectByID2("草图4", "SKETCH", 0, 0, 0, True, 4, Nothing, 0) '选中引导线
        Part.FeatureManager.InsertCutSwept4(False, True, 0, False, False, 0, 0, False, 0, 0, 0, 0, True, True, 0, True, True, True, False) '扫描切除
        Part.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, False, 2, Nothing, 0)
        Part.Extension.SelectByID2("切除-扫描1", "BODYFEATURE", 0, 0, 0, True, 1, Nothing, 0)
        Part.FeatureManager.InsertMirrorFeature(False, False, False, False) '镜像
        Part.Extension.SelectByID2("镜向1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0) '选中镜像
        Part.EditSuppress2() '压缩镜像
        Part.Extension.SelectByID2("切除-扫描1", "BODYFEATURE", 0, 0, 0, False, 4, Nothing, 0)
        Part.Extension.SelectByID2("基准轴1", "AXIS", 0, 0, 0, True, 1, Nothing, 0)
        Part.FeatureManager.FeatureCircularPattern3(z2, 2 * PI, False, "NULL", False, True)
        Part.Extension.SelectByID2("基准轴1", "AXIS", 0, 0, 0, False, 0, Nothing, 0)
        Part.Extension.SelectByID2("基准面1", "PLANE", 0, 0, 0, True, 0, Nothing, 0) '选中基准面1
        Part.BlankRefGeom() '隐藏基准轴1和基准面1
        Part.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, False, 32, Nothing, 0) '选中上视基准面
        Sketchmer.InsertSketch(True) '新建草图5
        Part.ShowNamedView2("*上视", 5)
        Part.ViewZoomtofit2() '整屏显示
        Part.SketchManager.CreateCenterRectangle(0, 0, 0, Keywayb / 2, L2Hub, 0)
        Part.SketchAddConstraints("sgFIXED")
        Sketchmer.InsertSketch(True) '完成草图5
        Select Case ComboBoxKeyWayNum.Text
            Case 1
                Featmgr.FeatureCut3(True, False, True, 0, 0, D2Axis / 2 + Keywayh / 2, D2Axis / 2 + Keywayh / 2, False, False, False, False, 0, 0, False, False, False, False, False, True, True, True, True, False, 0, 0, False) '拉伸切除
            Case 2
                Featmgr.FeatureCut3(False, False, True, 0, 0, D2Axis / 2 + Keywayh / 2, D2Axis / 2 + Keywayh / 2, False, False, False, False, 0, 0, False, False, False, False, False, True, True, True, True, False, 0, 0, False) '拉伸切除
        End Select
        Part.EditRebuild3() '重建模型
        Part.ClearSelection2（True）
        Part.ShowNamedView2("*等轴测", 7) '正等测视角
        Part.ViewZoomtofit2() '整屏显示
        Swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, Inputsize) '输入尺寸值还原
        Part.SaveAs3(IntegralTypWormGearRimPath, 0, 2) '文件保存
        If CheckBoxCloseFile.Checked = True Then Swapp.CloseAllDocuments(True)
    End Sub
    '创建螺栓连接式蜗轮
    Public Sub CreateBoltedTypeWormGear() '创建螺栓连接式蜗轮
        Dim BoltedTypeWormGearRimPath$ = FilePath & "\蜗轮(螺栓连接式轮缘)" & z2 & "X" & m & ".SLDPRT"
        Dim BoltedTypeWormGearHubPath$ = FilePath & "\蜗轮(螺栓连接式轮毂)" & z2 & "X" & m & ".SLDPRT"
        Dim BoltPath$ = FilePath & "\螺栓M" & Int(D2SpokeHole * 1000) & "X" & Int(BoltL * 10000) / 10 & ".SLDPRT"
        Dim NutPath$ = FilePath & "\螺母M" & Int(D2SpokeHole * 1000) & ".SLDPRT"
        Dim WasherPath$ = FilePath & "\垫圈M" & Int(D2SpokeHole * 1000) & ".SLDPRT"
        Dim BoltedTypeWormGearPath$ = FilePath & "\蜗轮(螺栓连接式)" & z1 & "X" & z2 & "X" & m & ".SLDASM"
        AssemWormGearTitle = "蜗轮(螺栓连接式)" & z1 & "X" & z2 & "X" & m & ".SLDASM"
        AssemWormGearPath = BoltedTypeWormGearPath
        Swapp = CreateObject("Sldworks.application")
        Swapp.Visible = True
        Swapp.NewDocument("C:\ProgramData\SolidWorks\SOLIDWORKS " & SolidworksEdition & "\templates\gb_part.prtdot", 0, 0, 0) '新建零件
        Part = Swapp.ActiveDoc
        Sketchmer = Part.SketchManager
        Featmgr = Part.FeatureManager
        Inputsize = Swapp.GetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate) '输入尺寸值记录
        Swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '输入尺寸值关闭
        Part.Extension.SelectByID2("右视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中右视基准面
        Part.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, True, 0, Nothing, 0) '选中上视基准面
        Part.InsertAxis2(True) '新建基准轴1
        Part.Extension.SelectByID2("基准轴1", "AXIS", 0, 0, 0, False, 0, Nothing, 0)
        Part.BlankRefGeom() '隐藏基准轴1
        Part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 32, Nothing, 0) '选中前视基准面
        Sketchmer.InsertSketch(True) '新建草图1
        Part.SketchManager.CreatePolygon(0, 0, 0, BoltNute / 2, 0, 0, 6, True) '创建六边形
        Part.FeatureManager.FeatureExtrusion2(True, False, False, 0, 0, Boltk, Boltk, False, False, False, False, 0, 0, False, False, False, False, True, True, True, 0, 0, False) '拉伸
        Part.Extension.SelectByID2("", "FACE", 0, 0, Boltk, False, 0, Nothing, 0)
        Sketchmer.InsertSketch(True) '新建草图2
        Sketchmer.CreateCircleByRadius(0, 0, 0, BoltNuts / 2) '画圆
        Part.FeatureManager.FeatureCut3(True, True, False, 0, 0, Boltk, Boltk, True, False, False, False, PI / 6, PI / 6, False, False, False, False, False, True, True, True, True, False, 0, 0, False) '拉伸切除
        Part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 32, Nothing, 0) '选中前视基准面
        Sketchmer.InsertSketch(True) '新建草图3
        Sketchmer.CreateCircleByRadius(0, 0, 0, BoltD / 2) '画圆
        Part.FeatureManager.FeatureExtrusion2(True, False, True, 0, 0, BoltL - Boltl2, BoltL - Boltl2, False, False, False, False, 0, 0, False, False, False, False, True, True, True, 0, 0, False) '拉伸
        Part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 32, Nothing, 0) '选中前视基准面
        Sketchmer.InsertSketch(True) '新建草图4
        Sketchmer.CreateCircleByRadius(0, 0, 0, Boltdp / 2) '画圆
        Part.FeatureManager.FeatureExtrusion2(True, False, True, 0, 0, BoltL, BoltL, False, False, False, False, 0, 0, False, False, False, False, True, True, True, 0, 0, False) '拉伸
        Part.EditRebuild3() '重建模型
        Part.ClearSelection2（True）
        Part.ShowNamedView2("*等轴测", 7) '正等测视角
        Part.ViewZoomtofit2() '整屏显示
        Part.SaveAs3(BoltPath, 0, 2) '文件保存
        Swapp.CloseAllDocuments(True) '文件关闭
        Swapp.NewDocument("C:\ProgramData\SolidWorks\SOLIDWORKS " & SolidworksEdition & "\templates\gb_part.prtdot", 0, 0, 0) '新建零件
        Part = Swapp.ActiveDoc
        Sketchmer = Part.SketchManager
        Featmgr = Part.FeatureManager
        Part.Extension.SelectByID2("右视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中右视基准面
        Part.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, True, 0, Nothing, 0) '选中上视基准面
        Part.InsertAxis2(True) '新建基准轴1
        Part.Extension.SelectByID2("基准轴1", "AXIS", 0, 0, 0, False, 0, Nothing, 0)
        Part.BlankRefGeom() '隐藏基准轴1
        Part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 32, Nothing, 0) '选中前视基准面
        Sketchmer.InsertSketch(True) '新建草图1
        Part.SketchManager.CreatePolygon(0, 0, 0, BoltNute / 2, 0, 0, 6, True) '创建六边形
        Part.FeatureManager.FeatureExtrusion2(True, False, False, 0, 0, Nutm, Nutm, False, False, False, False, 0, 0, False, False, False, False, True, True, True, 0, 0, False) '拉伸
        Part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 32, Nothing, 0) '选中前视基准面
        Sketchmer.InsertSketch(True) '新建草图2
        Sketchmer.CreateCircleByRadius(0, 0, 0, BoltNuts / 2) '画圆
        Part.FeatureManager.FeatureCut3(True, True, True, 0, 0, Nutm, Nutm, True, False, False, False, PI / 6, PI / 6, False, False, False, False, False, True, True, True, True, False, 0, 0, False) '拉伸切除
        Part.Extension.SelectByID2("", "FACE", 0, 0, Nutm, False, 0, Nothing, 0)
        Sketchmer.InsertSketch(True) '新建草图3
        Sketchmer.CreateCircleByRadius(0, 0, 0, BoltNuts / 2) '画圆
        Part.FeatureManager.FeatureCut3(True, True, False, 0, 0, Nutm, Nutm, True, False, False, False, PI / 6, PI / 6, False, False, False, False, False, True, True, True, True, False, 0, 0, False) '拉伸切除
        Part.Extension.SelectByID2("", "FACE", 0, 0, Nutm, False, 0, Nothing, 0)
        Part.FeatureManager.InsertRefPlane(4, 0, 0, 0, 0, 0) '新建基准面1
        Part.Extension.SelectByID2("基准面1", "PLANE", 0, 0, 0, True, 0, Nothing, 0) '选中基准面1
        Part.BlankRefGeom() '隐藏基准面1
        Part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 32, Nothing, 0) '选中前视基准面
        Sketchmer.InsertSketch(True) '新建草图4
        Sketchmer.CreateCircleByRadius(0, 0, 0, BoltD / 2) '画圆
        Part.FeatureManager.FeatureCut3(True, False, True, 1, 0, 0.01, 0.01, False, False, False, False, 0, 0, False, False, False, False, False, True, True, True, True, False, 0, 0, False) '拉伸切除
        Part.EditRebuild3() '重建模型
        Part.ClearSelection2（True）
        Part.ShowNamedView2("*等轴测", 7) '正等测视角
        Part.ViewZoomtofit2() '整屏显示
        Part.SaveAs3(NutPath, 0, 2) '文件保存
        Swapp.CloseAllDocuments(True) '文件关闭
        Swapp.NewDocument("C:\ProgramData\SolidWorks\SOLIDWORKS " & SolidworksEdition & "\templates\gb_part.prtdot", 0, 0, 0) '新建零件
        Part = Swapp.ActiveDoc
        Sketchmer = Part.SketchManager
        Featmgr = Part.FeatureManager
        Part.Extension.SelectByID2("右视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中右视基准面
        Part.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, True, 0, Nothing, 0) '选中上视基准面
        Part.InsertAxis2(True) '新建基准轴1
        Part.Extension.SelectByID2("基准轴1", "AXIS", 0, 0, 0, False, 0, Nothing, 0)
        Part.BlankRefGeom() '隐藏基准轴1
        Part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 32, Nothing, 0) '选中前视基准面
        Sketchmer.InsertSketch(True) '新建草图1
        Sketchmer.CreateCircleByRadius(0, 0, 0, Washerd2 / 2) '画圆
        Part.ViewZoomtofit2() '整屏显示
        Part.AddDimension2(Washerd2, 0, 0) '标尺寸
        Sketchmer.CreateCircleByRadius(0, 0, 0, Washerd1 / 2) '画圆
        Part.AddDimension2(Washerd1, 0, 0) '标尺寸
        Part.FeatureManager.FeatureExtrusion2(True, False, False, 0, 0, Washerh / 2, Washerh / 2, False, False, False, False, 0, 0, False, False, False, False, True, True, True, 0, 0, False) '拉伸
        Part.Extension.SelectByID2("", "FACE", 0, (Washerd1 + Washerd2) / 4, Washerh / 2, False, 0, Nothing, 0)
        Part.FeatureManager.InsertRefPlane(4, 0, 0, 0, 0, 0) '新建基准面1
        Part.Extension.SelectByID2("基准面1", "PLANE", 0, 0, 0, True, 0, Nothing, 0) '选中基准面1
        Part.BlankRefGeom() '隐藏基准面1
        Part.EditRebuild3() '重建模型
        Part.ClearSelection2（True）
        Part.ShowNamedView2("*等轴测", 7) '正等测视角
        Part.ViewZoomtofit2() '整屏显示
        Part.SaveAs3(WasherPath, 0, 2) '文件保存
        Swapp.CloseAllDocuments(True) '文件关闭
        Swapp.NewDocument("C:\ProgramData\SolidWorks\SOLIDWORKS " & SolidworksEdition & "\templates\gb_part.prtdot", 0, 0, 0) '新建零件
        Part = Swapp.ActiveDoc
        Sketchmer = Part.SketchManager
        Featmgr = Part.FeatureManager
        Part.Extension.SelectByID2("右视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中右视基准面
        Part.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, True, 0, Nothing, 0) '选中上视基准面
        Part.InsertAxis2(True) '新建基准轴1
        Part.Extension.SelectByID2("右视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中右视基准面
        Dim Line0, Line1, Line2, Line3, Line4, Line5, Line6, Line7, Line8 As SldWorks.SketchSegment
        Dim Arc0, Arc1, Arc2, Arc3, Arc4, Arc5 As SldWorks.SketchSegment
        Sketchmer.InsertSketch(True) '新建草图1
        Line1 = Sketchmer.CreateLine(0, D2Axis / 2, 0, L2Hub / 2, D2Axis / 2, 0) '孔径直线
        Line2 = Sketchmer.CreateLine(L2Hub / 2, D2Axis / 2, 0, L2Hub / 2, D2Hub / 2, 0) '轮毂侧面
        Line3 = Sketchmer.CreateLine(L2Hub / 2, D2Hub / 2, 0, B2Spoke / 2, D2Hub / 2, 0) '轮毂外圆
        Line4 = Sketchmer.CreateLine(B2Spoke / 2, D2Hub / 2, 0, B2Spoke / 2, dm2 / 2, 0) '凹槽面
        Arc2 = Sketchmer.CreateCircleByRadius(0, a, 0, rf2 + k) '画大圆
        Line4.Select4(False, Nothing) '选中凹槽面直线
        Sketchmer.SketchTrim(0, B2Spoke / 2, dm2 / 2, 0) '切割凹槽面直线
        Line5 = Sketchmer.CreateLine(0, dm2 / 2, 0, b2 / 2, dm2 / 2, 0) '最大圆直线
        Line6 = Sketchmer.CreateLine(b2 / 2, dm2 / 2, 0, b2 / 2, dm2 / 4, 0) '齿宽侧面
        Sketchmer.SketchTrim(0, b2 / 2, dm2 / 4, 0) '修剪齿宽侧面
        Arc2.Select4(False, Nothing) '选中大圆
        Sketchmer.SketchTrim(0, 0, a + rf2 + k, 0) '修剪大圆
        Arc1 = Sketchmer.CreateCircleByRadius(0, a, 0, rg2) '画小圆
        Line5.Select4(False, Nothing) '选中最大圆直线
        Sketchmer.SketchTrim(0, 0, dm2 / 2, 0) '修剪最大圆直线
        Line0 = Sketchmer.CreateCenterLine(0, 0, 0, 0, a - rg2, 0) '画中心线
        Arc1.Select4(False, Nothing) '选中大圆
        Sketchmer.SketchTrim(0, 0, a + rg2, 0) '修剪小圆
        If RCast <> 0 Then
            Line3.Select4(False, Nothing)
            Line4.Select4(True, Nothing)
            Arc3 = Sketchmer.CreateFillet(RCast, 2) '倒圆角
            Arc2.Select4(False, Nothing)
            Line4.Select4(True, Nothing)
            Arc4 = Sketchmer.CreateFillet(RCast, 2) '倒圆角
        End If
        Line5.Select4(False, Nothing)
        Line6.Select4(True, Nothing)
        Line7 = Sketchmer.CreateChamfer(1, m / 1000, m / 1000) '倒角
        Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
        Line1.Select4(True, Nothing)
        Part.AddDimension2(0, D2Axis / 4, L2Hub / 2 + m / 1000) '标注轴孔半径
        Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
        Line3.Select4(True, Nothing)
        Part.AddDimension2(0, D2Hub / 4, L2Hub / 2 + 2 * m / 1000) '标注轮毂外径
        Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
        Line5.Select4(True, Nothing)
        Part.AddDimension2(0, D2Hub / 4, L2Hub / 2 + 3 * m / 1000) '标注最大半径
        Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
        Line2.Select4(True, Nothing)
        Part.AddDimension2(0, D2Axis / 2 - 2 * m / 1000, L2Hub / 4) '标注轮廓宽度的一半
        Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
        Line4.Select4(True, Nothing)
        Part.AddDimension2(0, D2Axis / 2 - m / 1000, L2Hub / 4) '标注凹槽宽度的一半
        Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
        Line6.Select4(True, Nothing)
        Part.AddDimension2(0, D2Axis / 2 - 3 * m / 1000, L2Hub / 4) '标注齿宽
        Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
        Arc1.Select4(True, Nothing)
        Part.AddDimension2(0, D2Hub / 4, L2Hub / 2 + 4 * m / 1000) '标注中心距
        Arc2.Select4(True, Nothing)
        Part.AddDimension2(0, dm2 / 2, 0) '给大圆标尺寸
        Arc1.Select4(True, Nothing)
        Part.AddDimension2(-b1 / 2, dm2 / 2, 0) '给小圆标尺寸
        Line1.SelectChain(False, Nothing)
        Line0.Select4(True, Nothing)
        Part.SketchMirror() '镜像
        Part.Extension.SelectByID2("基准轴1", "AXIS", 0, 0, 0, False, 16, Nothing, 0)
        Part.FeatureManager.FeatureRevolve2(True, True, False, False, False, False, 0, 0, 2 * PI, 0, False, False, 0.01, 0.01, 0, 0, 0, True, True, True) '旋转
        Part.ShowNamedView2("*等轴测", 7) '正等测视角
        Part.Extension.SelectByID2("", "FACE", 0, D2Spoke / 2, B2Spoke / 2, False, 0, Nothing, 0)
        Part.FeatureManager.InsertRefPlane(4, 0, 0, 0, 0, 0) '新建基准面1
        Part.Extension.SelectByID2("", "FACE", 0, D2Spoke / 2, -B2Spoke / 2, False, 0, Nothing, 0)
        Part.FeatureManager.InsertRefPlane(4, 0, 0, 0, 0, 0) '新建基准面2
        Part.Extension.SelectByID2("基准面1", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中基准面1
        Part.Extension.SelectByID2("基准面2", "PLANE", 0, 0, 0, True, 0, Nothing, 0) '选中基准面2
        Part.BlankRefGeom() '隐藏基准面12
        Part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 32, Nothing, 0) '选中前视基准面
        Sketchmer.InsertSketch(True) '新建草图2
        Sketchmer.CreateCircleByRadius(0, D2Spoke / 2, 0, D2SpokeHole / 2) '画圆
        Featmgr.FeatureCut3(False, False, False, 9, 1, 0.01, 0.01, False, False, False, False, 0, 0, False, False, False, False, False, True, True, True, True, False, 0, 0, False) '拉伸切除轮辐孔
        Part.Extension.SelectByID2("", "FACE", 0, D2Spoke / 2 - D2SpokeHole / 2, 0, False, 0, Nothing, 0)
        Part.InsertAxis2(True) '新建基准轴2
        Part.Extension.SelectByID2("基准轴2", "AXIS", 0, 0, 0, False, 0, Nothing, 0)
        Part.BlankRefGeom() '隐藏基准轴2
        Part.Extension.SelectByID2("切除-拉伸1", "BODYFEATURE", 0, 0, 0, False, 4, Nothing, 0)
        Part.Extension.SelectByID2("基准轴1", "AXIS", 0, 0, 0, True, 1, Nothing, 0)
        Part.FeatureManager.FeatureCircularPattern4(SpokeHoleNum, 2 * PI, False, "NULL", False, True, False) '圆周阵列
        Part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 32, Nothing, 0) '选中前视基准面
        Sketchmer.InsertSketch(True) '新建草图3
        Part.ShowNamedView2("*前视", 7) '前视角
        Part.ViewZoomtofit2() '整屏显示
        Dim t1#, t2#
        t1 = 0
        t2 = Tan(Acos(db2 / dm2))
        Arc1 = Sketchmer.CreateCircleByRadius(0, 0, 0, df2 / 2) '齿根圆
        Arc2 = Sketchmer.CreateCircleByRadius(0, 0, 0, dm2 / 2) '齿顶圆
        Line1 = Sketchmer.CreateEquationSpline2(db2 * 1000 / 2 & "*(sin(t)-t*cos(t))", db2 * 1000 / 2 & "*(cos(t)+t*sin(t))", "", t1, t2, False, 0, 0, 0, True, True) '绘制渐开线
        Line8 = Sketchmer.CreateCenterLine(0, 0, 0, 0.005, 0.005, 0) '绘制中心线
        Line8.Angle = PI / 2 + (Acos(db2 / d2) - Tan(Acos(db2 / d2)) + PI * 0.5 / z2) '中心线角度赋值
        If db2 > df2 Then
            Line2 = Sketchmer.CreateLine(0, db2 / 2, 0, 0, df2 / 2, 0) '绘制过渡曲线
            Line1.Select4(False, Nothing)
            Line2.Select4(True, Nothing)
            Line8.Select4(True, Nothing)
            Part.SketchMirror() '镜像
        Else
            Line1.Select4(False, Nothing)
            Sketchmer.SketchTrim(0, 0, db2 / 2, 0)
            Line1.Select4(False, Nothing)
            Line8.Select4(True, Nothing)
            Part.SketchMirror() '镜像
        End If
        Arc1.Select4(False, Nothing)
        Sketchmer.SketchTrim(0, 0, -df2 / 2, 0)
        Arc2.Select4(False, Nothing)
        Sketchmer.SketchTrim(0, 0, -dm2 / 2, 0)
        Line1.SelectChain(False, Nothing)
        Part.SketchAddConstraints("sgFIXED")
        Sketchmer.InsertSketch(True) '完成草图3
        Part.Extension.SelectByID2("Line1@草图3", "EXTSKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
        Part.Extension.SelectByID2("基准轴1", "AXIS", 0, 0, 0, True, 1, Nothing, 1)
        Featmgr.InsertRefPlane(4, 0, 4, 0, 0, 0) '新建基准面3
        Part.Extension.SelectByID2("基准面3", "PLANE", 0, 0, 0, False, 32, Nothing, 0) '选中基准面3
        Sketchmer.InsertSketch(True) '新建草图4
        Arc5 = Sketchmer.CreateCircleByRadius(0, a, 0, d1 / 2) '画圆
        Part.AddDimension2(0, a, 0) '给圆标直径
        Sketchmer.InsertSketch(True) '完成草图4
        Part.Extension.SelectByID2("草图3", "SKETCH", 0, 0, 0, False, 1, Nothing, 0) '选中扫描轮廓
        Part.Extension.SelectByID2("草图4", "SKETCH", 0, 0, 0, True, 4, Nothing, 0) '选中引导线
        Part.FeatureManager.InsertCutSwept4(False, True, 0, False, False, 0, 0, False, 0, 0, 0, 0, True, True, 0, True, True, True, False) '扫描切除
        Part.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, False, 2, Nothing, 0)
        Part.Extension.SelectByID2("切除-扫描1", "BODYFEATURE", 0, 0, 0, True, 1, Nothing, 0)
        Part.FeatureManager.InsertMirrorFeature(False, False, False, False)
        Part.Extension.SelectByID2("镜向1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        Part.EditSuppress2()
        Part.Extension.SelectByID2("切除-扫描1", "BODYFEATURE", 0, 0, 0, False, 4, Nothing, 0)
        Part.Extension.SelectByID2("基准轴1", "AXIS", 0, 0, 0, True, 1, Nothing, 0)
        Part.FeatureManager.FeatureCircularPattern3(z2, 2 * PI, False, "NULL", False, True)
        Part.Extension.SelectByID2("基准轴1", "AXIS", 0, 0, 0, False, 0, Nothing, 0)
        Part.Extension.SelectByID2("基准面3", "PLANE", 0, 0, 0, True, 0, Nothing, 0) '选中基准面3
        Part.BlankRefGeom() '隐藏基准轴1和基准面3
        Part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 32, Nothing, 0) '选中前视基准面
        Sketchmer.InsertSketch(True) '新建草图5
        Arc0 = Sketchmer.CreateCircleByRadius(0, 0, 0, D2Hub / 2) '画圆
        Featmgr.FeatureCut3(False, False, False, 1, 1, 0.01, 0.01, False, False, False, False, 0, 0, False, False, False, False, False, True, True, True, True, False, 0, 0, False) '拉伸切除
        Part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 32, Nothing, 0) '选中前视基准面
        Sketchmer.InsertSketch(True) '新建草图6
        Arc0 = Sketchmer.CreateCircleByRadius(0, 0, 0, D2Rim / 2) '画圆
        Featmgr.FeatureCut3(True, False, True, 1, 1, 0.01, 0.01, False, False, False, False, 0, 0, False, False, False, False, False, True, True, True, True, False, 0, 0, False) '拉伸切除
        Part.EditRebuild3() '重建模型
        Part.ClearSelection2（True）
        Part.ShowNamedView2("*等轴测", 7) '正等测视角
        Part.ViewZoomtofit2() '整屏显示
        Part.SaveAs3(BoltedTypeWormGearRimPath, 0, 2) '文件保存
        Part.Extension.SelectByID2("切除-拉伸3", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        Part.Extension.SelectByID2("切除-拉伸2", "BODYFEATURE", 0, 0, 0, True, 0, Nothing, 0)
        Part.EditDelete()
        Part.Extension.SelectByID2("草图6", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
        Part.Extension.SelectByID2("草图5", "SKETCH", 0, 0, 0, True, 0, Nothing, 0)
        Part.EditDelete()
        Part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 32, Nothing, 0) '选中前视基准面
        Sketchmer.InsertSketch(True) '新建草图7
        Arc0 = Sketchmer.CreateCircleByRadius(0, 0, 0, D2Rim / 2) '画圆
        Featmgr.FeatureCut3(False, True, True, 1, 1, 0.01, 0.01, False, False, False, False, 0, 0, False, False, False, False, False, True, True, True, True, False, 0, 0, False) '拉伸切除
        Part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 32, Nothing, 0) '选中前视基准面
        Sketchmer.InsertSketch(True) '新建草图8
        Arc0 = Sketchmer.CreateCircleByRadius(0, 0, 0, D2Hub / 2) '画圆
        Featmgr.FeatureCut3(True, True, False, 1, 1, 0.01, 0.01, False, False, False, False, 0, 0, False, False, False, False, False, True, True, True, True, False, 0, 0, False) '拉伸切除
        Part.EditRebuild3() '重建模型
        Part.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, False, 32, Nothing, 0) '选中上视基准面
        Sketchmer.InsertSketch(True) '新建草图9
        Part.ShowNamedView2("*上视", 5)
        Part.ViewZoomtofit2() '整屏显示
        Part.SketchManager.CreateCenterRectangle(0, 0, 0, Keywayb / 2, L2Hub, 0)
        Part.SketchAddConstraints("sgFIXED")
        Sketchmer.InsertSketch(True) '完成草图9
        Select Case ComboBoxKeyWayNum.Text
            Case 1
                Featmgr.FeatureCut3(True, False, True, 0, 0, D2Axis / 2 + Keywayh / 2, D2Axis / 2 + Keywayh / 2, False, False, False, False, 0, 0, False, False, False, False, False, True, True, True, True, False, 0, 0, False) '拉伸切除
            Case 2
                Featmgr.FeatureCut3(False, False, True, 0, 0, D2Axis / 2 + Keywayh / 2, D2Axis / 2 + Keywayh / 2, False, False, False, False, 0, 0, False, False, False, False, False, True, True, True, True, False, 0, 0, False) '拉伸切除
        End Select
        Part.ClearSelection2（True）
        Part.ShowNamedView2("*等轴测", 7) '正等测视角
        Part.ViewZoomtofit2() '整屏显示
        Part.SaveAs3(BoltedTypeWormGearHubPath, 0, 2) '文件保存
        Swapp.CloseAllDocuments(True) '文件关闭
        Swapp.NewDocument("C:\ProgramData\SolidWorks\SolidWorks " & SolidworksEdition & "\templates\gb_assembly.asmdot", 0, 0, 0)
        Assem = Swapp.ActiveDoc
        Part = Swapp.ActiveDoc
        Dim DocumentError&, DocumentWarn&, CoordinationError&
        Dim PartRim, PartHub, PartBolt, PartNut, PartWasher As SldWorks.ModelDoc2
        Dim BoltedTypeWormGearTitle$, PartRimTitle$, PartHubTitle$, PartBoltTitle$, PartNutTitle$, PartWasherTitle$
        BoltedTypeWormGearTitle = Assem.GetTitle() '获取装配体名称
        PartRim = Swapp.OpenDoc6(BoltedTypeWormGearRimPath, 1, 0, "", DocumentError, DocumentWarn) '打开轮缘
        PartRimTitle = PartRim.GetTitle() '获取轮缘文件名
        PartRimTitle = PartRimTitle.Substring(0, InStrRev(PartRimTitle, ".") - 1) '提取轮缘名称
        PartRim.Visible = False '隐藏轮缘文件
        Part.AddComponent4(BoltedTypeWormGearRimPath, "默认", 0, 0, 0) '添加轮缘
        PartHub = Swapp.OpenDoc6(BoltedTypeWormGearHubPath, 1, 0, "", DocumentError, DocumentWarn) '打开轮毂
        PartHubTitle = PartHub.GetTitle() '获取轮毂文件名
        PartHubTitle = PartHubTitle.Substring(0, InStrRev(PartHubTitle, ".") - 1) '提取轮毂名称
        PartHub.Visible = False '隐藏轮毂文件
        Part.AddComponent4(BoltedTypeWormGearHubPath, "默认", 0, 0, 0) '添加轮毂
        PartBolt = Swapp.OpenDoc6(BoltPath, 1, 0, "", DocumentError, DocumentWarn) '打开螺栓
        PartBoltTitle = PartBolt.GetTitle() '获取螺栓文件名
        PartBoltTitle = PartBoltTitle.Substring(0, InStrRev(PartBoltTitle, ".") - 1) '提取螺栓名称
        PartBolt.Visible = False '隐藏螺栓文件
        Part.AddComponent4(BoltPath, "默认", 0, 0, 0) '添加螺栓
        PartNut = Swapp.OpenDoc6(NutPath, 1, 0, "", DocumentError, DocumentWarn) '打开螺母
        PartNutTitle = PartNut.GetTitle() '获取螺母文件名
        PartNutTitle = PartNutTitle.Substring(0, InStrRev(PartNutTitle, ".") - 1) '提取螺母名称
        PartNut.Visible = False '隐藏螺母文件
        Part.AddComponent4(NutPath, "默认", 0, 0, 0) '添加螺母
        PartWasher = Swapp.OpenDoc6(WasherPath, 1, 0, "", DocumentError, DocumentWarn) '打开垫圈
        PartWasherTitle = PartWasher.GetTitle() '获取垫圈文件名
        PartWasherTitle = PartWasherTitle.Substring(0, InStrRev(PartWasherTitle, ".") - 1) '提取垫圈名称
        PartWasher.Visible = False '隐藏垫圈文件
        Part.AddComponent4(WasherPath, "默认", 0, 0, 0) '添加垫圈
        Part.ShowNamedView2("*等轴测", 7) '正等测视角
        Part.ViewZoomtofit2() '整屏显示
        Part.Extension.SelectByID2("前视基准面@" & PartHubTitle & "-1@" & BoltedTypeWormGearTitle, "PLANE", 0, 0, 0, False, 1, Nothing, 0) '选中轮毂前视基准面
        Part.Extension.SelectByID2("前视基准面@" & PartRimTitle & "-1@" & BoltedTypeWormGearTitle, "PLANE", 0, 0, 0, True, 1, Nothing, 0) '选中轮缘前视基准面
        Assem.AddMate5(0, -1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, CoordinationError) '重合配合
        Part.EditRebuild3() '重建模型
        Part.Extension.SelectByID2("前视基准面@" & PartBoltTitle & "-1@" & BoltedTypeWormGearTitle, "PLANE", 0, 0, 0, False, 1, Nothing, 0) '选中螺栓前视基准面
        Part.Extension.SelectByID2("基准面1@" & PartRimTitle & "-1@" & BoltedTypeWormGearTitle, "PLANE", 0, 0, 0, True, 1, Nothing, 0) '选中轮缘基准面1
        Assem.AddMate5(5, -1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, CoordinationError) '距离配合
        Part.EditRebuild3() '重建模型
        Part.Extension.SelectByID2("基准面1@" & PartWasherTitle & "-1@" & BoltedTypeWormGearTitle, "PLANE", 0, 0, 0, False, 1, Nothing, 0) '选中垫片基准面1
        Part.Extension.SelectByID2("基准面2@" & PartRimTitle & "-1@" & BoltedTypeWormGearTitle, "PLANE", 0, 0, 0, True, 1, Nothing, 0) '选中轮缘基准面2
        Assem.AddMate5(5, -1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, CoordinationError) '距离配合
        Part.EditRebuild3() '重建模型
        Part.Extension.SelectByID2("前视基准面@" & PartWasherTitle & "-1@" & BoltedTypeWormGearTitle, "PLANE", 0, 0, 0, False, 1, Nothing, 0) '选中垫片前视基准面
        Part.Extension.SelectByID2("基准面1@" & PartNutTitle & "-1@" & BoltedTypeWormGearTitle, "PLANE", 0, 0, 0, True, 1, Nothing, 0) '选中螺母基准面1
        Assem.AddMate5(5, -1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, CoordinationError) '距离配合
        Part.EditRebuild3() '重建模型
        Part.Extension.SelectByID2("基准轴1@" & PartRimTitle & "-1@" & BoltedTypeWormGearTitle, "AXIS", 0, 0, 0, False, 1, Nothing, 0)
        Part.Extension.SelectByID2("基准轴1@" & PartHubTitle & "-1@" & BoltedTypeWormGearTitle, "AXIS", 0, 0, 0, True, 1, Nothing, 0)
        Assem.AddMate5(5, -1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, CoordinationError) '距离配合
        Part.EditRebuild3() '重建模型
        Part.Extension.SelectByID2("基准轴2@" & PartRimTitle & "-1@" & BoltedTypeWormGearTitle, "AXIS", 0, 0, 0, False, 1, Nothing, 0)
        Part.Extension.SelectByID2("基准轴2@" & PartHubTitle & "-1@" & BoltedTypeWormGearTitle, "AXIS", 0, 0, 0, True, 1, Nothing, 0)
        Assem.AddMate5(5, -1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, CoordinationError) '距离配合
        Part.EditRebuild3() '重建模型
        Part.Extension.SelectByID2("基准轴2@" & PartRimTitle & "-1@" & BoltedTypeWormGearTitle, "AXIS", 0, 0, 0, False, 1, Nothing, 0)
        Part.Extension.SelectByID2("基准轴1@" & PartBoltTitle & "-1@" & BoltedTypeWormGearTitle, "AXIS", 0, 0, 0, True, 1, Nothing, 0)
        Assem.AddMate5(5, -1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, CoordinationError) '距离配合
        Part.EditRebuild3() '重建模型
        Part.Extension.SelectByID2("基准轴2@" & PartRimTitle & "-1@" & BoltedTypeWormGearTitle, "AXIS", 0, 0, 0, False, 1, Nothing, 0)
        Part.Extension.SelectByID2("基准轴1@" & PartWasherTitle & "-1@" & BoltedTypeWormGearTitle, "AXIS", 0, 0, 0, True, 1, Nothing, 0)
        Assem.AddMate5(5, -1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, CoordinationError) '距离配合
        Part.EditRebuild3() '重建模型
        Part.Extension.SelectByID2("基准轴2@" & PartRimTitle & "-1@" & BoltedTypeWormGearTitle, "AXIS", 0, 0, 0, False, 1, Nothing, 0)
        Part.Extension.SelectByID2("基准轴1@" & PartNutTitle & "-1@" & BoltedTypeWormGearTitle, "AXIS", 0, 0, 0, True, 1, Nothing, 0)
        Assem.AddMate5(5, -1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, CoordinationError) '距离配合
        Part.EditRebuild3() '重建模型
        Part.Extension.SelectByID2("右视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中右视基准面
        Part.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, True, 0, Nothing, 0) '选中上视基准面
        Part.InsertAxis2(True) '新建基准轴1
        Part.ShowNamedView2("*等轴测", 7) '正等测视角
        Part.ViewZoomtofit2() '整屏显示
        Part.Extension.SelectByID2(PartBoltTitle & "-1@" & BoltedTypeWormGearTitle, "COMPONENT", 0, 0, 0, False, 1, Nothing, 0)
        Part.Extension.SelectByID2(PartNutTitle & "-1@" & BoltedTypeWormGearTitle, "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
        Part.Extension.SelectByID2(PartWasherTitle & "-1@" & BoltedTypeWormGearTitle, "COMPONENT", 0, 0, 0, True, 1, Nothing, 0)
        Part.Extension.SelectByID2("基准轴1", "AXIS", 0, 0, 0, True, 2, Nothing, 0)
        Part.FeatureManager.FeatureCircularPattern3(SpokeHoleNum, 1.5707963267949, False, "NULL", False, False) '圆周阵列零件
        Part.Extension.SelectByID2("基准轴1", "AXIS", 0, 0, 0, False, 0, Nothing, 0)
        Part.BlankRefGeom() '隐藏基准轴1
        Part.ClearSelection2（True）
        Part.ShowNamedView2("*等轴测", 7) '正等测视角
        Part.ViewZoomtofit2() '整屏显示
        Part.SaveAs3(BoltedTypeWormGearPath, 0, 2)
        If CheckBoxCloseFile.Checked = True Then Swapp.CloseAllDocuments(True)
        Swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, Inputsize) '输入尺寸值还原
    End Sub
    '创建镶铸式蜗轮
    Public Sub CreateInlayTypeWormGear() '创建镶铸式蜗轮
        Dim InlayTypeWormGearRimPath$ = FilePath & "\蜗轮(镶铸式轮缘)" & z2 & "X" & m & ".SLDPRT"
        Dim InlayTypeWormGearHubPath$ = FilePath & "\蜗轮(镶铸式轮毂)" & z2 & "X" & m & ".SLDPRT"
        Dim InlayTypeWormGearPath$ = FilePath & "\蜗轮(镶铸式)" & z1 & "X" & z2 & "X" & m & ".SLDASM"
        AssemWormGearTitle = "蜗轮(镶铸式)" & z1 & "X" & z2 & "X" & m & ".SLDASM"
        AssemWormGearPath = InlayTypeWormGearPath
        Swapp = CreateObject("Sldworks.application")
        Swapp.Visible = True
        Swapp.NewDocument("C:\ProgramData\SolidWorks\SOLIDWORKS " & SolidworksEdition & "\templates\gb_part.prtdot", 0, 0, 0) '新建零件
        Part = Swapp.ActiveDoc
        Sketchmer = Part.SketchManager
        Featmgr = Part.FeatureManager
        Inputsize = Swapp.GetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate) '输入尺寸值记录
        Swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '输入尺寸值关闭
        Part.Extension.SelectByID2("右视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中右视基准面
        Part.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, True, 0, Nothing, 0) '选中上视基准面
        Part.InsertAxis2(True) '新建基准轴1
        Part.Extension.SelectByID2("右视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中右视基准面
        Dim Line0, Line1, Line2, Line3, Line4, Line5, Line6, Line7, Line8 As SldWorks.SketchSegment
        Dim Arc0, Arc1, Arc2, Arc3, Arc4, Arc5 As SldWorks.SketchSegment
        Sketchmer.InsertSketch(True) '新建草图1
        Line1 = Sketchmer.CreateLine(0, D2Axis / 2, 0, L2Hub / 2, D2Axis / 2, 0) '孔径直线
        Line2 = Sketchmer.CreateLine(L2Hub / 2, D2Axis / 2, 0, L2Hub / 2, D2Hub / 2, 0) '轮毂侧面
        Line3 = Sketchmer.CreateLine(L2Hub / 2, D2Hub / 2, 0, B2Spoke / 2, D2Hub / 2, 0) '轮毂外圆
        Line4 = Sketchmer.CreateLine(B2Spoke / 2, D2Hub / 2, 0, B2Spoke / 2, D2Rim / 2 - k, 0) '凹槽面
        Line5 = Sketchmer.CreateLine(B2Spoke / 2, D2Rim / 2 - k, 0, b2 / 2, D2Rim / 2 - k, 0) '轮缘内侧
        Line6 = Sketchmer.CreateLine(b2 / 2, D2Rim / 2 - k, 0, b2 / 2, dm2 / 2, 0) '齿宽侧面
        Line7 = Sketchmer.CreateLine(b2 / 2, dm2 / 2, 0, 0, dm2 / 2, 0) '齿顶面
        Arc1 = Sketchmer.CreateCircleByRadius(0, a, 0, rg2) '画小圆
        Line0 = Sketchmer.CreateCenterLine(0, 0, 0, 0, a - rg2, 0) '画中心线
        Arc1.Select4(False, Nothing) '选中小圆
        Sketchmer.SketchTrim(0, 0, a + rg2, 0) '修剪小圆
        Line7.Select4(False, Nothing) '选中齿顶面
        Sketchmer.SketchTrim(0, 0, dm2 / 2, 0) '修剪齿顶面
        If RCast <> 0 Then
            Line3.Select4(False, Nothing)
            Line4.Select4(True, Nothing)
            Arc3 = Sketchmer.CreateFillet(RCast, 2) '倒圆角
            Line4.Select4(False, Nothing)
            Line5.Select4(True, Nothing)
            Arc4 = Sketchmer.CreateFillet(RCast, 2) '倒圆角
        End If
        Line6.Select4(False, Nothing)
        Line7.Select4(True, Nothing)
        Line8 = Sketchmer.CreateChamfer(1, m / 1000, m / 1000) '倒角
        Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
        Line1.Select4(True, Nothing)
        Part.AddDimension2(0, D2Axis / 4, L2Hub / 2 + m / 1000) '标注轴孔半径
        Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
        Line3.Select4(True, Nothing)
        Part.AddDimension2(0, D2Hub / 4, L2Hub / 2 + 2 * m / 1000) '标注轮毂外径
        Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
        Line5.Select4(True, Nothing)
        Part.AddDimension2(0, D2Hub / 4, L2Hub / 2 + 3 * m / 1000) '标注轮缘半径
        Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
        Line7.Select4(True, Nothing)
        Part.AddDimension2(0, D2Hub / 4, L2Hub / 2 + 4 * m / 1000) '标注最大半径
        Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
        Line2.Select4(True, Nothing)
        Part.AddDimension2(0, D2Axis / 2 - 2 * m / 1000, L2Hub / 4) '标注轮廓宽度的一半
        Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
        Line4.Select4(True, Nothing)
        Part.AddDimension2(0, D2Axis / 2 - m / 1000, L2Hub / 4) '标注凹槽宽度的一半
        Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
        Line6.Select4(True, Nothing)
        Part.AddDimension2(0, D2Axis / 2 - 3 * m / 1000, L2Hub / 4) '标注齿宽
        Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
        Arc1.Select4(True, Nothing)
        Part.AddDimension2(0, D2Hub / 4, L2Hub / 2 + 4 * m / 1000) '标注中心距
        Arc1.Select4(True, Nothing)
        Part.AddDimension2(-b1 / 2, dm2 / 2, 0) '给小圆标尺寸
        Line1.SelectChain(False, Nothing)
        Line0.Select4(True, Nothing)
        Part.SketchMirror() '镜像
        Part.Extension.SelectByID2("基准轴1", "AXIS", 0, 0, 0, False, 16, Nothing, 0)
        Part.FeatureManager.FeatureRevolve2(True, True, False, False, False, False, 0, 0, 2 * PI, 0, False, False, 0.01, 0.01, 0, 0, 0, True, True, True) '旋转
        Part.ShowNamedView2("*等轴测", 7) '正等测视角
        Part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 32, Nothing, 0) '选中前视基准面
        Sketchmer.InsertSketch(True) '新建草图2
        Sketchmer.CreateCircleByRadius(0, D2Spoke / 2, 0, D2SpokeHole / 2) '画圆
        Featmgr.FeatureCut3(False, False, False, 9, 1, 0.01, 0.01, False, False, False, False, 0, 0, False, False, False, False, False, True, True, True, True, False, 0, 0, False) '拉伸切除轮辐孔
        Part.Extension.SelectByID2("切除-拉伸1", "BODYFEATURE", 0, 0, 0, False, 4, Nothing, 0)
        Part.Extension.SelectByID2("基准轴1", "AXIS", 0, 0, 0, True, 1, Nothing, 0)
        Part.FeatureManager.FeatureCircularPattern4(SpokeHoleNum, 2 * PI, False, "NULL", False, True, False) '圆周阵列
        Part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 32, Nothing, 0) '选中前视基准面
        Sketchmer.InsertSketch(True) '新建草图3
        Part.ShowNamedView2("*前视", 7) '前视角
        Part.ViewZoomtofit2() '整屏显示
        Dim t1#, t2#
        t1 = 0
        t2 = Tan(Acos(db2 / dm2))
        Arc1 = Sketchmer.CreateCircleByRadius(0, 0, 0, df2 / 2) '齿根圆
        Arc2 = Sketchmer.CreateCircleByRadius(0, 0, 0, dm2 / 2) '齿顶圆
        Line1 = Sketchmer.CreateEquationSpline2(db2 * 1000 / 2 & "*(sin(t)-t*cos(t))", db2 * 1000 / 2 & "*(cos(t)+t*sin(t))", "", t1, t2, False, 0, 0, 0, True, True) '绘制渐开线
        Line8 = Sketchmer.CreateCenterLine(0, 0, 0, 0.005, 0.005, 0) '绘制中心线
        Line8.Angle = PI / 2 + (Acos(db2 / d2) - Tan(Acos(db2 / d2)) + PI * 0.5 / z2) '中心线角度赋值
        If db2 > df2 Then
            Line2 = Sketchmer.CreateLine(0, db2 / 2, 0, 0, df2 / 2, 0) '绘制过渡曲线
            Line1.Select4(False, Nothing)
            Line2.Select4(True, Nothing)
            Line8.Select4(True, Nothing)
            Part.SketchMirror() '镜像
        Else
            Line1.Select4(False, Nothing)
            Sketchmer.SketchTrim(0, 0, db2 / 2, 0)
            Line1.Select4(False, Nothing)
            Line8.Select4(True, Nothing)
            Part.SketchMirror() '镜像
        End If
        Arc1.Select4(False, Nothing)
        Sketchmer.SketchTrim(0, 0, -df2 / 2, 0)
        Arc2.Select4(False, Nothing)
        Sketchmer.SketchTrim(0, 0, -dm2 / 2, 0)
        Line1.SelectChain(False, Nothing)
        Part.SketchAddConstraints("sgFIXED")
        Sketchmer.InsertSketch(True) '完成草图3
        Part.Extension.SelectByID2("Line1@草图3", "EXTSKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
        Part.Extension.SelectByID2("基准轴1", "AXIS", 0, 0, 0, True, 1, Nothing, 1)
        Featmgr.InsertRefPlane(4, 0, 4, 0, 0, 0) '新建基准面1
        Part.Extension.SelectByID2("基准面1", "PLANE", 0, 0, 0, False, 32, Nothing, 0) '选中基准面1
        Sketchmer.InsertSketch(True) '新建草图4
        Arc5 = Sketchmer.CreateCircleByRadius(0, a, 0, d1 / 2) '画圆
        Part.AddDimension2(0, a, 0) '给圆标直径
        Sketchmer.InsertSketch(True) '完成草图4
        Part.Extension.SelectByID2("草图3", "SKETCH", 0, 0, 0, False, 1, Nothing, 0) '选中扫描轮廓
        Part.Extension.SelectByID2("草图4", "SKETCH", 0, 0, 0, True, 4, Nothing, 0) '选中引导线
        Part.FeatureManager.InsertCutSwept4(False, True, 0, False, False, 0, 0, False, 0, 0, 0, 0, True, True, 0, True, True, True, False) '扫描切除
        Part.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, False, 2, Nothing, 0)
        Part.Extension.SelectByID2("切除-扫描1", "BODYFEATURE", 0, 0, 0, True, 1, Nothing, 0)
        Part.FeatureManager.InsertMirrorFeature(False, False, False, False) '镜像
        Part.Extension.SelectByID2("镜向1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0) '选中镜像
        Part.EditSuppress2() '压缩镜像
        Part.Extension.SelectByID2("切除-扫描1", "BODYFEATURE", 0, 0, 0, False, 4, Nothing, 0)
        Part.Extension.SelectByID2("基准轴1", "AXIS", 0, 0, 0, True, 1, Nothing, 0)
        Part.FeatureManager.FeatureCircularPattern3(z2, 2 * PI, False, "NULL", False, True)
        Part.Extension.SelectByID2("基准轴1", "AXIS", 0, 0, 0, False, 0, Nothing, 0)
        Part.Extension.SelectByID2("基准面1", "PLANE", 0, 0, 0, True, 0, Nothing, 0) '选中基准面1
        Part.BlankRefGeom() '隐藏基准轴1和基准面1
        Part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 32, Nothing, 0) '选中前视基准面
        Sketchmer.InsertSketch(True) '新建草图5
        Sketchmer.CreateCircleByRadius(0, 0, 0, D2Rim / 2) '画圆
        Featmgr.FeatureCut3(False, False, False, 1, 1, 0.01, 0.01, False, False, False, False, 0, 0, False, False, False, False, False, True, True, True, True, False, 0, 0, False) '拉伸切除
        Part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 32, Nothing, 0) '选中前视基准面
        Sketchmer.InsertSketch(True) '新建草图6
        Arc0 = Sketchmer.CreateCircleByRadius(0, 0, 0, D2Rim / 2) '画大圆
        Arc1 = Sketchmer.CreateCircleByRadius(0, D2Rim / 2, 0, m / 1000) '画小圆
        Arc0.Select4(False, Nothing) '选大圆
        Sketchmer.SketchTrim(0, 0, -D2Rim / 2, 0) '修剪大圆
        Arc1.Select4(False, Nothing) '选小圆
        Sketchmer.SketchTrim(0, 0, D2Rim / 2 + m / 1000, 0) '修剪小圆
        Part.FeatureManager.FeatureExtrusion2(True, False, False, 6, 0, b2 / 4, b2 / 4, False, False, False, False, 0, 0, False, False, False, False, True, True, True, 0, 0, False) '拉伸
        Part.ShowNamedView2("*等轴测", 7) '正等测视角
        Part.ViewZoomtofit2() '整屏显示
        Part.Extension.SelectByID2("", "FACE", 0, D2Rim / 2 - m / 1000, 0, False, 0, Nothing, 0)
        Part.InsertAxis2(True) '基准轴2
        Part.Extension.SelectByID2("基准轴2", "AXIS", 0, 0, 0, False, 0, Nothing, 0)
        Part.BlankRefGeom() '隐藏基准轴2
        Part.Extension.SelectByID2("凸台-拉伸1", "BODYFEATURE", 0, 0, 0, False, 4, Nothing, 0)
        Part.Extension.SelectByID2("基准轴1", "AXIS", 0, 0, 0, True, 1, Nothing, 0)
        Part.FeatureManager.FeatureCircularPattern4(InlayTypeWormGearKeyNum, 2 * PI, False, "NULL", False, True, False) '圆周阵列
        Part.EditRebuild3() '重建模型
        Part.ClearSelection2（True）
        Part.ShowNamedView2("*等轴测", 7) '正等测视角
        Part.ViewZoomtofit2() '整屏显示
        Part.SaveAs3(InlayTypeWormGearRimPath, 0, 2) '文件保存
        Part.Extension.SelectByID2("阵列(圆周)3", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
        Part.Extension.SelectByID2("凸台-拉伸1", "BODYFEATURE", 0, 0, 0, True, 0, Nothing, 0)
        Part.Extension.SelectByID2("切除-拉伸2", "BODYFEATURE", 0, 0, 0, True, 0, Nothing, 0)
        Part.EditDelete()
        Part.Extension.SelectByID2("草图6", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
        Part.Extension.SelectByID2("草图5", "SKETCH", 0, 0, 0, True, 0, Nothing, 0)
        Part.EditDelete()
        Part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 32, Nothing, 0) '选中前视基准面
        Sketchmer.InsertSketch(True) '新建草图5
        Sketchmer.CreateCircleByRadius(0, 0, 0, D2Rim / 2) '画圆
        Featmgr.FeatureCut3(False, True, False, 1, 1, 0.01, 0.01, False, False, False, False, 0, 0, False, False, False, False, False, True, True, True, True, False, 0, 0, False) '拉伸切除
        Part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 32, Nothing, 0) '选中前视基准面
        Sketchmer.InsertSketch(True) '新建草图6
        Sketchmer.CreateCircleByRadius(0, D2Rim / 2, 0, m / 1000) '画圆
        Featmgr.FeatureCut3(True, False, False, 6, 0, b2 / 4, b2 / 4, False, False, False, False, 0, 0, False, False, False, False, False, True, True, True, True, False, 0, 0, False) '拉伸切除
        Part.ShowNamedView2("*等轴测", 7) '正等测视角
        Part.ViewZoomtofit2() '整屏显示
        Part.Extension.SelectByID2("", "FACE", 0, D2Rim / 2 - m / 1000, 0, False, 0, Nothing, 0)
        Part.InsertAxis2(True) '基准轴3
        Part.Extension.SelectByID2("基准轴2", "AXIS", 0, 0, 0, False, 0, Nothing, 0)
        Part.BlankRefGeom() '隐藏基准轴3
        Part.Extension.SelectByID2("切除-拉伸4", "BODYFEATURE", 0, 0, 0, False, 4, Nothing, 0)
        Part.Extension.SelectByID2("基准轴1", "AXIS", 0, 0, 0, True, 1, Nothing, 0)
        Part.FeatureManager.FeatureCircularPattern4(InlayTypeWormGearKeyNum, 2 * PI, False, "NULL", False, True, False) '圆周阵列
        Part.EditRebuild3() '重建模型
        Part.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, False, 32, Nothing, 0) '选中上视基准面
        Sketchmer.InsertSketch(True) '新建草图7
        Part.ShowNamedView2("*上视", 5)
        Part.ViewZoomtofit2() '整屏显示
        Part.SketchManager.CreateCenterRectangle(0, 0, 0, Keywayb / 2, L2Hub, 0)
        Part.SketchAddConstraints("sgFIXED")
        Sketchmer.InsertSketch(True) '完成草图7
        Select Case ComboBoxKeyWayNum.Text
            Case 1
                Featmgr.FeatureCut3(True, False, True, 0, 0, D2Axis / 2 + Keywayh / 2, D2Axis / 2 + Keywayh / 2, False, False, False, False, 0, 0, False, False, False, False, False, True, True, True, True, False, 0, 0, False) '拉伸切除
            Case 2
                Featmgr.FeatureCut3(False, False, True, 0, 0, D2Axis / 2 + Keywayh / 2, D2Axis / 2 + Keywayh / 2, False, False, False, False, 0, 0, False, False, False, False, False, True, True, True, True, False, 0, 0, False) '拉伸切除
        End Select
        Part.ClearSelection2（True）
        Part.ShowNamedView2("*等轴测", 7) '正等测视角
        Part.ViewZoomtofit2() '整屏显示
        Part.SaveAs3(InlayTypeWormGearHubPath, 0, 2) '文件保存
        Swapp.CloseAllDocuments(True) '文件关闭
        Swapp.NewDocument("C:\ProgramData\SolidWorks\SolidWorks " & SolidworksEdition & "\templates\gb_assembly.asmdot", 0, 0, 0)
        Assem = Swapp.ActiveDoc
        Part = Swapp.ActiveDoc
        Dim DocumentError&, DocumentWarn&, CoordinationError&
        Dim PartRim, PartHub As SldWorks.ModelDoc2
        Dim InlayTypeWormGearTitle$, PartRimTitle$, PartHubTitle$
        InlayTypeWormGearTitle = Assem.GetTitle() '获取装配体名称
        PartRim = Swapp.OpenDoc6(InlayTypeWormGearRimPath, 1, 0, "", DocumentError, DocumentWarn) '打开轮缘
        PartRimTitle = PartRim.GetTitle() '获取轮缘文件名
        PartRimTitle = PartRimTitle.Substring(0, InStrRev(PartRimTitle, ".") - 1) '提取轮缘名称
        PartRim.Visible = False '隐藏轮缘文件
        Part.AddComponent4(InlayTypeWormGearRimPath, "默认", 0, 0, 0) '添加轮缘
        PartHub = Swapp.OpenDoc6(InlayTypeWormGearHubPath, 1, 0, "", DocumentError, DocumentWarn) '打开轮毂
        PartHubTitle = PartHub.GetTitle() '获取轮毂文件名
        PartHubTitle = PartHubTitle.Substring(0, InStrRev(PartHubTitle, ".") - 1) '提取轮毂名称
        PartHub.Visible = False '隐藏轮毂文件
        Part.AddComponent4(InlayTypeWormGearHubPath, "默认", 0, 0, 0) '添加轮毂
        Part.ShowNamedView2("*等轴测", 7) '正等测视角
        Part.ViewZoomtofit2() '整屏显示
        Part.Extension.SelectByID2("前视基准面@" & PartHubTitle & "-1@" & InlayTypeWormGearTitle, "PLANE", 0, 0, 0, False, 1, Nothing, 0) '选中轮毂前视基准面
        Part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, True, 1, Nothing, 0) '选中前视基准面
        Assem.AddMate5(0, -1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, CoordinationError) '重合配合
        Part.EditRebuild3() '重建模型
        Part.Extension.SelectByID2("基准轴1@" & PartRimTitle & "-1@" & InlayTypeWormGearTitle, "AXIS", 0, 0, 0, False, 1, Nothing, 0)
        Part.Extension.SelectByID2("基准轴1@" & PartHubTitle & "-1@" & InlayTypeWormGearTitle, "AXIS", 0, 0, 0, True, 1, Nothing, 0)
        Assem.AddMate5(5, -1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, CoordinationError) '距离配合
        Part.EditRebuild3() '重建模型
        Part.Extension.SelectByID2("基准轴2@" & PartRimTitle & "-1@" & InlayTypeWormGearTitle, "AXIS", 0, 0, 0, False, 1, Nothing, 0)
        Part.Extension.SelectByID2("基准轴3@" & PartHubTitle & "-1@" & InlayTypeWormGearTitle, "AXIS", 0, 0, 0, True, 1, Nothing, 0)
        Assem.AddMate5(5, -1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, CoordinationError) '距离配合
        Part.EditRebuild3() '重建模型
        Part.ClearSelection2（True）
        Part.ShowNamedView2("*等轴测", 7) '正等测视角
        Part.ViewZoomtofit2() '整屏显示
        Part.SaveAs3(InlayTypeWormGearPath, 0, 2)
        If CheckBoxCloseFile.Checked = True Then Swapp.CloseAllDocuments(True)
        Swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, Inputsize) '输入尺寸值还原
    End Sub
    '创建轮箍式蜗轮
    Public Sub CreateTyreTypeWormGear() '创建轮箍式蜗轮
        If TextBoxNutBoltNum.Enabled = True Then
            Dim TyreTypeWormGearRimPath$ = FilePath & "\蜗轮(轮箍式轮缘带螺钉)" & z2 & "X" & m & ".SLDPRT"
            Dim TyreTypeWormGearHubPath$ = FilePath & "\蜗轮(轮箍式轮毂带螺钉)" & z2 & "X" & m & ".SLDPRT"
            Dim TyreTypeWormGearPath$ = FilePath & "\蜗轮(轮箍式带螺钉)" & z1 & "X" & z2 & "X" & m & ".SLDASM"
            Dim NutBoltPath$ = FilePath & "\螺钉M" & TextBoxNutBoltNum.Text & ".SLDPRT"
            AssemWormGearTitle = "蜗轮(轮箍式带螺钉)" & z1 & "X" & z2 & "X" & m & ".SLDASM"
            AssemWormGearPath = TyreTypeWormGearPath
            Swapp = CreateObject("Sldworks.application")
            Swapp.Visible = True
            Swapp.NewDocument("C:\ProgramData\SolidWorks\SOLIDWORKS " & SolidworksEdition & "\templates\gb_part.prtdot", 0, 0, 0) '新建零件
            Part = Swapp.ActiveDoc
            Sketchmer = Part.SketchManager
            Featmgr = Part.FeatureManager
            Inputsize = Swapp.GetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate) '输入尺寸值记录
            Swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '输入尺寸值关闭
            Part.Extension.SelectByID2("右视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中右视基准面
            Part.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, True, 0, Nothing, 0) '选中上视基准面
            Part.InsertAxis2(True) '新建基准轴1
            Part.Extension.SelectByID2("基准轴1", "AXIS", 0, 0, 0, False, 0, Nothing, 0)
            Part.BlankRefGeom() '隐藏基准轴1
            Part.Extension.SelectByID2("右视基准面", "PLANE", 0, 0, 0, False, 32, Nothing, 0) '选中右视基准面
            Sketchmer.InsertSketch(True) '新建草图1
            Part.SketchManager.CreateCenterLine(0#, 0#, 0#, -NutBoltL, 0#, 0#) '中心线
            Part.ViewZoomtofit2() '整屏显示
            Part.SketchManager.CreateCornerRectangle(0#, 0#, 0#, NutBoltL, NutBoltD / 2, 0#) '矩形
            Part.FeatureManager.FeatureRevolve2(True, True, False, False, False, False, 0, 0, 2 * PI, 0, False, False, 0.01, 0.01, 0, 0, 0, True, True, True) '旋转
            Part.ShowNamedView2("*等轴测", 7) '正等测视角
            Part.ViewZoomtofit2() '整屏显示
            Part.Extension.SelectByID2("", "EDGE", NutBoltD / 2, 0, -NutBoltL, False, 0, Nothing, 0) '选择边
            Part.FeatureManager.InsertFeatureChamfer(4, 1, (NutBoltD - NutBoltdt) / 2, PI / 4, 0, 0, 0, 0) '倒角
            Part.Extension.SelectByID2("", "EDGE", NutBoltD / 2, 0, 0, False, 0, Nothing, 0) '选择边
            Part.FeatureManager.InsertFeatureChamfer(4, 1, NutBoltc, PI / 4, 0, 0, 0, 0) '倒角
            Part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 32, Nothing, 0)
            Sketchmer.InsertSketch(True) '新建草图2
            Part.SketchManager.CreateCenterRectangle(0, 0, 0, NutBoltD, NutBoltn / 2, 0) '画矩形
            Part.FeatureManager.FeatureCut3(True, False, False, 0, 0, NutBoltt, NutBoltt, False, False, False, False, 0, 0, False, False, False, False, False, True, True, True, True, False, 0, 0, False) '拉伸切除
            Part.EditRebuild3() '重建模型
            Part.ClearSelection2（True）
            Part.ShowNamedView2("*等轴测", 7) '正等测视角
            Part.ViewZoomtofit2() '整屏显示
            Part.SaveAs3(NutBoltPath, 0, 2) '文件保存
            Swapp.CloseAllDocuments(True) '文件关闭
            Swapp.NewDocument("C:\ProgramData\SolidWorks\SOLIDWORKS " & SolidworksEdition & "\templates\gb_part.prtdot", 0, 0, 0) '新建零件
            Part = Swapp.ActiveDoc
            Sketchmer = Part.SketchManager
            Featmgr = Part.FeatureManager
            Part.Extension.SelectByID2("右视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中右视基准面
            Part.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, True, 0, Nothing, 0) '选中上视基准面
            Part.InsertAxis2(True) '新建基准轴1
            Part.Extension.SelectByID2("右视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中右视基准面
            Dim Line0, Line1, Line2, Line3, Line4, Line5, Line6, Line7, Line8 As SldWorks.SketchSegment
            Dim Arc1, Arc2, Arc3, Arc4, Arc5 As SldWorks.SketchSegment
            Sketchmer.InsertSketch(True) '新建草图1
            Line1 = Sketchmer.CreateLine(0, D2Axis / 2, 0, L2Hub / 2, D2Axis / 2, 0) '孔径直线
            Line2 = Sketchmer.CreateLine(L2Hub / 2, D2Axis / 2, 0, L2Hub / 2, D2Hub / 2, 0) '轮毂侧面
            Line3 = Sketchmer.CreateLine(L2Hub / 2, D2Hub / 2, 0, B2Spoke / 2, D2Hub / 2, 0) '轮毂外圆
            Line4 = Sketchmer.CreateLine(B2Spoke / 2, D2Hub / 2, 0, B2Spoke / 2, D2Rim / 2 - k, 0) '凹槽面
            Line5 = Sketchmer.CreateLine(B2Spoke / 2, D2Rim / 2 - k, 0, b2 / 2, D2Rim / 2 - k, 0) '轮缘内侧
            Line6 = Sketchmer.CreateLine(b2 / 2, D2Rim / 2 - k, 0, b2 / 2, dm2 / 2, 0) '齿宽侧面
            Line7 = Sketchmer.CreateLine(b2 / 2, dm2 / 2, 0, 0, dm2 / 2, 0) '齿顶面
            Arc1 = Sketchmer.CreateCircleByRadius(0, a, 0, rg2) '画小圆
            Line0 = Sketchmer.CreateCenterLine(0, 0, 0, 0, a - rg2, 0) '画中心线
            Arc1.Select4(False, Nothing) '选中小圆
            Sketchmer.SketchTrim(0, 0, a + rg2, 0) '修剪小圆
            Line7.Select4(False, Nothing) '选中齿顶面
            Sketchmer.SketchTrim(0, 0, dm2 / 2, 0) '修剪齿顶面
            If RCast <> 0 Then
                Line3.Select4(False, Nothing)
                Line4.Select4(True, Nothing)
                Arc3 = Sketchmer.CreateFillet(RCast, 2) '倒圆角
                Line4.Select4(False, Nothing)
                Line5.Select4(True, Nothing)
                Arc4 = Sketchmer.CreateFillet(RCast, 2) '倒圆角
            End If
            Line6.Select4(False, Nothing)
            Line7.Select4(True, Nothing)
            Line8 = Sketchmer.CreateChamfer(1, m / 1000, m / 1000) '倒角
            Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
            Line1.Select4(True, Nothing)
            Part.AddDimension2(0, D2Axis / 4, L2Hub / 2 + m / 1000) '标注轴孔半径
            Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
            Line3.Select4(True, Nothing)
            Part.AddDimension2(0, D2Hub / 4, L2Hub / 2 + 2 * m / 1000) '标注轮毂外径
            Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
            Line5.Select4(True, Nothing)
            Part.AddDimension2(0, D2Hub / 4, L2Hub / 2 + 3 * m / 1000) '标注轮缘半径
            Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
            Line7.Select4(True, Nothing)
            Part.AddDimension2(0, D2Hub / 4, L2Hub / 2 + 4 * m / 1000) '标注最大半径
            Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
            Line2.Select4(True, Nothing)
            Part.AddDimension2(0, D2Axis / 2 - 2 * m / 1000, L2Hub / 4) '标注轮廓宽度的一半
            Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
            Line4.Select4(True, Nothing)
            Part.AddDimension2(0, D2Axis / 2 - m / 1000, L2Hub / 4) '标注凹槽宽度的一半
            Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
            Line6.Select4(True, Nothing)
            Part.AddDimension2(0, D2Axis / 2 - 3 * m / 1000, L2Hub / 4) '标注齿宽
            Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
            Arc1.Select4(True, Nothing)
            Part.AddDimension2(0, D2Hub / 4, L2Hub / 2 + 4 * m / 1000) '标注中心距
            Arc1.Select4(True, Nothing)
            Part.AddDimension2(-b1 / 2, dm2 / 2, 0) '给小圆标尺寸
            Line1.SelectChain(False, Nothing)
            Line0.Select4(True, Nothing)
            Part.SketchMirror() '镜像
            Part.Extension.SelectByID2("基准轴1", "AXIS", 0, 0, 0, False, 16, Nothing, 0)
            Part.FeatureManager.FeatureRevolve2(True, True, False, False, False, False, 0, 0, 2 * PI, 0, False, False, 0.01, 0.01, 0, 0, 0, True, True, True) '旋转
            Part.ShowNamedView2("*等轴测", 7) '正等测视角
            Part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 32, Nothing, 0) '选中前视基准面
            Sketchmer.InsertSketch(True) '新建草图2
            Sketchmer.CreateCircleByRadius(0, D2Spoke / 2, 0, D2SpokeHole / 2) '画圆
            Featmgr.FeatureCut3(False, False, False, 9, 1, 0.01, 0.01, False, False, False, False, 0, 0, False, False, False, False, False, True, True, True, True, False, 0, 0, False) '拉伸切除轮辐孔
            Part.Extension.SelectByID2("切除-拉伸1", "BODYFEATURE", 0, 0, 0, False, 4, Nothing, 0)
            Part.Extension.SelectByID2("基准轴1", "AXIS", 0, 0, 0, True, 1, Nothing, 0)
            Part.FeatureManager.FeatureCircularPattern4(SpokeHoleNum, 2 * PI, False, "NULL", False, True, False) '圆周阵列
            Part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 32, Nothing, 0) '选中前视基准面
            Sketchmer.InsertSketch(True) '新建草图3
            Part.ShowNamedView2("*前视", 7) '前视角
            Part.ViewZoomtofit2() '整屏显示
            Dim t1#, t2#
            t1 = 0
            t2 = Tan(Acos(db2 / dm2))
            Arc1 = Sketchmer.CreateCircleByRadius(0, 0, 0, df2 / 2) '齿根圆
            Arc2 = Sketchmer.CreateCircleByRadius(0, 0, 0, dm2 / 2) '齿顶圆
            Line1 = Sketchmer.CreateEquationSpline2(db2 * 1000 / 2 & "*(sin(t)-t*cos(t))", db2 * 1000 / 2 & "*(cos(t)+t*sin(t))", "", t1, t2, False, 0, 0, 0, True, True) '绘制渐开线
            Line8 = Sketchmer.CreateCenterLine(0, 0, 0, 0.005, 0.005, 0) '绘制中心线
            Line8.Angle = PI / 2 + (Acos(db2 / d2) - Tan(Acos(db2 / d2)) + PI * 0.5 / z2) '中心线角度赋值
            If db2 > df2 Then
                Line2 = Sketchmer.CreateLine(0, db2 / 2, 0, 0, df2 / 2, 0) '绘制过渡曲线
                Line1.Select4(False, Nothing)
                Line2.Select4(True, Nothing)
                Line8.Select4(True, Nothing)
                Part.SketchMirror() '镜像
            Else
                Line1.Select4(False, Nothing)
                Sketchmer.SketchTrim(0, 0, db2 / 2, 0)
                Line1.Select4(False, Nothing)
                Line8.Select4(True, Nothing)
                Part.SketchMirror() '镜像
            End If
            Arc1.Select4(False, Nothing)
            Sketchmer.SketchTrim(0, 0, -df2 / 2, 0)
            Arc2.Select4(False, Nothing)
            Sketchmer.SketchTrim(0, 0, -dm2 / 2, 0)
            Line1.SelectChain(False, Nothing)
            Part.SketchAddConstraints("sgFIXED")
            Sketchmer.InsertSketch(True) '完成草图3
            Part.Extension.SelectByID2("Line1@草图3", "EXTSKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
            Part.Extension.SelectByID2("基准轴1", "AXIS", 0, 0, 0, True, 1, Nothing, 1)
            Featmgr.InsertRefPlane(4, 0, 4, 0, 0, 0) '新建基准面1
            Part.Extension.SelectByID2("基准面1", "PLANE", 0, 0, 0, False, 32, Nothing, 0) '选中基准面1
            Sketchmer.InsertSketch(True) '新建草图4
            Arc5 = Sketchmer.CreateCircleByRadius(0, a, 0, d1 / 2) '画圆
            Part.AddDimension2(0, a, 0) '给圆标直径
            Sketchmer.InsertSketch(True) '完成草图4
            Part.Extension.SelectByID2("草图3", "SKETCH", 0, 0, 0, False, 1, Nothing, 0) '选中扫描轮廓
            Part.Extension.SelectByID2("草图4", "SKETCH", 0, 0, 0, True, 4, Nothing, 0) '选中引导线
            Part.FeatureManager.InsertCutSwept4(False, True, 0, False, False, 0, 0, False, 0, 0, 0, 0, True, True, 0, True, True, True, False) '扫描切除
            Part.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, False, 2, Nothing, 0)
            Part.Extension.SelectByID2("切除-扫描1", "BODYFEATURE", 0, 0, 0, True, 1, Nothing, 0)
            Part.FeatureManager.InsertMirrorFeature(False, False, False, False) '镜像
            Part.Extension.SelectByID2("镜向1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0) '选中镜像
            Part.EditSuppress2() '压缩镜像
            Part.Extension.SelectByID2("切除-扫描1", "BODYFEATURE", 0, 0, 0, False, 4, Nothing, 0)
            Part.Extension.SelectByID2("基准轴1", "AXIS", 0, 0, 0, True, 1, Nothing, 0)
            Part.FeatureManager.FeatureCircularPattern3(z2, 2 * PI, False, "NULL", False, True) '旋转齿
            Part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
            Part.FeatureManager.InsertRefPlane(8, b2 / 2, 0, 0, 0, 0) '基准面2
            Part.Extension.SelectByID2("基准轴1", "AXIS", 0, 0, 0, False, 0, Nothing, 0)
            Part.Extension.SelectByID2("基准面1", "PLANE", 0, 0, 0, True, 0, Nothing, 0) '选中基准面1
            Part.Extension.SelectByID2("基准面2", "PLANE", 0, 0, 0, True, 0, Nothing, 0) '选中基准面2
            Part.BlankRefGeom() '隐藏基准轴1和基准面12
            Part.Extension.SelectByID2("", "FACE", 0, D2Rim / 2, b2 / 2, False, 0, Nothing, 0)
            Sketchmer.InsertSketch(True) '新建草图5
            Sketchmer.CreateCircleByRadius(0, D2Rim / 2, 0, NutBoltD / 2) '画圆
            Part.FeatureManager.FeatureCut3(True, False, False, 0, 0, NutBoltHoleL, NutBoltHoleL, False, False, False, False, 0, 0, False, False, False, False, False, True, True, True, True, False, 0, 0, False) '拉伸切除
            Part.ShowNamedView2("*等轴测", 7) '正等测视角
            Part.ViewZoomtofit2() '整屏显示
            Part.Extension.SelectByID2("", "FACE", 0, D2Rim / 2 + NutBoltD / 2, b2 / 2 - NutBoltHoleL / 2, False, 0, Nothing, 0)
            Part.InsertAxis2(True) '基准轴2
            Part.Extension.SelectByID2("基准轴2", "AXIS", 0, 0, 0, False, 0, Nothing, 0)
            Part.BlankRefGeom() '隐藏基准轴2
            Part.Extension.SelectByID2("切除-拉伸2", "BODYFEATURE", 0, 0, 0, False, 4, Nothing, 0)
            Part.Extension.SelectByID2("基准轴1", "AXIS", 0, 0, 0, True, 1, Nothing, 0)
            Part.FeatureManager.FeatureCircularPattern4(NutBoltNum, 2 * PI, False, "NULL", False, True, False) '圆周阵列
            Part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 32, Nothing, 0) '选中前视基准面
            Sketchmer.InsertSketch(True) '新建草图6
            Sketchmer.CreateCircleByRadius(0, 0, 0, D2Rim / 2) '画圆
            Part.FeatureManager.FeatureCut3(False, False, False, 1, 1, 0.01, 0.01, False, False, False, False, 0, 0, False, False, False, False, False, True, True, True, True, False, 0, 0, False) '拉伸切除
            Part.Extension.SelectByID2("", "FACE", 0, D2Rim / 2 + k, -b2 / 2, False, 0, Nothing, 0)
            Sketchmer.InsertSketch(True) '新建草图7
            Sketchmer.CreateCircleByRadius(0, 0, 0, D2Rim / 2 + k / 2) '画圆
            Part.FeatureManager.FeatureCut3(True, False, False, 0, 0, k / 2, k / 2, False, False, False, False, 0, 0, False, False, False, False, False, True, True, True, True, False, 0, 0, False) '拉伸切除
            Part.EditRebuild3() '重建模型
            Part.ClearSelection2（True）
            Part.ShowNamedView2("*等轴测", 7) '正等测视角
            Part.ViewZoomtofit2() '整屏显示
            Part.SaveAs3(TyreTypeWormGearRimPath, 0, 2) '文件保存
            Part.Extension.SelectByID2("切除-拉伸4", "BODYFEATURE", 0, 0, 0, True, 0, Nothing, 0)
            Part.Extension.SelectByID2("切除-拉伸3", "BODYFEATURE", 0, 0, 0, True, 0, Nothing, 0)
            Part.EditDelete()
            Part.Extension.SelectByID2("草图7", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
            Part.Extension.SelectByID2("草图6", "SKETCH", 0, 0, 0, True, 0, Nothing, 0)
            Part.EditDelete()
            Part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 32, Nothing, 0) '选中前视基准面
            Sketchmer.InsertSketch(True) '新建草图8
            Sketchmer.CreateCircleByRadius(0, 0, 0, D2Rim / 2 + k / 2) '画圆
            Part.FeatureManager.FeatureCut3(False, True, False, 9, 1, 0.01, 0.01, False, False, False, False, 0, 0, False, False, False, False, False, True, True, True, True, False, 0, 0, False) '拉伸切除
            Part.Extension.SelectByID2("", "FACE", 0, D2Rim / 2 - k / 2, b2 / 2, False, 0, Nothing, 0)
            Sketchmer.InsertSketch(True) '新建草图9
            Sketchmer.CreateCircleByRadius(0, 0, 0, D2Rim / 2) '画圆
            Part.FeatureManager.FeatureCut3(True, True, False, 0, 0, b2 - k / 2, b2 - k / 2, False, False, False, False, 0, 0, False, False, False, False, False, True, True, True, True, False, 0, 0, False) '拉伸切除
            Part.EditRebuild3() '重建模型
            Part.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, False, 32, Nothing, 0) '选中上视基准面
            Sketchmer.InsertSketch(True) '新建草图10
            Part.ShowNamedView2("*上视", 5)
            Part.ViewZoomtofit2() '整屏显示
            Part.SketchManager.CreateCenterRectangle(0, 0, 0, Keywayb / 2, L2Hub, 0)
            Part.SketchAddConstraints("sgFIXED")
            Sketchmer.InsertSketch(True) '完成草图10
            Select Case ComboBoxKeyWayNum.Text
                Case 1
                    Featmgr.FeatureCut3(True, False, True, 0, 0, D2Axis / 2 + Keywayh / 2, D2Axis / 2 + Keywayh / 2, False, False, False, False, 0, 0, False, False, False, False, False, True, True, True, True, False, 0, 0, False) '拉伸切除
                Case 2
                    Featmgr.FeatureCut3(False, False, True, 0, 0, D2Axis / 2 + Keywayh / 2, D2Axis / 2 + Keywayh / 2, False, False, False, False, 0, 0, False, False, False, False, False, True, True, True, True, False, 0, 0, False) '拉伸切除
            End Select
            Part.ClearSelection2（True）
            Part.ShowNamedView2("*等轴测", 7) '正等测视角
            Part.ViewZoomtofit2() '整屏显示
            Part.SaveAs3(TyreTypeWormGearHubPath, 0, 2) '文件保存
            Swapp.CloseAllDocuments(True) '文件关闭
            Swapp.NewDocument("C:\ProgramData\SolidWorks\SolidWorks " & SolidworksEdition & "\templates\gb_assembly.asmdot", 0, 0, 0)
            Assem = Swapp.ActiveDoc
            Part = Swapp.ActiveDoc
            Dim DocumentError&, DocumentWarn&, CoordinationError&
            Dim PartRim, PartHub, PartNutBolt As SldWorks.ModelDoc2
            Dim TyreTypeWormGearTitle$, PartRimTitle$, PartHubTitle$, NutBoltTitle$
            TyreTypeWormGearTitle = Assem.GetTitle() '获取装配体名称
            PartRim = Swapp.OpenDoc6(TyreTypeWormGearRimPath, 1, 0, "", DocumentError, DocumentWarn) '打开轮缘
            PartRimTitle = PartRim.GetTitle() '获取轮缘文件名
            PartRimTitle = PartRimTitle.Substring(0, InStrRev(PartRimTitle, ".") - 1) '提取轮缘名称
            PartRim.Visible = False '隐藏轮缘文件
            Part.AddComponent4(TyreTypeWormGearRimPath, "默认", 0, 0, 0) '添加轮缘
            PartHub = Swapp.OpenDoc6(TyreTypeWormGearHubPath, 1, 0, "", DocumentError, DocumentWarn) '打开轮毂
            PartHubTitle = PartHub.GetTitle() '获取轮毂文件名
            PartHubTitle = PartHubTitle.Substring(0, InStrRev(PartHubTitle, ".") - 1) '提取轮毂名称
            PartHub.Visible = False '隐藏轮毂文件
            Part.AddComponent4(TyreTypeWormGearHubPath, "默认", 0, 0, 0) '添加轮毂
            PartNutBolt = Swapp.OpenDoc6(NutBoltPath, 1, 0, "", DocumentError, DocumentWarn) '打开螺钉
            NutBoltTitle = PartNutBolt.GetTitle() '获取螺钉文件名
            NutBoltTitle = NutBoltTitle.Substring(0, InStrRev(NutBoltTitle, ".") - 1) '提取螺钉名称
            PartNutBolt.Visible = False '隐藏螺钉文件
            Part.AddComponent4(NutBoltPath, "默认", 0, 0, 0) '添加螺钉
            Part.ShowNamedView2("*等轴测", 7) '正等测视角
            Part.ViewZoomtofit2() '整屏显示
            Part.Extension.SelectByID2("前视基准面@" & PartRimTitle & "-1@" & TyreTypeWormGearTitle, "PLANE", 0, 0, 0, False, 1, Nothing, 0) '选中轮缘前视基准面
            Part.Extension.SelectByID2("前视基准面@" & PartHubTitle & "-1@" & TyreTypeWormGearTitle, "PLANE", 0, 0, 0, True, 1, Nothing, 0) '选中轮毂前视基准面
            Assem.AddMate5(0, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, CoordinationError) '重合配合
            Part.EditRebuild3() '重建模型
            Part.Extension.SelectByID2("前视基准面@" & NutBoltTitle & "-1@" & TyreTypeWormGearTitle, "PLANE", 0, 0, 0, False, 1, Nothing, 0)
            Part.Extension.SelectByID2("基准面2@" & PartRimTitle & "-1@" & TyreTypeWormGearTitle, "PLANE", 0, 0, 0, True, 1, Nothing, 0)
            Assem.AddMate5(5, -1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, CoordinationError) '距离配合
            Part.EditRebuild3() '重建模型
            Part.Extension.SelectByID2("基准轴1@" & PartRimTitle & "-1@" & TyreTypeWormGearTitle, "AXIS", 0, 0, 0, False, 1, Nothing, 0)
            Part.Extension.SelectByID2("基准轴1@" & PartHubTitle & "-1@" & TyreTypeWormGearTitle, "AXIS", 0, 0, 0, True, 1, Nothing, 0)
            Assem.AddMate5(5, -1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, CoordinationError) '距离配合
            Part.EditRebuild3() '重建模型
            Part.Extension.SelectByID2("基准轴2@" & PartRimTitle & "-1@" & TyreTypeWormGearTitle, "AXIS", 0, 0, 0, False, 1, Nothing, 0)
            Part.Extension.SelectByID2("基准轴2@" & PartHubTitle & "-1@" & TyreTypeWormGearTitle, "AXIS", 0, 0, 0, True, 1, Nothing, 0)
            Assem.AddMate5(5, -1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, CoordinationError) '距离配合
            Part.EditRebuild3() '重建模型
            Part.Extension.SelectByID2("基准轴2@" & PartRimTitle & "-1@" & TyreTypeWormGearTitle, "AXIS", 0, 0, 0, False, 1, Nothing, 0)
            Part.Extension.SelectByID2("基准轴1@" & NutBoltTitle & "-1@" & TyreTypeWormGearTitle, "AXIS", 0, 0, 0, True, 1, Nothing, 0)
            Assem.AddMate5(5, -1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, CoordinationError) '距离配合
            Part.EditRebuild3() '重建模型
            Part.Extension.SelectByID2("右视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中右视基准面
            Part.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, True, 0, Nothing, 0) '选中上视基准面
            Part.InsertAxis2(True) '新建基准轴1
            Part.ShowNamedView2("*等轴测", 7) '正等测视角
            Part.ViewZoomtofit2() '整屏显示
            Part.Extension.SelectByID2(NutBoltTitle & "-1@" & TyreTypeWormGearTitle, "COMPONENT", 0, 0, 0, False, 1, Nothing, 0)
            Part.Extension.SelectByID2("基准轴1", "AXIS", 0, 0, 0, True, 2, Nothing, 0)
            Part.FeatureManager.FeatureCircularPattern3(NutBoltNum, 1.5707963267949, False, "NULL", False, False) '圆周阵列零件
            Part.Extension.SelectByID2("基准轴1", "AXIS", 0, 0, 0, False, 0, Nothing, 0)
            Part.BlankRefGeom() '隐藏基准轴1
            Part.EditRebuild3() '重建模型
            Part.ClearSelection2（True）
            Part.ShowNamedView2("*等轴测", 7) '正等测视角
            Part.ViewZoomtofit2() '整屏显示
            Part.SaveAs3(TyreTypeWormGearPath, 0, 2)
            If CheckBoxCloseFile.Checked = True Then Swapp.CloseAllDocuments(True)
            Swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, Inputsize) '输入尺寸值还原
        Else
            Dim TyreTypeWormGearRimPath$ = FilePath & "\蜗轮(轮箍式轮缘)" & z2 & "X" & m & ".SLDPRT"
            Dim TyreTypeWormGearHubPath$ = FilePath & "\蜗轮(轮箍式轮毂)" & z2 & "X" & m & ".SLDPRT"
            Dim TyreTypeWormGearPath$ = FilePath & "\蜗轮(轮箍式)" & z1 & "X" & z2 & "X" & m & ".SLDASM"
            AssemWormGearTitle = "蜗轮(轮箍式)" & z1 & "X" & z2 & "X" & m & ".SLDASM"
            AssemWormGearPath = TyreTypeWormGearPath
            Swapp = CreateObject("Sldworks.application")
            Swapp.NewDocument("C:\ProgramData\SolidWorks\SOLIDWORKS " & SolidworksEdition & "\templates\gb_part.prtdot", 0, 0, 0) '新建零件
            Part = Swapp.ActiveDoc
            Sketchmer = Part.SketchManager
            Featmgr = Part.FeatureManager
            Inputsize = Swapp.GetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate) '输入尺寸值记录
            Swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, False) '输入尺寸值关闭
            Part.Extension.SelectByID2("右视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中右视基准面
            Part.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, True, 0, Nothing, 0) '选中上视基准面
            Part.InsertAxis2(True) '新建基准轴1
            Part.Extension.SelectByID2("右视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0) '选中右视基准面
            Dim Line0, Line1, Line2, Line3, Line4, Line5, Line6, Line7, Line8 As SldWorks.SketchSegment
            Dim Arc1, Arc2, Arc3, Arc4, Arc5 As SldWorks.SketchSegment
            Sketchmer.InsertSketch(True) '新建草图1
            Line1 = Sketchmer.CreateLine(0, D2Axis / 2, 0, L2Hub / 2, D2Axis / 2, 0) '孔径直线
            Line2 = Sketchmer.CreateLine(L2Hub / 2, D2Axis / 2, 0, L2Hub / 2, D2Hub / 2, 0) '轮毂侧面
            Line3 = Sketchmer.CreateLine(L2Hub / 2, D2Hub / 2, 0, B2Spoke / 2, D2Hub / 2, 0) '轮毂外圆
            Line4 = Sketchmer.CreateLine(B2Spoke / 2, D2Hub / 2, 0, B2Spoke / 2, D2Rim / 2 - k, 0) '凹槽面
            Line5 = Sketchmer.CreateLine(B2Spoke / 2, D2Rim / 2 - k, 0, b2 / 2, D2Rim / 2 - k, 0) '轮缘内侧
            Line6 = Sketchmer.CreateLine(b2 / 2, D2Rim / 2 - k, 0, b2 / 2, dm2 / 2, 0) '齿宽侧面
            Line7 = Sketchmer.CreateLine(b2 / 2, dm2 / 2, 0, 0, dm2 / 2, 0) '齿顶面
            Arc1 = Sketchmer.CreateCircleByRadius(0, a, 0, rg2) '画小圆
            Line0 = Sketchmer.CreateCenterLine(0, 0, 0, 0, a - rg2, 0) '画中心线
            Arc1.Select4(False, Nothing) '选中小圆
            Sketchmer.SketchTrim(0, 0, a + rg2, 0) '修剪小圆
            Line7.Select4(False, Nothing) '选中齿顶面
            Sketchmer.SketchTrim(0, 0, dm2 / 2, 0) '修剪齿顶面
            If RCast <> 0 Then
                Line3.Select4(False, Nothing)
                Line4.Select4(True, Nothing)
                Arc3 = Sketchmer.CreateFillet(RCast, 2) '倒圆角
                Line4.Select4(False, Nothing)
                Line5.Select4(True, Nothing)
                Arc4 = Sketchmer.CreateFillet(RCast, 2) '倒圆角
            End If
            Line6.Select4(False, Nothing)
            Line7.Select4(True, Nothing)
            Line8 = Sketchmer.CreateChamfer(1, m / 1000, m / 1000) '倒角
            Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
            Line1.Select4(True, Nothing)
            Part.AddDimension2(0, D2Axis / 4, L2Hub / 2 + m / 1000) '标注轴孔半径
            Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
            Line3.Select4(True, Nothing)
            Part.AddDimension2(0, D2Hub / 4, L2Hub / 2 + 2 * m / 1000) '标注轮毂外径
            Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
            Line5.Select4(True, Nothing)
            Part.AddDimension2(0, D2Hub / 4, L2Hub / 2 + 3 * m / 1000) '标注轮缘半径
            Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
            Line7.Select4(True, Nothing)
            Part.AddDimension2(0, D2Hub / 4, L2Hub / 2 + 4 * m / 1000) '标注最大半径
            Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
            Line2.Select4(True, Nothing)
            Part.AddDimension2(0, D2Axis / 2 - 2 * m / 1000, L2Hub / 4) '标注轮廓宽度的一半
            Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
            Line4.Select4(True, Nothing)
            Part.AddDimension2(0, D2Axis / 2 - m / 1000, L2Hub / 4) '标注凹槽宽度的一半
            Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
            Line6.Select4(True, Nothing)
            Part.AddDimension2(0, D2Axis / 2 - 3 * m / 1000, L2Hub / 4) '标注齿宽
            Part.Extension.SelectByID2("", "EXTSKETCHPOINT", 0, 0, 0, False, 0, Nothing, 0)
            Arc1.Select4(True, Nothing)
            Part.AddDimension2(0, D2Hub / 4, L2Hub / 2 + 4 * m / 1000) '标注中心距
            Arc1.Select4(True, Nothing)
            Part.AddDimension2(-b1 / 2, dm2 / 2, 0) '给小圆标尺寸
            Line1.SelectChain(False, Nothing)
            Line0.Select4(True, Nothing)
            Part.SketchMirror() '镜像
            Part.Extension.SelectByID2("基准轴1", "AXIS", 0, 0, 0, False, 16, Nothing, 0)
            Part.FeatureManager.FeatureRevolve2(True, True, False, False, False, False, 0, 0, 2 * PI, 0, False, False, 0.01, 0.01, 0, 0, 0, True, True, True) '旋转
            Part.ShowNamedView2("*等轴测", 7) '正等测视角
            Part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 32, Nothing, 0) '选中前视基准面
            Sketchmer.InsertSketch(True) '新建草图2
            Sketchmer.CreateCircleByRadius(0, D2Spoke / 2, 0, D2SpokeHole / 2) '画圆
            Featmgr.FeatureCut3(False, False, False, 9, 1, 0.01, 0.01, False, False, False, False, 0, 0, False, False, False, False, False, True, True, True, True, False, 0, 0, False) '拉伸切除轮辐孔
            Part.Extension.SelectByID2("", "FACE", 0, D2Spoke / 2 - D2SpokeHole / 2, 0, False, 0, Nothing, 0)
            Part.InsertAxis2(True) '新建基准轴2
            Part.Extension.SelectByID2("基准轴2", "AXIS", 0, 0, 0, False, 0, Nothing, 0)
            Part.BlankRefGeom() '隐藏基准轴2
            Part.Extension.SelectByID2("切除-拉伸1", "BODYFEATURE", 0, 0, 0, False, 4, Nothing, 0)
            Part.Extension.SelectByID2("基准轴1", "AXIS", 0, 0, 0, True, 1, Nothing, 0)
            Part.FeatureManager.FeatureCircularPattern4(SpokeHoleNum, 2 * PI, False, "NULL", False, True, False) '圆周阵列
            Part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 32, Nothing, 0) '选中前视基准面
            Sketchmer.InsertSketch(True) '新建草图3
            Part.ShowNamedView2("*前视", 7) '前视角
            Part.ViewZoomtofit2() '整屏显示
            Dim t1#, t2#
            t1 = 0
            t2 = Tan(Acos(db2 / dm2))
            Arc1 = Sketchmer.CreateCircleByRadius(0, 0, 0, df2 / 2) '齿根圆
            Arc2 = Sketchmer.CreateCircleByRadius(0, 0, 0, dm2 / 2) '齿顶圆
            Line1 = Sketchmer.CreateEquationSpline2(db2 * 1000 / 2 & "*(sin(t)-t*cos(t))", db2 * 1000 / 2 & "*(cos(t)+t*sin(t))", "", t1, t2, False, 0, 0, 0, True, True) '绘制渐开线
            Line8 = Sketchmer.CreateCenterLine(0, 0, 0, 0.005, 0.005, 0) '绘制中心线
            Line8.Angle = PI / 2 + (Acos(db2 / d2) - Tan(Acos(db2 / d2)) + PI * 0.5 / z2) '中心线角度赋值
            If db2 > df2 Then
                Line2 = Sketchmer.CreateLine(0, db2 / 2, 0, 0, df2 / 2, 0) '绘制过渡曲线
                Line1.Select4(False, Nothing)
                Line2.Select4(True, Nothing)
                Line8.Select4(True, Nothing)
                Part.SketchMirror() '镜像
            Else
                Line1.Select4(False, Nothing)
                Sketchmer.SketchTrim(0, 0, db2 / 2, 0)
                Line1.Select4(False, Nothing)
                Line8.Select4(True, Nothing)
                Part.SketchMirror() '镜像
            End If
            Arc1.Select4(False, Nothing)
            Sketchmer.SketchTrim(0, 0, -df2 / 2, 0)
            Arc2.Select4(False, Nothing)
            Sketchmer.SketchTrim(0, 0, -dm2 / 2, 0)
            Line1.SelectChain(False, Nothing)
            Part.SketchAddConstraints("sgFIXED")
            Sketchmer.InsertSketch(True) '完成草图3
            Part.Extension.SelectByID2("Line1@草图3", "EXTSKETCHSEGMENT", 0, 0, 0, False, 0, Nothing, 0)
            Part.Extension.SelectByID2("基准轴1", "AXIS", 0, 0, 0, True, 1, Nothing, 1)
            Featmgr.InsertRefPlane(4, 0, 4, 0, 0, 0) '新建基准面1
            Part.Extension.SelectByID2("基准面1", "PLANE", 0, 0, 0, False, 32, Nothing, 0) '选中基准面1
            Sketchmer.InsertSketch(True) '新建草图4
            Arc5 = Sketchmer.CreateCircleByRadius(0, a, 0, d1 / 2) '画圆
            Part.AddDimension2(0, a, 0) '给圆标直径
            Sketchmer.InsertSketch(True) '完成草图4
            Part.Extension.SelectByID2("草图3", "SKETCH", 0, 0, 0, False, 1, Nothing, 0) '选中扫描轮廓
            Part.Extension.SelectByID2("草图4", "SKETCH", 0, 0, 0, True, 4, Nothing, 0) '选中引导线
            Part.FeatureManager.InsertCutSwept4(False, True, 0, False, False, 0, 0, False, 0, 0, 0, 0, True, True, 0, True, True, True, False) '扫描切除
            Part.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, False, 2, Nothing, 0)
            Part.Extension.SelectByID2("切除-扫描1", "BODYFEATURE", 0, 0, 0, True, 1, Nothing, 0)
            Part.FeatureManager.InsertMirrorFeature(False, False, False, False) '镜像
            Part.Extension.SelectByID2("镜向1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0) '选中镜像
            Part.EditSuppress2() '压缩镜像
            Part.Extension.SelectByID2("切除-扫描1", "BODYFEATURE", 0, 0, 0, False, 4, Nothing, 0)
            Part.Extension.SelectByID2("基准轴1", "AXIS", 0, 0, 0, True, 1, Nothing, 0)
            Part.FeatureManager.FeatureCircularPattern3(z2, 2 * PI, False, "NULL", False, True)
            Part.Extension.SelectByID2("基准轴1", "AXIS", 0, 0, 0, False, 0, Nothing, 0)
            Part.Extension.SelectByID2("基准面1", "PLANE", 0, 0, 0, True, 0, Nothing, 0) '选中基准面1
            Part.BlankRefGeom() '隐藏基准轴1和基准面1
            Part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 32, Nothing, 0) '选中前视基准面
            Sketchmer.InsertSketch(True) '新建草图5
            Sketchmer.CreateCircleByRadius(0, 0, 0, D2Rim / 2) '画圆
            Part.FeatureManager.FeatureCut3(False, False, False, 1, 1, 0.01, 0.01, False, False, False, False, 0, 0, False, False, False, False, False, True, True, True, True, False, 0, 0, False) '拉伸切除
            Part.Extension.SelectByID2("", "FACE", 0, D2Rim / 2 + k, -b2 / 2, False, 0, Nothing, 0)
            Sketchmer.InsertSketch(True) '新建草图6
            Sketchmer.CreateCircleByRadius(0, 0, 0, D2Rim / 2 + k / 2) '画圆
            Part.FeatureManager.FeatureCut3(True, False, False, 0, 0, k / 2, k / 2, False, False, False, False, 0, 0, False, False, False, False, False, True, True, True, True, False, 0, 0, False) '拉伸切除
            Part.Extension.SelectByID2("", "EDGE", 0, D2Rim / 2, b2 / 2, False, 0, Nothing, 0) '选择倒角的边
            Part.FeatureManager.InsertFeatureChamfer(4, 1, 0.003, PI / 4, 0, 0, 0, 0) '倒角
            Part.EditRebuild3() '重建模型
            Part.ClearSelection2（True）
            Part.ShowNamedView2("*等轴测", 7) '正等测视角
            Part.ViewZoomtofit2() '整屏显示
            Part.SaveAs3(TyreTypeWormGearRimPath, 0, 2) '文件保存
            Part.Extension.SelectByID2("倒角1", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
            Part.Extension.SelectByID2("切除-拉伸3", "BODYFEATURE", 0, 0, 0, True, 0, Nothing, 0)
            Part.Extension.SelectByID2("切除-拉伸2", "BODYFEATURE", 0, 0, 0, True, 0, Nothing, 0)
            Part.EditDelete()
            Part.Extension.SelectByID2("草图6", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
            Part.Extension.SelectByID2("草图5", "SKETCH", 0, 0, 0, True, 0, Nothing, 0)
            Part.EditDelete()
            Part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 32, Nothing, 0) '选中前视基准面
            Sketchmer.InsertSketch(True) '新建草图7
            Sketchmer.CreateCircleByRadius(0, 0, 0, D2Rim / 2 + k / 2) '画圆
            Part.FeatureManager.FeatureCut3(False, True, False, 9, 1, 0.01, 0.01, False, False, False, False, 0, 0, False, False, False, False, False, True, True, True, True, False, 0, 0, False) '拉伸切除
            Part.Extension.SelectByID2("", "FACE", 0, D2Rim / 2 - k / 2, b2 / 2, False, 0, Nothing, 0)
            Sketchmer.InsertSketch(True) '新建草图7
            Sketchmer.CreateCircleByRadius(0, 0, 0, D2Rim / 2) '画圆
            Part.FeatureManager.FeatureCut3(True, True, False, 0, 0, b2 - k / 2, b2 - k / 2, False, False, False, False, 0, 0, False, False, False, False, False, True, True, True, True, False, 0, 0, False) '拉伸切除
            Part.Extension.SelectByID2("", "EDGE", 0, D2Rim / 2, b2 / 2, False, 0, Nothing, 0) '选择倒角的边
            Part.FeatureManager.InsertFeatureChamfer(4, 1, 0.001, PI / 18, 0, 0, 0, 0) '倒角
            Part.EditRebuild3() '重建模型
            Part.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, False, 32, Nothing, 0) '选中上视基准面
            Sketchmer.InsertSketch(True) '新建草图8
            Part.ShowNamedView2("*上视", 5)
            Part.ViewZoomtofit2() '整屏显示
            Part.SketchManager.CreateCenterRectangle(0, 0, 0, Keywayb / 2, L2Hub, 0)
            Part.SketchAddConstraints("sgFIXED")
            Sketchmer.InsertSketch(True) '完成草图8
            Select Case ComboBoxKeyWayNum.Text
                Case 1
                    Featmgr.FeatureCut3(True, False, True, 0, 0, D2Axis / 2 + Keywayh / 2, D2Axis / 2 + Keywayh / 2, False, False, False, False, 0, 0, False, False, False, False, False, True, True, True, True, False, 0, 0, False) '拉伸切除
                Case 2
                    Featmgr.FeatureCut3(False, False, True, 0, 0, D2Axis / 2 + Keywayh / 2, D2Axis / 2 + Keywayh / 2, False, False, False, False, 0, 0, False, False, False, False, False, True, True, True, True, False, 0, 0, False) '拉伸切除
            End Select
            Part.ClearSelection2（True）
            Part.ShowNamedView2("*等轴测", 7) '正等测视角
            Part.ViewZoomtofit2() '整屏显示
            Part.SaveAs3(TyreTypeWormGearHubPath, 0, 2) '文件保存
            Swapp.CloseAllDocuments(True) '文件关闭
            Swapp.NewDocument("C:\ProgramData\SolidWorks\SolidWorks " & SolidworksEdition & "\templates\gb_assembly.asmdot", 0, 0, 0)
            Assem = Swapp.ActiveDoc
            Part = Swapp.ActiveDoc
            Dim DocumentError&, DocumentWarn&, CoordinationError&
            Dim PartRim, PartHub As SldWorks.ModelDoc2
            Dim TyreTypeWormGearTitle$, PartRimTitle$, PartHubTitle$
            TyreTypeWormGearTitle = Assem.GetTitle() '获取装配体名称
            PartRim = Swapp.OpenDoc6(TyreTypeWormGearRimPath, 1, 0, "", DocumentError, DocumentWarn) '打开轮缘
            PartRimTitle = PartRim.GetTitle() '获取轮缘文件名
            PartRimTitle = PartRimTitle.Substring(0, InStrRev(PartRimTitle, ".") - 1) '提取轮缘名称
            PartRim.Visible = False '隐藏轮缘文件
            Part.AddComponent4(TyreTypeWormGearRimPath, "默认", 0, 0, 0) '添加轮缘
            PartHub = Swapp.OpenDoc6(TyreTypeWormGearHubPath, 1, 0, "", DocumentError, DocumentWarn) '打开轮毂
            PartHubTitle = PartHub.GetTitle() '获取轮毂文件名
            PartHubTitle = PartHubTitle.Substring(0, InStrRev(PartHubTitle, ".") - 1) '提取轮毂名称
            PartHub.Visible = False '隐藏轮毂文件
            Part.AddComponent4(TyreTypeWormGearHubPath, "默认", 0, 0, 0) '添加轮毂
            Part.ShowNamedView2("*等轴测", 7) '正等测视角
            Part.ViewZoomtofit2() '整屏显示
            Part.Extension.SelectByID2("前视基准面@" & PartHubTitle & "-1@" & TyreTypeWormGearTitle, "PLANE", 0, 0, 0, False, 1, Nothing, 0) '选中轮毂前视基准面
            Part.Extension.SelectByID2("前视基准面@" & PartRimTitle & "-1@" & TyreTypeWormGearTitle, "PLANE", 0, 0, 0, True, 1, Nothing, 0) '选中轮毂前视基准面
            Assem.AddMate5(0, -1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, CoordinationError) '重合配合
            Part.EditRebuild3() '重建模型
            Part.Extension.SelectByID2("基准轴1@" & PartRimTitle & "-1@" & TyreTypeWormGearTitle, "AXIS", 0, 0, 0, False, 1, Nothing, 0)
            Part.Extension.SelectByID2("基准轴1@" & PartHubTitle & "-1@" & TyreTypeWormGearTitle, "AXIS", 0, 0, 0, True, 1, Nothing, 0)
            Assem.AddMate5(5, -1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, CoordinationError) '距离配合
            Part.EditRebuild3() '重建模型
            Part.Extension.SelectByID2("基准轴2@" & PartRimTitle & "-1@" & TyreTypeWormGearTitle, "AXIS", 0, 0, 0, False, 1, Nothing, 0)
            Part.Extension.SelectByID2("基准轴2@" & PartHubTitle & "-1@" & TyreTypeWormGearTitle, "AXIS", 0, 0, 0, True, 1, Nothing, 0)
            Assem.AddMate5(5, -1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, CoordinationError) '距离配合
            Part.EditRebuild3() '重建模型
            Part.ClearSelection2（True）
            Part.ShowNamedView2("*等轴测", 7) '正等测视角
            Part.ViewZoomtofit2() '整屏显示
            Part.SaveAs3(TyreTypeWormGearPath, 0, 2)
            If CheckBoxCloseFile.Checked = True Then Swapp.CloseAllDocuments(True)
            Swapp.SetUserPreferenceToggle(SwConst.swUserPreferenceToggle_e.swInputDimValOnCreate, Inputsize) '输入尺寸值还原
        End If
    End Sub
    '轮箍式蜗轮是否添加螺钉
    Private Sub CheckBoxAddNutBolt_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxAddNutBolt.CheckedChanged, CheckBoxCloseFile.CheckedChanged '轮箍式蜗轮是否添加螺钉
        Select Case LabelNutBoltNum.Enabled
            Case False
                LabelNutBoltNum.Enabled = True
                TextBoxNutBoltNum.Enabled = True
            Case True
                LabelNutBoltNum.Enabled = False
                TextBoxNutBoltNum.Enabled = False
        End Select
    End Sub



    '选择需要装配的蜗轮类型
    Private Sub ComboBoxWormGearTypeAssem_TextChanged(sender As Object, e As EventArgs) Handles ComboBoxWormGearTypeAssem.TextChanged
        Select Case ComboBoxWormGearTypeAssem.Text
            Case "整体式"
                AssemWormGearType = 1
                OpenFileDialogWormGearPath.Filter = "SWPart (*.SLDPRT)|*.SLDPRT|SWAssem (*.SLDASM)|*.SLDASM|所有文件|*.*"
            Case Else
                AssemWormGearType = 2
                OpenFileDialogWormGearPath.Filter = "SWAssem (*.SLDASM)|*.SLDASM|SWPart (*.SLDPRT)|*.SLDPRT|所有文件|*.*"
        End Select
    End Sub
    '选择需要装配的蜗杆
    Private Sub ButtonSelectWorm_Click(sender As Object, e As EventArgs) Handles ButtonSelectWorm.Click
        If OpenFileDialogWormPath.ShowDialog = DialogResult.OK Then TextboxWormPath.Text = OpenFileDialogWormPath.FileName
    End Sub
    '选择需要装配的蜗轮
    Private Sub ButtonSelectWormGear_Click(sender As Object, e As EventArgs) Handles ButtonSelectWormGear.Click
        If OpenFileDialogWormGearPath.ShowDialog = DialogResult.OK Then TextboxWormGearPath.Text = OpenFileDialogWormGearPath.FileName
    End Sub
    '单击装配
    Private Sub ButtonCreateAssem_Click(sender As Object, e As EventArgs) Handles ButtonCreateAssem.Click
        If TextboxWormPath.Text = "" Then
            MessageBox.Show("请选择蜗杆路径！", "未选择"， MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        If TextboxWormGearPath.Text = "" Then
            MessageBox.Show("请选择蜗轮路径！", "未选择"， MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        If IsNumeric(TextBoxAssema.Text) Then '判断Assema输入的值是否正确
            If TextBoxAssema.Text <= 0 Then
                MessageBox.Show("请输入正确的蜗轮轴径！", "格式不正确"， MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If
        Else
            MessageBox.Show("请输入正确的蜗轮轴径！", "格式不正确"， MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        Me.Hide()
        Select Case ComboBoxWormTypeAssem.Text
            Case "阿基米德圆柱蜗杆(ZA型)"
                Call AssemZA()
            Case "法向直廓圆柱蜗杆(ZN型)"
                Call AssemZN()
        End Select
        Me.Show()
    End Sub

    '装配过程
    Public Sub AssemZA()
        SolidworksEdition = ComboboxEdition.Text
        Assema = TextBoxAssema.Text / 1000
        Direction = ComboBoxDirectionAssem.Text
        Swapp = CreateObject("Sldworks.application")
        Swapp.Visible = True
        Swapp.NewDocument("C:\ProgramData\SolidWorks\SolidWorks " & SolidworksEdition & "\templates\gb_assembly.asmdot", 0, 0, 0)
        Assem = Swapp.ActiveDoc
        Part = Swapp.ActiveDoc
        Dim DocumentError&, DocumentWarn&, CoordinationError&
        Dim Part1, Part2 As SldWorks.ModelDoc2
        Dim AssemblyTitle$, WormTitle$, WormGearTitle$
        AssemblyTitle = Assem.GetTitle() '获取装配体名称
        Part1 = Swapp.OpenDoc6(TextboxWormPath.Text, 1, 0, "", DocumentError, DocumentWarn) '打开蜗杆
        WormTitle = Part1.GetTitle() '获取蜗杆文件名
        WormTitle = WormTitle.Substring(0, InStrRev(WormTitle, ".") - 1) '提取蜗杆名称
        Part1.Visible = False '隐藏蜗杆文件
        Part.AddComponent4(TextboxWormPath.Text, "默认", 0, Assema, 0) '添加蜗杆1
        Part.AddComponent4(TextboxWormPath.Text, "默认", 0, 0, 0) '添加蜗杆2
        Part2 = Swapp.OpenDoc6(TextboxWormGearPath.Text, AssemWormGearType, 0, "", DocumentError, DocumentWarn) '打开蜗轮
        WormGearTitle = Part2.GetTitle() '获取蜗轮文件名
        WormGearTitle = WormGearTitle.Substring(0, InStrRev(WormGearTitle, ".") - 1) '提取蜗轮名称
        Part2.Visible = False '隐藏蜗轮文件
        Part.AddComponent4(TextboxWormGearPath.Text, "默认", 0, -Assema, 0) '添加蜗轮
        Dim WormAndWormGearPath$ = FilePath & "\蜗轮蜗杆(" & WormTitle & WormGearTitle & ").SLDASM"
        Part.Extension.SelectByID2(WormTitle & "-1@" & AssemblyTitle, "COMPONENT", 0, 0, 0, False, 0, Nothing, 0) '选择蜗杆1
        Part.EditDelete() '删除蜗杆1
        Part.Extension.SelectByID2("基准轴1@" & WormTitle & "-2@" & AssemblyTitle, "AXIS", 0, 0, 0, False, 1, Nothing, 0) '选择蜗杆基准轴
        Part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, True, 1, Nothing, 0) '选中前视基准面
        Assem.AddMate5(0, -1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, CoordinationError) '重合配合
        Part.EditRebuild3() '重建模型
        Part.Extension.SelectByID2("基准轴1@" & WormTitle & "-2@" & AssemblyTitle, "AXIS", 0, 0, 0, False, 1, Nothing, 0) '选择蜗杆基准轴
        Part.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, True, 1, Nothing, 0) '选中上视基准面
        Assem.AddMate5(0, -1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, CoordinationError) '重合配合
        Part.EditRebuild3() '重建模型
        Part.Extension.SelectByID2("右视基准面@" & WormTitle & "-2@" & AssemblyTitle, "PLANE", 0, 0, 0, False, 1, Nothing, 0) '选择蜗杆右视基准面
        Part.Extension.SelectByID2("右视基准面", "PLANE", 0, 0, 0, True, 1, Nothing, 0) '选中右视基准面
        Assem.AddMate5(0, -1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, CoordinationError) '重合配合
        Part.EditRebuild3() '重建模型
        Part.Extension.SelectByID2("右视基准面@" & WormTitle & "-2@" & AssemblyTitle, "PLANE", 0, 0, 0, False, 1, Nothing, 0) '选择蜗杆右视基准面
        Part.Extension.SelectByID2("基准轴1@" & WormGearTitle & "-1@" & AssemblyTitle, "AXIS", 0, 0, 0, True, 1, Nothing, 0) '选择蜗轮基准轴
        Assem.AddMate5(0, -1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, CoordinationError) '重合配合
        Part.EditRebuild3() '重建模型
        Part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 1, Nothing, 0) '选中前视基准面
        Part.Extension.SelectByID2("前视基准面@" & WormGearTitle & "-1@" & AssemblyTitle, "PLANE", 0, 0, 0, True, 1, Nothing, 0) '选中蜗轮前视基准面
        Assem.AddMate5(0, -1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, CoordinationError) '重合配合
        Part.EditRebuild3() '重建模型
        Part.Extension.SelectByID2("基准轴1@" & WormTitle & "-2@" & AssemblyTitle, "AXIS", 0, 0, 0, False, 1, Nothing, 0) '选择蜗杆基准轴
        Part.Extension.SelectByID2("基准轴1@" & WormGearTitle & "-1@" & AssemblyTitle, "AXIS", 0, 0, 0, True, 1, Nothing, 0) '选择蜗轮基准轴
        Assem.AddMate5(5, -1, False, Assema, Assema, Assema, 0, 0, 0, 0, 0, False, False, 0, CoordinationError) '距离配合
        Part.EditRebuild3() '重建模型

        '未能添加机械配合
        'If Direction = "右旋" Then
        '    Part.Extension.SelectByID2("", "EDGE", -(l1 + b1 / 2), d11 / 2, 0, False, 1, Nothing, 0)
        'Else
        '    Part.Extension.SelectByID2("", "EDGE", l1 + b1 / 2, d11 / 2, 0, False, 1, Nothing, 0)
        'End If
        'Part.Extension.SelectByID2("", "EDGE", 0, -(Assema + D2Hub / 2), L2Hub / 2, True, 1, Nothing, 0)
        'Assem.AddMate5(10, -1, False, 0, 0, 0, z1 / 1000, z2 / 1000, 0, 0, 0, False, False, 0, CoordinationError)

        Part.EditRebuild3() '重建模型
        Part.ClearSelection2（True）
        Part.ShowNamedView2("*等轴测", 7) '正等测视角
        Part.ViewZoomtofit2() '整屏显示
        Part.SaveAs3(WormAndWormGearPath, 0, 2)
        If CheckBoxCloseFile.Checked = True Then Swapp.CloseAllDocuments(True)
    End Sub
    Public Sub AssemZN()
        SolidworksEdition = ComboboxEdition.Text
        Assema = TextBoxAssema.Text / 1000
        Direction = ComboBoxDirectionAssem.Text
        Swapp = CreateObject("Sldworks.application")
        Swapp.Visible = True
        Swapp.NewDocument("C:\ProgramData\SolidWorks\SolidWorks " & SolidworksEdition & "\templates\gb_assembly.asmdot", 0, 0, 0)
        Assem = Swapp.ActiveDoc
        Part = Swapp.ActiveDoc
        Dim DocumentError&, DocumentWarn&, CoordinationError&
        Dim Part1, Part2 As SldWorks.ModelDoc2
        Dim AssemblyTitle$, WormTitle$, WormGearTitle$
        AssemblyTitle = Assem.GetTitle() '获取装配体名称
        Part1 = Swapp.OpenDoc6(TextboxWormPath.Text, 1, 0, "", DocumentError, DocumentWarn) '打开蜗杆
        WormTitle = Part1.GetTitle() '获取蜗杆文件名
        WormTitle = WormTitle.Substring(0, InStrRev(WormTitle, ".") - 1) '提取蜗杆名称
        Part1.Visible = False '隐藏蜗杆文件
        Part.AddComponent4(TextboxWormPath.Text, "默认", 0, Assema, 0) '添加蜗杆1
        Part.AddComponent4(TextboxWormPath.Text, "默认", 0, 0, 0) '添加蜗杆2
        Part2 = Swapp.OpenDoc6(TextboxWormGearPath.Text, AssemWormGearType, 0, "", DocumentError, DocumentWarn) '打开蜗轮
        WormGearTitle = Part2.GetTitle() '获取蜗轮文件名
        WormGearTitle = WormGearTitle.Substring(0, InStrRev(WormGearTitle, ".") - 1) '提取蜗轮名称
        Part2.Visible = False '隐藏蜗轮文件
        Part.AddComponent4(TextboxWormGearPath.Text, "默认", 0, -Assema, 0) '添加蜗轮
        Dim WormAndWormGearPath$ = FilePath & "\蜗轮蜗杆(" & WormTitle & "+" & WormGearTitle & ").SLDASM"
        Part.Extension.SelectByID2(WormTitle & "-1@" & AssemblyTitle, "COMPONENT", 0, 0, 0, False, 0, Nothing, 0) '选择蜗杆1
        Part.EditDelete() '删除蜗杆1
        Part.Extension.SelectByID2("基准轴1@" & WormTitle & "-2@" & AssemblyTitle, "AXIS", 0, 0, 0, False, 1, Nothing, 0) '选择蜗杆基准轴
        Part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, True, 1, Nothing, 0) '选中前视基准面
        Assem.AddMate5(0, -1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, CoordinationError) '重合配合
        Part.EditRebuild3() '重建模型
        Part.Extension.SelectByID2("基准轴1@" & WormTitle & "-2@" & AssemblyTitle, "AXIS", 0, 0, 0, False, 1, Nothing, 0) '选择蜗杆基准轴
        Part.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, True, 1, Nothing, 0) '选中上视基准面
        Assem.AddMate5(0, -1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, CoordinationError) '重合配合
        Part.EditRebuild3() '重建模型
        Part.Extension.SelectByID2("基准面1@" & WormTitle & "-2@" & AssemblyTitle, "PLANE", 0, 0, 0, False, 1, Nothing, 0) '选择蜗杆基准面1
        Part.Extension.SelectByID2("右视基准面", "PLANE", 0, 0, 0, True, 1, Nothing, 0) '选中右视基准面
        Assem.AddMate5(0, -1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, CoordinationError) '重合配合
        Part.EditRebuild3() '重建模型
        Part.Extension.SelectByID2("基准面1@" & WormTitle & "-2@" & AssemblyTitle, "PLANE", 0, 0, 0, False, 1, Nothing, 0) '选择蜗杆基准面1
        Part.Extension.SelectByID2("基准轴1@" & WormGearTitle & "-1@" & AssemblyTitle, "AXIS", 0, 0, 0, True, 1, Nothing, 0) '选择蜗轮基准轴
        Assem.AddMate5(0, -1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, CoordinationError) '重合配合
        Part.EditRebuild3() '重建模型
        Part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 1, Nothing, 0) '选中前视基准面
        Part.Extension.SelectByID2("前视基准面@" & WormGearTitle & "-1@" & AssemblyTitle, "PLANE", 0, 0, 0, True, 1, Nothing, 0) '选中蜗轮前视基准面
        Assem.AddMate5(0, -1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, CoordinationError) '重合配合
        Part.EditRebuild3() '重建模型
        Part.Extension.SelectByID2("基准轴1@" & WormTitle & "-2@" & AssemblyTitle, "AXIS", 0, 0, 0, False, 1, Nothing, 0) '选择蜗杆基准轴
        Part.Extension.SelectByID2("基准轴1@" & WormGearTitle & "-1@" & AssemblyTitle, "AXIS", 0, 0, 0, True, 1, Nothing, 0) '选择蜗轮基准轴
        Assem.AddMate5(5, -1, False, Assema, Assema, Assema, 0, 0, 0, 0, 0, False, False, 0, CoordinationError) '距离配合
        Part.EditRebuild3() '重建模型

        '未能添加机械配合
        'If Direction = "右旋" Then
        '    Part.Extension.SelectByID2("", "EDGE", -(l1 + b1 / 2), d11 / 2, 0, False, 1, Nothing, 0)
        'Else
        '    Part.Extension.SelectByID2("", "EDGE", l1 + b1 / 2, d11 / 2, 0, False, 1, Nothing, 0)
        'End If
        'Part.Extension.SelectByID2("", "EDGE", 0, -(Assema + D2Hub / 2), L2Hub / 2, True, 1, Nothing, 0)
        'Assem.AddMate5(10, -1, False, 0, 0, 0, z1 / 1000, z2 / 1000, 0, 0, 0, False, False, 0, CoordinationError)

        Part.EditRebuild3() '重建模型
        Part.ClearSelection2（True）
        Part.ShowNamedView2("*等轴测", 7) '正等测视角
        Part.ViewZoomtofit2() '整屏显示
        Part.SaveAs3(WormAndWormGearPath, 0, 2)
        If CheckBoxCloseFile.Checked = True Then Swapp.CloseAllDocuments(True)
    End Sub

    '单击一键建模装配
    Private Sub ButtonCreatPartAndAssem_Click(sender As Object, e As EventArgs) Handles ButtonCreatPartAndAssem.Click
        Dim Face$ = "右视基准面"
        Call NumericalCheck()
        If ParameterError = True Then Exit Sub
        Call NumericalCalculation()
        Me.Hide()
        Select Case ComboBoxWormType.Text
            Case "阿基米德圆柱蜗杆(ZA型)"
                Call CreateZATypeWorm()
                Face = "右视基准面"
            Case "法向直廓圆柱蜗杆(ZN型)"
                Call CreateZNTypeWorm()
                Face = "基准面1"
        End Select
        Select Case ComboBoxWormGearType.Text
            Case "整体式"
                Call CreateIntegralTypeWormGear()
                AssemWormGearType = 1
            Case "螺栓连接式"
                Call CreateBoltedTypeWormGear()
                AssemWormGearType = 2
            Case "镶铸式"
                Call CreateInlayTypeWormGear()
                AssemWormGearType = 2
            Case "轮箍式"
                Call CreateTyreTypeWormGear()
                AssemWormGearType = 2
        End Select
        Assema = a
        Swapp = CreateObject("Sldworks.application")
        Swapp.Visible = True
        Swapp.NewDocument("C:\ProgramData\SolidWorks\SolidWorks " & SolidworksEdition & "\templates\gb_assembly.asmdot", 0, 0, 0)
        Assem = Swapp.ActiveDoc
        Part = Swapp.ActiveDoc
        Dim DocumentError&, DocumentWarn&, CoordinationError&
        Dim Part1, Part2 As SldWorks.ModelDoc2
        Dim AssemblyTitle$, WormTitle$, WormGearTitle$
        AssemblyTitle = Assem.GetTitle() '获取装配体名称
        Part1 = Swapp.OpenDoc6(AssemWormPath, 1, 0, "", DocumentError, DocumentWarn) '打开蜗杆
        WormTitle = Part1.GetTitle() '获取蜗杆文件名
        WormTitle = WormTitle.Substring(0, InStrRev(WormTitle, ".") - 1) '提取蜗杆名称
        Part1.Visible = False '隐藏蜗杆文件
        Part.AddComponent4(AssemWormPath, "默认", 0, Assema, 0) '添加蜗杆1
        Part.AddComponent4(AssemWormPath, "默认", 0, 0, 0) '添加蜗杆2
        Part2 = Swapp.OpenDoc6(AssemWormGearPath, AssemWormGearType, 0, "", DocumentError, DocumentWarn) '打开蜗轮
        WormGearTitle = Part2.GetTitle() '获取蜗轮文件名
        WormGearTitle = WormGearTitle.Substring(0, InStrRev(WormGearTitle, ".") - 1) '提取蜗轮名称
        Part2.Visible = False '隐藏蜗轮文件
        Part.AddComponent4(AssemWormGearPath, "默认", 0, -Assema, 0) '添加蜗轮
        Dim WormAndWormGearPath$ = FilePath & "\蜗轮蜗杆(" & WormTitle & WormGearTitle & ").SLDASM"
        Part.Extension.SelectByID2(WormTitle & "-1@" & AssemblyTitle, "COMPONENT", 0, 0, 0, False, 0, Nothing, 0) '选择蜗杆1
        Part.EditDelete() '删除蜗杆1
        Part.Extension.SelectByID2("基准轴1@" & WormTitle & "-2@" & AssemblyTitle, "AXIS", 0, 0, 0, False, 1, Nothing, 0) '选择蜗杆基准轴
        Part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, True, 1, Nothing, 0) '选中前视基准面
        Assem.AddMate5(0, -1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, CoordinationError) '重合配合
        Part.EditRebuild3() '重建模型
        Part.Extension.SelectByID2("基准轴1@" & WormTitle & "-2@" & AssemblyTitle, "AXIS", 0, 0, 0, False, 1, Nothing, 0) '选择蜗杆基准轴
        Part.Extension.SelectByID2("上视基准面", "PLANE", 0, 0, 0, True, 1, Nothing, 0) '选中上视基准面
        Assem.AddMate5(0, -1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, CoordinationError) '重合配合
        Part.EditRebuild3() '重建模型
        Part.Extension.SelectByID2(Face & "@" & WormTitle & "-2@" & AssemblyTitle, "PLANE", 0, 0, 0, False, 1, Nothing, 0) '选择蜗杆基准面
        Part.Extension.SelectByID2("右视基准面", "PLANE", 0, 0, 0, True, 1, Nothing, 0) '选中右视基准面
        Assem.AddMate5(0, -1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, CoordinationError) '重合配合
        Part.EditRebuild3() '重建模型
        Part.Extension.SelectByID2(Face & "@" & WormTitle & "-2@" & AssemblyTitle, "PLANE", 0, 0, 0, False, 1, Nothing, 0) '选择蜗杆右视基准面
        Part.Extension.SelectByID2("基准轴1@" & WormGearTitle & "-1@" & AssemblyTitle, "AXIS", 0, 0, 0, True, 1, Nothing, 0) '选择蜗轮基准轴
        Assem.AddMate5(0, -1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, CoordinationError) '重合配合
        Part.EditRebuild3() '重建模型
        Part.Extension.SelectByID2("前视基准面", "PLANE", 0, 0, 0, False, 1, Nothing, 0) '选中前视基准面
        Part.Extension.SelectByID2("前视基准面@" & WormGearTitle & "-1@" & AssemblyTitle, "PLANE", 0, 0, 0, True, 1, Nothing, 0) '选中蜗轮前视基准面
        Assem.AddMate5(0, -1, False, 0, 0, 0, 0, 0, 0, 0, 0, False, False, 0, CoordinationError) '重合配合
        Part.EditRebuild3() '重建模型
        Part.Extension.SelectByID2("基准轴1@" & WormTitle & "-2@" & AssemblyTitle, "AXIS", 0, 0, 0, False, 1, Nothing, 0) '选择蜗杆基准轴
        Part.Extension.SelectByID2("基准轴1@" & WormGearTitle & "-1@" & AssemblyTitle, "AXIS", 0, 0, 0, True, 1, Nothing, 0) '选择蜗轮基准轴
        Assem.AddMate5(5, -1, False, Assema, Assema, Assema, 0, 0, 0, 0, 0, False, False, 0, CoordinationError) '距离配合
        Part.EditRebuild3() '重建模型
        '未能添加机械配合
        If ComboBoxWormTypeAssem.Text = "阿基米德圆柱蜗杆(ZA型)" Then
            If Direction = "右旋" Then
                Part.Extension.SelectByID2("", "EDGE", -(l1 + b1 / 2), d11 / 2, 0, False, 1, Nothing, 0)
            Else
                Part.Extension.SelectByID2("", "EDGE", l1 + b1 / 2, d11 / 2, 0, False, 1, Nothing, 0)
            End If
            Part.Extension.SelectByID2("", "EDGE", 0, -(Assema + D2Hub / 2), L2Hub / 2, True, 1, Nothing, 0)
            Assem.AddMate5(10, -1, False, 0, 0, 0, z1 / 1000, z2 / 1000, 0, 0, 0, False, False, 0, CoordinationError)
        End If
        Part.EditRebuild3() '重建模型
        Part.ClearSelection2（True）
        Part.ShowNamedView2("*等轴测", 7) '正等测视角
        Part.ViewZoomtofit2() '整屏显示
        Part.SaveAs3(WormAndWormGearPath, 0, 2)
        If CheckBoxCloseFile.Checked = True Then Swapp.CloseAllDocuments(True)
        Me.Show()
    End Sub

    '生成文档
    Private Sub ButtonOutputParameter_Click(sender As Object, e As EventArgs) Handles ButtonOutputParameter.Click
        Call NumericalCheck()
        If ParameterError = True Then
            MessageBox.Show("因输入参数不正确，未能生成参数文件！", "格式不正确"， MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        Call NumericalCalculation()
        Dim NowTime$ = Format(DateTime.Now, "yyyy年MM月dd日hh时mm分ss秒")
        Dim txtlujing$ = Application.StartupPath & "\蜗轮蜗杆参数-" & NowTime & ".txt"
        Dim t As System.IO.StreamWriter = New System.IO.StreamWriter(txtlujing, True, System.Text.Encoding.UTF8)
        Dim Digit% = 3
        t.WriteLine("蜗杆类型:         " & ComboBoxWormType.Text)
        t.WriteLine("蜗轮类型:         " & ComboBoxWormGearType.Text)
        t.WriteLine("")
        t.WriteLine("蜗杆轴向模数m:    " & m & "mm")
        t.WriteLine("蜗杆头数z1:       " & z1)
        t.WriteLine("蜗轮齿数z2:       " & z2)
        t.WriteLine("")
        t.WriteLine("蜗杆直径系数q:    " & FormatNumber(q, Digit))
        t.WriteLine("蜗轮变位系数x2:   " & x2)
        t.WriteLine("齿顶高系数ha*:    " & hax)
        t.WriteLine("齿顶间隙系数c*:   " & cx)
        t.WriteLine("")
        t.WriteLine("法向压力角αn:    " & FormatNumber(Alpha * 180 / PI, Digit) & "°")
        t.WriteLine("导程角γ:         " & FormatNumber(Gamma * 180 / PI, Digit) & "°")
        t.WriteLine("旋向:             " & Direction)
        t.WriteLine("传动比i:          " & FormatNumber(i, Digit))
        t.WriteLine("中心距a:          " & FormatNumber(a * 1000, Digit) & "mm")
        t.WriteLine("")
        t.WriteLine("蜗杆分度圆直径d1: " & FormatNumber(d1 * 1000, Digit) & "mm")
        t.WriteLine("蜗杆齿顶圆直径da1:" & FormatNumber(da1 * 1000, Digit) & "mm")
        t.WriteLine("蜗杆齿根圆直径df1:" & FormatNumber(df1 * 1000, Digit) & "mm")
        t.WriteLine("蜗杆齿顶高ha1:    " & FormatNumber(ha1 * 1000, Digit) & "mm")
        t.WriteLine("蜗杆齿根高hf1:    " & FormatNumber(hf1 * 1000, Digit) & "mm")
        t.WriteLine("蜗杆全齿高h1:     " & FormatNumber(h1 * 1000, Digit) & "mm")
        t.WriteLine("")
        t.WriteLine("蜗轮分度圆直径d2: " & FormatNumber(d2 * 1000, Digit) & "mm")
        t.WriteLine("蜗轮齿顶圆直径da2:" & FormatNumber(da2 * 1000, Digit) & "mm")
        t.WriteLine("蜗轮齿根圆直径df2:" & FormatNumber(df2 * 1000, Digit) & "mm")
        t.WriteLine("蜗轮齿顶高ha2:    " & FormatNumber(ha2 * 1000, Digit) & "mm")
        t.WriteLine("蜗轮齿根高hf2:    " & FormatNumber(hf2 * 1000, Digit) & "mm")
        t.WriteLine("蜗轮全齿高h2:     " & FormatNumber(h2 * 1000, Digit) & "mm")
        t.WriteLine("蜗轮轴径:         " & FormatNumber(D2Axis * 1000, Digit) & "mm")
        t.WriteLine("")
        t.WriteLine("蜗杆齿宽B1:       " & FormatNumber(b1 * 1000, Digit) & "mm")
        t.WriteLine("蜗轮齿宽B2:       " & FormatNumber(b2 * 1000, Digit) & "mm")
        t.WriteLine("")
        t.WriteLine("参数生成时间：" & NowTime)
        t.Close()
        Shell("notepad.exe " & txtlujing, vbNormalFocus)
    End Sub


    '模块1计算
    Private Sub Button1Calculation_Click(sender As Object, e As EventArgs) Handles Button1Calculation.Click
        If (Not IsNumeric(TextBoxiDesign.Text)) Or Val(TextBoxiDesign.Text) <= 0 Then
            MsgBox("传动比的格式不正确！", 0 + 48, "格式不正确")
            Exit Sub
        End If
        If (Not IsNumeric(TextBoxP1Design.Text)) Or Val(TextBoxP1Design.Text) <= 0 Then
            MsgBox("输入功率的格式不正确！", 0 + 48, "格式不正确")
            Exit Sub
        End If
        If (Not IsNumeric(TextBoxn1Design.Text)) Or Val(TextBoxn1Design.Text) < 500 Or Val(TextBoxn1Design.Text) > 1500 Then
            MsgBox("蜗杆转速在500~1500内！", 0 + 48, "格式不正确")
            Exit Sub
        End If
        If (Not IsNumeric(TextBoxTDesign.Text)) Or Val(TextBoxTDesign.Text) <= 0 Then
            MsgBox("使用寿命的格式不正确！", 0 + 48, "格式不正确")
            Exit Sub
        End If
        Panel1_2.Enabled = True
        Panel1_1.Enabled = False
        iDesign = Val(TextBoxiDesign.Text)
        P1Design = Val(TextBoxP1Design.Text)
        n1Design = Val(TextBoxn1Design.Text)
        TDesign = Val(TextBoxTDesign.Text)
        Select Case iDesign
            Case <= 6
                LabeliToz1Design.Text = "由i=" & iDesign & "推荐蜗杆头数为6"
                ComboBoxz1Design.Text = 6
            Case 7, 8
                LabeliToz1Design.Text = "由i=" & iDesign & "推荐蜗杆头数为4"
                ComboBoxz1Design.Text = 4
            Case 9 To 13
                LabeliToz1Design.Text = "由i=" & iDesign & "推荐蜗杆头数为3~4"
                ComboBoxz1Design.Text = 3
            Case 14 To 27
                LabeliToz1Design.Text = "由i=" & iDesign & "推荐蜗杆头数为2~3"
                ComboBoxz1Design.Text = 2
            Case 28 To 40
                LabeliToz1Design.Text = "由i=" & iDesign & "推荐蜗杆头数为1~2"
                ComboBoxz1Design.Text = 1
            Case > 40
                LabeliToz1Design.Text = "由i=" & iDesign & "推荐蜗杆头数为1"
                ComboBoxz1Design.Text = 1
        End Select
        TextBoxz2Design.Text = Round(Val(ComboBoxz1Design.Text) * iDesign)
        TextBoxn2Design.Text = Format(n1Design / iDesign, "##########.###")
        TextBoxEfficiencyDesign.Text = Format((100 - 3.5 * Sqrt(iDesign)) * 0.01, "0.###")
        TextBoxT2Design.Text = Format(9550 * P1Design * Val(TextBoxEfficiencyDesign.Text) * iDesign / n1Design, "##########.###")
        vsDesign500 = 0.000002003 * Pow(P1Design, 5) - 0.000102 * Pow(P1Design, 4) + 0.001854 * Pow(P1Design, 3) - 0.02348 * Pow(P1Design, 2) + 0.416 * P1Design + 1.242
        vsDesign750 = 0.0000005008 * Pow(P1Design, 5) - 0.0000428 * Pow(P1Design, 4) + 0.001665 * Pow(P1Design, 3) - 0.04088 * Pow(P1Design, 2) + 0.687 * P1Design + 1.183
        vsDesign1000 = -0.000001753 * Pow(P1Design, 5) + 0.0000997 * Pow(P1Design, 4) - 0.001659 * Pow(P1Design, 3) - 0.008321 * Pow(P1Design, 2) + 0.6283 * P1Design + 1.791
        vsDesign1500 = 0.000007512 * Pow(P1Design, 5) - 0.0004744 * Pow(P1Design, 4) + 0.01142 * Pow(P1Design, 3) - 0.1432 * Pow(P1Design, 2) + 1.327 * P1Design + 1.739
        Select Case n1Design
            Case 500 To 750
                vsDesign = (vsDesign750 - vsDesign500) / 250 * (n1Design - 500) + vsDesign500
            Case 750 To 1000
                vsDesign = (vsDesign1000 - vsDesign750) / 250 * (n1Design - 750) + vsDesign750
            Case 1000 To 1500
                vsDesign = (vsDesign1500 - vsDesign1000) / 500 * (n1Design - 1000) + vsDesign1000
        End Select
        TextBoxvsDesign.Text = Format(vsDesign, "##.###")
        z1Design = Val(ComboBoxz1Design.Text)
        z2Design = Val(TextBoxz2Design.Text)
        Button1OK.Focus()
    End Sub
    '更新z2
    Private Sub ComboBox1z1Design_TextChanged(sender As Object, e As EventArgs) Handles ComboBoxz1Design.TextChanged
        TextBoxz2Design.Text = Round(Val(ComboBoxz1Design.Text) * iDesign)
    End Sub
    '模块1更改
    Private Sub Button1Change_Click(sender As Object, e As EventArgs) Handles Button1Change.Click
        Panel1_2.Enabled = False
        Panel1_1.Enabled = True
    End Sub
    '模块1确定
    Private Sub Button1OK_Click(sender As Object, e As EventArgs) Handles Button1OK.Click
        TextBoxiDesign.Text = Val(TextBoxz2Design.Text) / Val(ComboBoxz1Design.Text)
        iDesign = Val(TextBoxiDesign.Text)
        TextBoxEfficiencyDesign.Text = Format((100 - 3.5 * Sqrt(iDesign)) * 0.01, "0.###")
        EfficiencyDesign = Val(TextBoxEfficiencyDesign.Text)
        TextBoxT2Design.Text = Format(9550 * P1Design * Val(TextBoxEfficiencyDesign.Text) * iDesign / n1Design, "##########.###")
        T2Design = Val(TextBoxT2Design.Text)
        n2Design = Val(TextBoxn2Design.Text)
        EfficiencyDesign = Val(TextBoxEfficiencyDesign.Text)
        T2Design = Val(TextBoxT2Design.Text)
        vsDesign = Val(TextBoxvsDesign.Text)
        GroupBox2.Enabled = True
        GroupBox1.Enabled = False
        ComboBoxWormGearMaterial.Items.Clear()
        If vsDesign >= 0.25 And vsDesign < 0.5 Then
            ComboBoxWormGearMaterial.Items.Add("ZCuSn10Pb1-砂模")
            ComboBoxWormGearMaterial.Items.Add("ZCuSn10Pb1-金属模")
            ComboBoxWormGearMaterial.Items.Add("ZCuSn5Pb5Zn5-砂模")
            ComboBoxWormGearMaterial.Items.Add("ZCuSn5Pb5Zn5-金属模")
            ComboBoxWormGearMaterial.Items.Add("HT150-砂模")
            ComboBoxWormGearMaterial.Items.Add("HT200-砂模")
        ElseIf vsDesign >= 0.5 And vsDesign <= 2 Then
            ComboBoxWormGearMaterial.Items.Add("ZCuSn10Pb1-砂模")
            ComboBoxWormGearMaterial.Items.Add("ZCuSn10Pb1-金属模")
            ComboBoxWormGearMaterial.Items.Add("ZCuSn5Pb5Zn5-砂模")
            ComboBoxWormGearMaterial.Items.Add("ZCuSn5Pb5Zn5-金属模")
            ComboBoxWormGearMaterial.Items.Add("ZCuAl10Fe3-砂模")
            ComboBoxWormGearMaterial.Items.Add("ZCuAl10Fe3-金属模")
            ComboBoxWormGearMaterial.Items.Add("ZCuAl10Fe3Mn2-金属模")
            ComboBoxWormGearMaterial.Items.Add("ZCuZn38Mn2Pb2-砂模")
            ComboBoxWormGearMaterial.Items.Add("HT150-砂模")
            ComboBoxWormGearMaterial.Items.Add("HT200-砂模")
        ElseIf vsDesign > 2 And vsDesign <= 8 Then
            ComboBoxWormGearMaterial.Items.Add("ZCuSn10Pb1-砂模")
            ComboBoxWormGearMaterial.Items.Add("ZCuSn10Pb1-金属模")
            ComboBoxWormGearMaterial.Items.Add("ZCuSn5Pb5Zn5-砂模")
            ComboBoxWormGearMaterial.Items.Add("ZCuSn5Pb5Zn5-金属模")
            ComboBoxWormGearMaterial.Items.Add("ZCuAl10Fe3-砂模")
            ComboBoxWormGearMaterial.Items.Add("ZCuAl10Fe3-金属模")
            ComboBoxWormGearMaterial.Items.Add("ZCuAl10Fe3Mn2-金属模")
            ComboBoxWormGearMaterial.Items.Add("ZCuZn38Mn2Pb2-砂模")
        End If
        ComboBoxWormGearMaterial.Text = "ZCuSn10Pb1-砂模"
        ComboBoxLubricationMode.Text = "喷油润滑"
        ComboBoxLoadDirection.Text = "一侧受载"
        Button2Calculation.Focus()
    End Sub
    '模块2返回模块1
    Private Sub Button2BackToThePreviousStep_Click(sender As Object, e As EventArgs) Handles Button2BackToThePreviousStep.Click
        GroupBox2.Enabled = False
        GroupBox1.Enabled = True
    End Sub
    '蜗杆材料更新
    Private Sub ComboBoxWormGearMaterial_TextChanged(sender As Object, e As EventArgs) Handles ComboBoxWormGearMaterial.TextChanged
        ComboBoxWormMaterial.Items.Clear()
        Select Case ComboBoxWormGearMaterial.Text
            Case "ZCuSn10Pb1-砂模", "ZCuSn10Pb1-金属模", "ZCuSn5Pb5Zn5-砂模", "ZCuSn5Pb5Zn5-金属模"
                ComboBoxWormMaterial.Items.Add("45-表面淬火")
                ComboBoxWormMaterial.Items.Add("42SiMn-表面淬火")
                ComboBoxWormMaterial.Items.Add("37SiMn2MoV-表面淬火")
                ComboBoxWormMaterial.Items.Add("40Cr-表面淬火")
                ComboBoxWormMaterial.Items.Add("35CrMo-表面淬火")
                ComboBoxWormMaterial.Items.Add("38SiMnMo-表面淬火")
                ComboBoxWormMaterial.Items.Add("42CrMo-表面淬火")
                ComboBoxWormMaterial.Items.Add("40CrNi-表面淬火")
                ComboBoxWormMaterial.Items.Add("15CrMn-渗碳淬火")
                ComboBoxWormMaterial.Items.Add("20CrMn-渗碳淬火")
                ComboBoxWormMaterial.Items.Add("20Cr-渗碳淬火")
                ComboBoxWormMaterial.Items.Add("20CrNi-渗碳淬火")
                ComboBoxWormMaterial.Items.Add("20CrMnTi-渗碳淬火")
                ComboBoxWormMaterial.Items.Add("18Cr2Ni4W-渗碳淬火")
                ComboBoxWormMaterial.Items.Add("45-调质")
                ComboBoxWormMaterial.Text = "45-表面淬火"
            Case "ZCuAl10Fe3-砂模", "ZCuAl10Fe3-金属模", "ZCuAl10Fe3Mn2-金属模", "ZCuZn38Mn2Pb2-砂模"
                ComboBoxWormMaterial.Items.Add("钢未淬火")
                ComboBoxWormMaterial.Items.Add("钢经淬火")
                ComboBoxWormMaterial.Text = "钢未淬火"
            Case "HT150-砂模"
                ComboBoxWormMaterial.Items.Add("渗碳钢")
                ComboBoxWormMaterial.Items.Add("调质或淬火钢")
                ComboBoxWormMaterial.Text = "渗碳钢"
            Case "HT200-砂模"
                ComboBoxWormMaterial.Items.Add("渗碳钢")
                ComboBoxWormMaterial.Text = "渗碳钢"
        End Select
    End Sub
    '模块2计算
    Private Sub Button2Calculation_Click(sender As Object, e As EventArgs) Handles Button2Calculation.Click
        WormGearMaterialDesign = ComboBoxWormGearMaterial.Text
        WormMaterialDesign = ComboBoxWormMaterial.Text
        LubricationModeDesign = ComboBoxLubricationMode.Text
        LoadDriectionDesign = ComboBoxLoadDirection.Text
        Panel2_2.Enabled = True
        Panel2_1.Enabled = False
        Select Case ComboBoxLubricationMode.Text
            Case "喷油润滑"
                ZVSDesign = 0.000000651 * Pow(vsDesign, 6) - 0.00002987 * Pow(vsDesign, 5) + 0.0005001 * Pow(vsDesign, 4) - 0.003735 * Pow(vsDesign, 3) + 0.01337 * Pow(vsDesign, 2) - 0.0316 * vsDesign + 1.022
            Case "浸油润滑"
                ZVSDesign = 0.000001753 * Pow(vsDesign, 5) - 0.00007514 * Pow(vsDesign, 4) + 0.001123 * Pow(vsDesign, 3) - 0.005583 * Pow(vsDesign, 2) - 0.01535 * vsDesign + 1.02
        End Select
        ZVSDesign = Format(ZVSDesign, "0.###")
        NDesign = 60 * n2Design * TDesign
        Select Case LoadDriectionDesign
            Case "一侧受载"
                Select Case WormGearMaterialDesign
                    Case "ZCuSn10Pb1-砂模"
                        SigmaFPSkimDesign = 50
                    Case "ZCuSn10Pb1-金属模"
                        SigmaFPSkimDesign = 70
                    Case "ZCuSn5Pb5Zn5-砂模"
                        SigmaFPSkimDesign = 32
                    Case "ZCuSn5Pb5Zn5-金属模"
                        SigmaFPSkimDesign = 40
                    Case "ZCuAl10Fe3-砂模"
                        SigmaFPSkimDesign = 80
                    Case "ZCuAl10Fe3-金属模"
                        SigmaFPSkimDesign = 90
                    Case "ZCuAl10Fe3Mn2-金属模"
                        SigmaFPSkimDesign = 100
                    Case "ZCuZn38Mn2Pb2-砂模"
                        SigmaFPSkimDesign = 60
                    Case "HT150-砂模"
                        SigmaFPSkimDesign = 40
                    Case "HT200-砂模"
                        SigmaFPSkimDesign = 47
                End Select
            Case "两侧受载"
                Select Case WormGearMaterialDesign
                    Case "ZCuSn10Pb1-砂模"
                        SigmaFPSkimDesign = 30
                    Case "ZCuSn10Pb1-金属模"
                        SigmaFPSkimDesign = 40
                    Case "ZCuSn5Pb5Zn5-砂模"
                        SigmaFPSkimDesign = 24
                    Case "ZCuSn5Pb5Zn5-金属模"
                        SigmaFPSkimDesign = 28
                    Case "ZCuAl10Fe3-砂模"
                        SigmaFPSkimDesign = 63
                    Case "ZCuAl10Fe3-金属模"
                        SigmaFPSkimDesign = 80
                    Case "ZCuAl10Fe3Mn2-金属模"
                        SigmaFPSkimDesign = 90
                    Case "ZCuZn38Mn2Pb2-砂模"
                        SigmaFPSkimDesign = 55
                    Case "HT150-砂模"
                        SigmaFPSkimDesign = 25
                    Case "HT200-砂模"
                        SigmaFPSkimDesign = 30
                End Select
        End Select
        If WormGearMaterialDesign = "HT150-砂模" Or WormGearMaterialDesign = "HT200-砂模" Then
            Select Case NDesign
                Case < 1000000
                    YNDesign = 1
                Case 1000000 To 6000000
                    YNDesign = (0.84 - 1) / (Log10(6000000) - Log10(1000000)) * (Log10(NDesign) - Log10(1000000)) + 1
                Case > 6000000
                    YNDesign = 0.84
            End Select
        Else
            Select Case NDesign
                Case < 1000000
                    YNDesign = 1
                Case 1000000 To 250000000
                    YNDesign = (0.542 - 1) / (Log10(250000000) - Log10(1000000)) * (Log10(NDesign) - Log10(1000000)) + 1
                Case > 250000000
                    YNDesign = 0.542
            End Select
        End If
        SigmaFPDesign = SigmaFPSkimDesign * YNDesign
        TextBoxSigmaFPDesign.Text = Format(SigmaFPDesign, "##########.###")
        SigmaFPDesign = TextBoxSigmaFPDesign.Text
        Select Case WormGearMaterialDesign
            Case "ZCuSn10Pb1-砂模"
                SigmaHPSkimDesign = IIf(WormMaterialDesign = "45-调质", 180, 200)
            Case "ZCuSn10Pb1-金属模"
                SigmaHPSkimDesign = IIf(WormMaterialDesign = "45-调质", 200, 220)
            Case "ZCuSn5Pb5Zn5-砂模"
                SigmaHPSkimDesign = IIf(WormMaterialDesign = "45-调质", 110, 125)
            Case "ZCuSn5Pb5Zn5-金属模"
                SigmaHPSkimDesign = IIf(WormMaterialDesign = "45-调质", 135, 150)
            Case "ZCuAl10Fe3-砂模", "ZCuAl10Fe3-金属模", "ZCuAl10Fe3Mn2-金属模"
                Select Case WormMaterialDesign
                    Case "钢未淬火"
                        Select Case vsDesign
                            Case 0.5 To 1
                                SigmaHPDesign = ((230 - 250) / 0.5 * (vsDesign - 0.5) + 250) * 0.8
                            Case 1 To 2
                                SigmaHPDesign = ((210 - 230) / 1 * (vsDesign - 1) + 230) * 0.8
                            Case 2 To 3
                                SigmaHPDesign = ((180 - 210) / 1 * (vsDesign - 2) + 210) * 0.8
                            Case 3 To 4
                                SigmaHPDesign = ((160 - 180) / 1 * (vsDesign - 3) + 180) * 0.8
                            Case 4 To 6
                                SigmaHPDesign = ((120 - 160) / 2 * (vsDesign - 4) + 160) * 0.8
                            Case 6 To 8
                                SigmaHPDesign = ((90 - 120) / 2 * (vsDesign - 6) + 120) * 0.8
                        End Select
                    Case "钢经淬火"
                        Select Case vsDesign
                            Case 0.5 To 1
                                SigmaHPDesign = (230 - 250) / 0.5 * (vsDesign - 0.5) + 250
                            Case 1 To 2
                                SigmaHPDesign = (210 - 230) / 1 * (vsDesign - 1) + 230
                            Case 2 To 3
                                SigmaHPDesign = (180 - 210) / 1 * (vsDesign - 2) + 210
                            Case 3 To 4
                                SigmaHPDesign = (160 - 180) / 1 * (vsDesign - 3) + 180
                            Case 4 To 6
                                SigmaHPDesign = (120 - 160) / 2 * (vsDesign - 4) + 160
                            Case 6 To 8
                                SigmaHPDesign = (90 - 120) / 2 * (vsDesign - 6) + 120
                        End Select
                End Select
                TextBoxSigmaHPDesign.Text = Format(SigmaHPDesign, "##########.###")
                SigmaHPDesign = Val(TextBoxSigmaHPDesign.Text)
                Exit Sub
            Case "ZCuZn38Mn2Pb2-砂模"
                Select Case WormMaterialDesign
                    Case "钢未淬火"
                        Select Case vsDesign
                            Case 0.5 To 1
                                SigmaHPDesign = ((200 - 215) / 0.5 * (vsDesign - 0.5) + 215) * 0.8
                            Case 1 To 2
                                SigmaHPDesign = ((180 - 200) / 1 * (vsDesign - 1) + 200) * 0.8
                            Case 2 To 3
                                SigmaHPDesign = ((150 - 180) / 1 * (vsDesign - 2) + 180) * 0.8
                            Case 3 To 4
                                SigmaHPDesign = ((135 - 150) / 1 * (vsDesign - 3) + 150) * 0.8
                            Case 4 To 6
                                SigmaHPDesign = ((95 - 135) / 2 * (vsDesign - 4) + 135) * 0.8
                            Case 6 To 8
                                SigmaHPDesign = ((75 - 95) / 2 * (vsDesign - 6) + 95) * 0.8
                        End Select
                    Case "钢经淬火"
                        Select Case vsDesign
                            Case 0.5 To 1
                                SigmaHPDesign = (200 - 215) / 0.5 * (vsDesign - 0.5) + 215
                            Case 1 To 2
                                SigmaHPDesign = (180 - 200) / 1 * (vsDesign - 1) + 200
                            Case 2 To 3
                                SigmaHPDesign = (150 - 180) / 1 * (vsDesign - 2) + 180
                            Case 3 To 4
                                SigmaHPDesign = (135 - 150) / 1 * (vsDesign - 3) + 150
                            Case 4 To 6
                                SigmaHPDesign = (95 - 135) / 2 * (vsDesign - 4) + 135
                            Case 6 To 8
                                SigmaHPDesign = (75 - 95) / 2 * (vsDesign - 6) + 95
                        End Select
                End Select
                TextBoxSigmaHPDesign.Text = Format(SigmaHPDesign, "##########.###")
                SigmaHPDesign = Val(TextBoxSigmaHPDesign.Text)
                Exit Sub
            Case "HT150-砂模", "HT200-砂模"
                Select Case WormMaterialDesign
                    Case "渗碳钢"
                        Select Case vsDesign
                            Case 0.25 To 0.5
                                SigmaHPDesign = (130 - 160) / 0.25 * (vsDesign - 0.25) + 160
                            Case 0.5 To 1
                                SigmaHPDesign = (115 - 130) / 0.5 * (vsDesign - 0.5) + 130
                            Case 1 To 2
                                SigmaHPDesign = (90 - 115) / 1 * (vsDesign - 1) + 115
                        End Select
                    Case "调质或淬火钢"
                        Select Case vsDesign
                            Case 0.25 To 0.5
                                SigmaHPDesign = (110 - 140) / 0.25 * (vsDesign - 0.25) + 140
                            Case 0.5 To 1
                                SigmaHPDesign = (90 - 110) / 0.5 * (vsDesign - 0.5) + 110
                            Case 1 To 2
                                SigmaHPDesign = (70 - 90) / 1 * (vsDesign - 1) + 90
                        End Select
                End Select
                TextBoxSigmaHPDesign.Text = Format(SigmaHPDesign, "##########.###")
                SigmaHPDesign = Val(TextBoxSigmaHPDesign.Text)
                Exit Sub
        End Select
        Select Case NDesign
            Case > 2.5 * Pow(10, 8)
                ZNDesign = 0.67
            Case Else
                ZNDesign = (0.67 - 1.5) / (Log10(250000000) - Log10(300000)) * (Log10(NDesign) - Log10(300000)) + 1.5
        End Select
        SigmaHPDesign = SigmaHPSkimDesign * ZVSDesign * ZNDesign
        TextBoxSigmaHPDesign.Text = Format(SigmaHPDesign, "##########.###")
        SigmaHPDesign = Val(TextBoxSigmaHPDesign.Text)
        Button2OK.Focus()
    End Sub
    '模块2更改
    Private Sub Button2Change_Click(sender As Object, e As EventArgs) Handles Button2Change.Click
        Panel2_2.Enabled = False
        Panel2_1.Enabled = True
    End Sub
    '模块2确定
    Private Sub Button2OK_Click(sender As Object, e As EventArgs) Handles Button2OK.Click
        GroupBox2.Enabled = False
        GroupBox3.Enabled = True
        Button3Calculation.Focus()
    End Sub
    '模块3返回模块2
    Private Sub Button3BackToThePreviousStep_Click(sender As Object, e As EventArgs) Handles Button3BackToThePreviousStep.Click
        GroupBox2.Enabled = True
        GroupBox3.Enabled = False
    End Sub
    '模块3计算
    Private Sub Button3Calculation_Click(sender As Object, e As EventArgs) Handles Button3Calculation.Click
        If (Not IsNumeric(TextBoxKDesign.Text)) Or Val(TextBoxKDesign.Text) < 1 Or Val(TextBoxKDesign.Text) > 1.4 Then
            MsgBox("载荷系数K的格式不正确！", 0 + 48, "格式不正确")
            Exit Sub
        End If
        KDesign = Val(TextBoxKDesign.Text)
        m2d1MIN = Pow((15000 / SigmaHPDesign / z2Design), 2) * KDesign * T2Design
        TextBoxm2d1MIN.Text = Format(m2d1MIN, "##########.###")
        m2d1MIN = Val(TextBoxm2d1MIN.Text)
        ComboBoxmDesign.Items.Clear()
        Dim i%, j#
        objDataSetm2d1MIN1.Clear() '清空数据集
        conStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data source=Worm and Worm Gear.mdb" '字符串链接到数据库
        sqlStr = "select * from m2d1 where m2d1>=" & m2d1MIN '字符串选择符合的值
        objConn.ConnectionString = conStr '链接到数据库
        objAdp = New OleDb.OleDbDataAdapter(sqlStr, objConn) '符合的值给数据适配器
        objAdp.Fill(objDataSetm2d1MIN1, "m2d11") '数据适配器赋值给数据集
        For i = 0 To objDataSetm2d1MIN1.Tables("m2d11").Rows.Count - 1 '数据集内容给Combobox
            j = objDataSetm2d1MIN1.Tables("m2d11").Rows(i).Item(0)
            If Not ComboBoxmDesign.Items.Contains(j) Then ComboBoxmDesign.Items.Add(j)
        Next
        mDesign = objDataSetm2d1MIN1.Tables("m2d11").Rows(0).Item(0) 'Combobox默认为最小的
        ComboBoxmDesign.Text = mDesign
        Panel3_1.Enabled = False
        Panel3_2.Enabled = True
        Button3OK.Focus()
    End Sub
    '更新d1
    Private Sub ComboBoxmDesign_TextChanged(sender As Object, e As EventArgs) Handles ComboBoxmDesign.TextChanged, ComboBoxLodeDesign.TextChanged
        If Panel3_2.Enabled Then mDesign = Val(ComboBoxmDesign.Text)
        ComboBoxd1Design.Items.Clear()
        Dim i%, j#
        objDataSetm2d1MIN2.Clear() '清空数据集
        conStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data source=Worm and Worm Gear.mdb" '字符串链接到数据库
        sqlStr = "select * from m2d1 where m2d1>=" & m2d1MIN & "and m=" & mDesign '字符串选择符合的值
        objConn.ConnectionString = conStr '链接到数据库
        objAdp = New OleDb.OleDbDataAdapter(sqlStr, objConn) '符合的值给数据适配器
        objAdp.Fill(objDataSetm2d1MIN2, "m2d12") '数据适配器赋值给数据集
        For i = 0 To objDataSetm2d1MIN2.Tables("m2d12").Rows.Count - 1 '数据集内容给Combobox
            j = objDataSetm2d1MIN2.Tables("m2d12").Rows(i).Item(1)
            ComboBoxd1Design.Items.Add(j)
        Next
        d1Design = objDataSetm2d1MIN2.Tables("m2d12").Rows(0).Item(1) 'Combobox默认为最小的
        ComboBoxd1Design.Text = d1Design
    End Sub
    '模块3更改
    Private Sub Button3Change_Click(sender As Object, e As EventArgs) Handles Button3Change.Click
        Panel3_1.Enabled = True
        Panel3_2.Enabled = False
    End Sub
    '模块3确定
    Private Sub Button3OK_Click(sender As Object, e As EventArgs) Handles Button3OK.Click
        ComboBoxLodeDesign.Text = "载荷平稳"
        GroupBox3.Enabled = False
        GroupBox4.Enabled = True
        Button4Calculation.Focus()
    End Sub
    '模块4返回模块3
    Private Sub Button4BackToThePreviousStep_Click(sender As Object, e As EventArgs) Handles Button4BackToThePreviousStep.Click
        GroupBox3.Enabled = True
        GroupBox4.Enabled = False
    End Sub
    '模块4计算
    Private Sub Button4Calculation_Click(sender As Object, e As EventArgs) Handles Button4Calculation.Click
        If (Not IsNumeric(TextBoxKADesign.Text)) Or Val(TextBoxKADesign.Text) < 0.8 Or Val(TextBoxKADesign.Text) > 2.25 Then
            MsgBox("KA的值在0.8~2.25内！", 0 + 48, "格式不正确")
            Exit Sub
        End If
        Select Case WormGearMaterialDesign
            Case "ZCuSn10Pb1-砂模", "ZCuSn10Pb1-金属模", "ZCuSn5Pb5Zn5-砂模", "ZCuSn5Pb5Zn5-金属模"
                ZEDesign = 115
            Case "ZCuAl10Fe3-砂模", "ZCuAl10Fe3-金属模", "ZCuAl10Fe3Mn2-砂模", "ZCuAl10Fe3Mn2-金属模", "ZCuZn38Mn2Pb2-砂模", "ZCuZn38Mn2Pb2-金属模"
                ZEDesign = 156
            Case "HT150-砂模", "HT200-砂模"
                ZEDesign = 162
        End Select
        KADesign = Val(TextBoxKADesign.Text)
        d2Design = mDesign * z2Design
        v2Design = PI * d2Design * n2Design / 60000
        KVDesign = IIf(v2Design <= 3, 1.05, 1.15)
        KBetaDesign = IIf(ComboBoxLodeDesign.Text = "载荷平稳", 1, 1.2)
        SigmaHDesign = ZEDesign * Sqrt(9400 * T2Design * KADesign * KVDesign * KBetaDesign / d1Design / d2Design / d2Design)
        TextBoxSigmaH.Text = Format(SigmaHDesign, "##########.###")
        SigmaHDesign = Val(TextBoxSigmaH.Text)
        TextBoxSigmaHPDesign2.Text = TextBoxSigmaHPDesign.Text
        LabelSymbol1.Text = IIf(SigmaHDesign <= SigmaHPDesign, "<=", ">")
        LabelConclusion1.Text = IIf(SigmaHDesign <= SigmaHPDesign, "接触应力符合要求！", "接触应力不符合要求！")
        GammaDesign = Atan(z1Design * mDesign / d1Design)
        YBetaDesign = 1 - GammaDesign / (PI * 2 / 3)
        zvDesign = z2Design / Pow(Cos(GammaDesign), 3)
        YFADesign = -0.000005488 * Pow(zvDesign, 3) + 0.001082 * Pow(zvDesign, 2) - 0.06968 * zvDesign + 3.805
        YSADesign = 0.0000008929 * Pow(zvDesign, 3) - 0.0001893 * Pow(zvDesign, 2) + 0.01477 * zvDesign + 1.323
        YFSDesign = YFADesign * YSADesign
        SigmaFDesign = 666 * T2Design * KADesign * KVDesign * KBetaDesign / d1Design / d2Design / mDesign * YFSDesign * YBetaDesign
        TextBoxSigmaFPDesign2.Text = SigmaFPDesign
        TextBoxSigmaF.Text = Format(SigmaFDesign, "##########.###")
        SigmaFDesign = TextBoxSigmaF.Text
        LabelSymbol2.Text = IIf(SigmaFDesign <= SigmaFPDesign, "<=", ">")
        LabelConclusion2.Text = IIf(SigmaFDesign <= SigmaFPDesign, "弯曲应力符合要求！", "弯曲应力不符合要求！")
        Panel4_1.Enabled = False
        Panel4_2.Enabled = True
        Button4OutputNum.Focus()
    End Sub
    '模块4更改
    Private Sub Button4Change_Click(sender As Object, e As EventArgs) Handles Button4Change.Click
        Panel4_1.Enabled = True
        Panel4_2.Enabled = False
    End Sub
    '导出设计数据
    Private Sub Button4OutputNum_Click(sender As Object, e As EventArgs) Handles Button4OutputNum.Click
        Dim NowTime$ = Format(DateTime.Now, "yyyy年MM月dd日hh时mm分ss秒")
        Dim DesignTXTPath$ = FilePath & "\蜗轮蜗杆设计结果-" & NowTime & ".txt"
        Dim t As System.IO.StreamWriter = New System.IO.StreamWriter(DesignTXTPath, True, System.Text.Encoding.UTF8)
        Dim Digit% = 3
        t.WriteLine("设计结果:")
        t.WriteLine("")
        t.WriteLine("传动比i:          " & iDesign)
        t.WriteLine("蜗杆输入功率:     " & P1Design & "KW")
        t.WriteLine("蜗杆转速:         " & n1Design & "r/min")
        t.WriteLine("使用寿命:         " & TDesign & "h")
        t.WriteLine("")
        t.WriteLine("蜗轮材料及其铸造方法:  " & WormGearMaterialDesign)
        t.WriteLine("蜗杆材料及其热处理方法:" & WormMaterialDesign)
        t.WriteLine("润滑方式:              " & LubricationModeDesign)
        t.WriteLine("齿受载方式:            " & LoadDriectionDesign)
        t.WriteLine("")
        t.WriteLine("模数m:            " & mDesign & "mm")
        t.WriteLine("蜗杆头数:         " & z1Design)
        t.WriteLine("蜗轮齿数:         " & z2Design)
        t.WriteLine("蜗杆分度圆直径d1: " & d1Design & "mm")
        t.WriteLine("蜗轮分度圆直径d2: " & d2Design & "mm")
        t.WriteLine("")
        t.WriteLine("蜗轮转速:         " & n2Design & "r/min")
        t.WriteLine("传动效率:         " & EfficiencyDesign)
        t.WriteLine("蜗轮输出转矩:     " & T2Design & "N*m")
        t.WriteLine("滑动速度:         " & vsDesign & "m/s")
        t.WriteLine("")
        t.WriteLine("许用接触应力: " & SigmaHPDesign & "MPa")
        t.WriteLine("接触应力:     " & SigmaHDesign & "MPa")
        t.WriteLine("许用弯曲应力: " & SigmaFPDesign & "MPa")
        t.WriteLine("弯曲应力:     " & SigmaFDesign & "MPa")
        t.WriteLine("")
        t.WriteLine("参数生成时间：" & NowTime)
        t.Close()
        Shell("notepad.exe " & DesignTXTPath, vbNormalFocus)
    End Sub
    '准备建模
    Private Sub ButtonPrepareModeling_Click(sender As Object, e As EventArgs) Handles ButtonPrepareModeling.Click
        ComboBoxm.Text = mDesign
        ComboBoxz1.Text = z1Design
        ComboBoxz2.Text = z2Design
        ComboBoxd1.Text = d1Design
        MsgBox("数据导入成功！", 0 + 64, "成功")
        TabControl1.SelectedIndex = 0
    End Sub
    '切换到Page1、2时
    Private Sub TabPageCreatePartAndAssem_Enter(sender As Object, e As EventArgs) Handles TabPageCreatePart.Enter, TabPageCreateAssem.Enter
        PanelUP.Enabled = True
    End Sub
    '切换到Page3时
    Private Sub TabPageDesign_Enter(sender As Object, e As EventArgs) Handles TabPageDesign.Enter
        PanelUP.Enabled = False
        Button1Calculation.Focus()
    End Sub




End Class

