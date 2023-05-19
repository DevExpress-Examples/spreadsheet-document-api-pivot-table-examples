Namespace SpreadsheetDocServerPivotAPI
    Partial Public Class Form1
        ''' <summary>
        ''' Required designer variable.
        ''' </summary>
        Private components As System.ComponentModel.IContainer = Nothing

        ''' <summary>
        ''' Clean up any resources being used.
        ''' </summary>
        ''' <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        Protected Overrides Sub Dispose(ByVal disposing As Boolean)
            If disposing AndAlso (components IsNot Nothing) Then
                components.Dispose()
            End If
            MyBase.Dispose(disposing)
        End Sub

        #Region "Windows Form Designer generated code"

        ''' <summary>
        ''' Required method for Designer support - do not modify
        ''' the contents of this method with the code editor.
        ''' </summary>
        Private Sub InitializeComponent()
            Me.components = New System.ComponentModel.Container()
            Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
            Me.xtraTabControl1 = New DevExpress.XtraTab.XtraTabControl()
            Me.xtraTabPage1 = New DevExpress.XtraTab.XtraTabPage()
            Me.richEditControlCS = New DevExpress.XtraRichEdit.RichEditControl()
            Me.xtraTabPage2 = New DevExpress.XtraTab.XtraTabPage()
            Me.richEditControlVB = New DevExpress.XtraRichEdit.RichEditControl()
            Me.btnOpenExcel = New DevExpress.XtraEditors.SimpleButton()
            Me.treeList1 = New DevExpress.XtraTreeList.TreeList()
            Me.defaultLookAndFeel1 = New DevExpress.LookAndFeel.DefaultLookAndFeel(Me.components)
            Me.LayoutControl1 = New DevExpress.XtraLayout.LayoutControl()
            Me.Root = New DevExpress.XtraLayout.LayoutControlGroup()
            Me.codeExampleNameLbl = New DevExpress.XtraLayout.SimpleLabelItem()
            Me.LayoutControlItem1 = New DevExpress.XtraLayout.LayoutControlItem()
            Me.LayoutControlItem2 = New DevExpress.XtraLayout.LayoutControlItem()
            Me.LayoutControlItem3 = New DevExpress.XtraLayout.LayoutControlItem()
            Me.SplitterItem1 = New DevExpress.XtraLayout.SplitterItem()
            CType(Me.xtraTabControl1, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.xtraTabControl1.SuspendLayout()
            Me.xtraTabPage1.SuspendLayout()
            Me.xtraTabPage2.SuspendLayout()
            CType(Me.treeList1, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.LayoutControl1, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.LayoutControl1.SuspendLayout()
            CType(Me.Root, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.codeExampleNameLbl, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.LayoutControlItem1, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.LayoutControlItem2, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.LayoutControlItem3, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.SplitterItem1, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.SuspendLayout()
            '
            'xtraTabControl1
            '
            Me.xtraTabControl1.AppearancePage.PageClient.BackColor = System.Drawing.Color.Transparent
            Me.xtraTabControl1.AppearancePage.PageClient.BackColor2 = System.Drawing.Color.Transparent
            Me.xtraTabControl1.AppearancePage.PageClient.BorderColor = System.Drawing.Color.Transparent
            Me.xtraTabControl1.AppearancePage.PageClient.Options.UseBackColor = True
            Me.xtraTabControl1.AppearancePage.PageClient.Options.UseBorderColor = True
            Me.xtraTabControl1.HeaderAutoFill = DevExpress.Utils.DefaultBoolean.[True]
            Me.xtraTabControl1.Location = New System.Drawing.Point(12, 53)
            Me.xtraTabControl1.Name = "xtraTabControl1"
            Me.xtraTabControl1.SelectedTabPage = Me.xtraTabPage1
            Me.xtraTabControl1.Size = New System.Drawing.Size(898, 520)
            Me.xtraTabControl1.TabIndex = 1
            Me.xtraTabControl1.TabPages.AddRange(New DevExpress.XtraTab.XtraTabPage() {Me.xtraTabPage1, Me.xtraTabPage2})
            '
            'xtraTabPage1
            '
            Me.xtraTabPage1.Appearance.HeaderActive.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold)
            Me.xtraTabPage1.Appearance.HeaderActive.Options.UseFont = True
            Me.xtraTabPage1.Controls.Add(Me.richEditControlCS)
            Me.xtraTabPage1.Name = "xtraTabPage1"
            Me.xtraTabPage1.Size = New System.Drawing.Size(896, 495)
            Me.xtraTabPage1.Text = "C#"
            '
            'richEditControlCS
            '
            Me.richEditControlCS.ActiveViewType = DevExpress.XtraRichEdit.RichEditViewType.Draft
            Me.richEditControlCS.Dock = System.Windows.Forms.DockStyle.Fill
            Me.richEditControlCS.LayoutUnit = DevExpress.XtraRichEdit.DocumentLayoutUnit.Pixel
            Me.richEditControlCS.Location = New System.Drawing.Point(0, 0)
            Me.richEditControlCS.Name = "richEditControlCS"
            Me.richEditControlCS.Options.Annotations.ShowAllAuthors = False
            Me.richEditControlCS.Options.HorizontalRuler.Visibility = DevExpress.XtraRichEdit.RichEditRulerVisibility.Hidden
            Me.richEditControlCS.Size = New System.Drawing.Size(896, 495)
            Me.richEditControlCS.TabIndex = 0
            '
            'xtraTabPage2
            '
            Me.xtraTabPage2.Appearance.HeaderActive.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold)
            Me.xtraTabPage2.Appearance.HeaderActive.Options.UseFont = True
            Me.xtraTabPage2.Controls.Add(Me.richEditControlVB)
            Me.xtraTabPage2.Name = "xtraTabPage2"
            Me.xtraTabPage2.Size = New System.Drawing.Size(947, 510)
            Me.xtraTabPage2.Text = "VB"
            '
            'richEditControlVB
            '
            Me.richEditControlVB.ActiveViewType = DevExpress.XtraRichEdit.RichEditViewType.Draft
            Me.richEditControlVB.Dock = System.Windows.Forms.DockStyle.Fill
            Me.richEditControlVB.LayoutUnit = DevExpress.XtraRichEdit.DocumentLayoutUnit.Pixel
            Me.richEditControlVB.Location = New System.Drawing.Point(0, 0)
            Me.richEditControlVB.Name = "richEditControlVB"
            Me.richEditControlVB.Options.Annotations.ShowAllAuthors = False
            Me.richEditControlVB.Options.HorizontalRuler.Visibility = DevExpress.XtraRichEdit.RichEditRulerVisibility.Hidden
            Me.richEditControlVB.Size = New System.Drawing.Size(947, 510)
            Me.richEditControlVB.TabIndex = 0
            '
            'btnOpenExcel
            '
            Me.btnOpenExcel.ImageOptions.Image = CType(resources.GetObject("btnOpenExcel.ImageOptions.Image"), System.Drawing.Image)
            Me.btnOpenExcel.Location = New System.Drawing.Point(20, 585)
            Me.btnOpenExcel.MaximumSize = New System.Drawing.Size(200, 0)
            Me.btnOpenExcel.Name = "btnOpenExcel"
            Me.btnOpenExcel.Size = New System.Drawing.Size(200, 50)
            Me.btnOpenExcel.StyleController = Me.LayoutControl1
            Me.btnOpenExcel.TabIndex = 0
            Me.btnOpenExcel.Text = "Open in Microsoft Excel"
            '
            'treeList1
            '
            Me.treeList1.Location = New System.Drawing.Point(926, 12)
            Me.treeList1.Name = "treeList1"
            Me.treeList1.Size = New System.Drawing.Size(274, 631)
            Me.treeList1.TabIndex = 0
            '
            'defaultLookAndFeel1
            '
            Me.defaultLookAndFeel1.LookAndFeel.SkinName = "Office 2016 Colorful"
            '
            'LayoutControl1
            '
            Me.LayoutControl1.Controls.Add(Me.treeList1)
            Me.LayoutControl1.Controls.Add(Me.btnOpenExcel)
            Me.LayoutControl1.Controls.Add(Me.xtraTabControl1)
            Me.LayoutControl1.Dock = System.Windows.Forms.DockStyle.Fill
            Me.LayoutControl1.Location = New System.Drawing.Point(0, 0)
            Me.LayoutControl1.Name = "LayoutControl1"
            Me.LayoutControl1.OptionsCustomizationForm.DesignTimeCustomizationFormPositionAndSize = New System.Drawing.Rectangle(617, 63, 650, 400)
            Me.LayoutControl1.Root = Me.Root
            Me.LayoutControl1.Size = New System.Drawing.Size(1212, 655)
            Me.LayoutControl1.TabIndex = 1
            Me.LayoutControl1.Text = "LayoutControl1"
            '
            'Root
            '
            Me.Root.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.[True]
            Me.Root.GroupBordersVisible = False
            Me.Root.Items.AddRange(New DevExpress.XtraLayout.BaseLayoutItem() {Me.codeExampleNameLbl, Me.LayoutControlItem1, Me.LayoutControlItem2, Me.LayoutControlItem3, Me.SplitterItem1})
            Me.Root.Name = "Root"
            Me.Root.Size = New System.Drawing.Size(1212, 655)
            Me.Root.TextVisible = False
            '
            'codeExampleNameLbl
            '
            Me.codeExampleNameLbl.AllowHotTrack = False
            Me.codeExampleNameLbl.AppearanceItemCaption.Font = New System.Drawing.Font("Segoe UI", 20.25!)
            Me.codeExampleNameLbl.AppearanceItemCaption.Options.UseFont = True
            Me.codeExampleNameLbl.Location = New System.Drawing.Point(0, 0)
            Me.codeExampleNameLbl.Name = "codeExampleNameLbl"
            Me.codeExampleNameLbl.Size = New System.Drawing.Size(902, 41)
            Me.codeExampleNameLbl.TextSize = New System.Drawing.Size(329, 37)
            '
            'LayoutControlItem1
            '
            Me.LayoutControlItem1.Control = Me.xtraTabControl1
            Me.LayoutControlItem1.Location = New System.Drawing.Point(0, 41)
            Me.LayoutControlItem1.Name = "LayoutControlItem1"
            Me.LayoutControlItem1.Size = New System.Drawing.Size(902, 524)
            Me.LayoutControlItem1.TextSize = New System.Drawing.Size(0, 0)
            Me.LayoutControlItem1.TextVisible = False
            '
            'LayoutControlItem2
            '
            Me.LayoutControlItem2.Control = Me.btnOpenExcel
            Me.LayoutControlItem2.Location = New System.Drawing.Point(0, 565)
            Me.LayoutControlItem2.MaxSize = New System.Drawing.Size(0, 70)
            Me.LayoutControlItem2.MinSize = New System.Drawing.Size(174, 70)
            Me.LayoutControlItem2.Name = "LayoutControlItem2"
            Me.LayoutControlItem2.Padding = New DevExpress.XtraLayout.Utils.Padding(10, 10, 10, 10)
            Me.LayoutControlItem2.Size = New System.Drawing.Size(902, 70)
            Me.LayoutControlItem2.SizeConstraintsType = DevExpress.XtraLayout.SizeConstraintsType.Custom
            Me.LayoutControlItem2.TextSize = New System.Drawing.Size(0, 0)
            Me.LayoutControlItem2.TextVisible = False
            '
            'LayoutControlItem3
            '
            Me.LayoutControlItem3.Control = Me.treeList1
            Me.LayoutControlItem3.Location = New System.Drawing.Point(914, 0)
            Me.LayoutControlItem3.Name = "LayoutControlItem3"
            Me.LayoutControlItem3.Size = New System.Drawing.Size(278, 635)
            Me.LayoutControlItem3.TextSize = New System.Drawing.Size(0, 0)
            Me.LayoutControlItem3.TextVisible = False
            '
            'SplitterItem1
            '
            Me.SplitterItem1.AllowHotTrack = True
            Me.SplitterItem1.Location = New System.Drawing.Point(902, 0)
            Me.SplitterItem1.Name = "SplitterItem1"
            Me.SplitterItem1.Size = New System.Drawing.Size(12, 635)
            '
            'Form1
            '
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.ClientSize = New System.Drawing.Size(1212, 655)
            Me.Controls.Add(Me.LayoutControl1)
            Me.Name = "Form1"
            Me.Text = "Pivot Table API"
            CType(Me.xtraTabControl1, System.ComponentModel.ISupportInitialize).EndInit()
            Me.xtraTabControl1.ResumeLayout(False)
            Me.xtraTabPage1.ResumeLayout(False)
            Me.xtraTabPage2.ResumeLayout(False)
            CType(Me.treeList1, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.LayoutControl1, System.ComponentModel.ISupportInitialize).EndInit()
            Me.LayoutControl1.ResumeLayout(False)
            CType(Me.Root, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.codeExampleNameLbl, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.LayoutControlItem1, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.LayoutControlItem2, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.LayoutControlItem3, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.SplitterItem1, System.ComponentModel.ISupportInitialize).EndInit()
            Me.ResumeLayout(False)

        End Sub

#End Region
        Private treeList1 As DevExpress.XtraTreeList.TreeList
        Private xtraTabControl1 As DevExpress.XtraTab.XtraTabControl
        Private xtraTabPage1 As DevExpress.XtraTab.XtraTabPage
        Private xtraTabPage2 As DevExpress.XtraTab.XtraTabPage
        Private richEditControlCS As DevExpress.XtraRichEdit.RichEditControl
        Private richEditControlVB As DevExpress.XtraRichEdit.RichEditControl
        Private WithEvents btnOpenExcel As DevExpress.XtraEditors.SimpleButton
        Private defaultLookAndFeel1 As DevExpress.LookAndFeel.DefaultLookAndFeel
        Friend WithEvents LayoutControl1 As DevExpress.XtraLayout.LayoutControl
        Friend WithEvents Root As DevExpress.XtraLayout.LayoutControlGroup
        Friend WithEvents codeExampleNameLbl As DevExpress.XtraLayout.SimpleLabelItem
        Friend WithEvents LayoutControlItem1 As DevExpress.XtraLayout.LayoutControlItem
        Friend WithEvents LayoutControlItem2 As DevExpress.XtraLayout.LayoutControlItem
        Friend WithEvents LayoutControlItem3 As DevExpress.XtraLayout.LayoutControlItem
        Friend WithEvents SplitterItem1 As DevExpress.XtraLayout.SplitterItem
    End Class
End Namespace

