' WINFBE FORM
' WINFBE VERSION 2.1.7
' LOCKCONTROLS=False
' SNAPLINES=True
' WINFBE FORM_START
' WINFBE CONTROL_START Form
'   PROPERTIES_START
'     PROP_NAME=Name
'     PROP_VALUE=Misc
'     PROP_NAME=Left
'     PROP_VALUE=10
'     PROP_NAME=Top
'     PROP_VALUE=10
'     PROP_NAME=Width
'     PROP_VALUE=699
'     PROP_NAME=Height
'     PROP_VALUE=420
'     PROP_NAME=ChildForm
'     PROP_VALUE=True
'     PROP_NAME=Text
'     PROP_VALUE=Form1
'     PROP_NAME=WindowState
'     PROP_VALUE=FormWindowState.Normal
'     PROP_NAME=StartPosition
'     PROP_VALUE=FormStartPosition.Manual
'     PROP_NAME=BorderStyle
'     PROP_VALUE=FormBorderStyle.None
'     PROP_NAME=MinimizeBox
'     PROP_VALUE=True
'     PROP_NAME=MaximizeBox
'     PROP_VALUE=True
'     PROP_NAME=ControlBox
'     PROP_VALUE=True
'     PROP_NAME=Enabled
'     PROP_VALUE=True
'     PROP_NAME=Visible
'     PROP_VALUE=True
'     PROP_NAME=BackColor
'     PROP_VALUE=SYSTEM|Control
'     PROP_NAME=AcceptButton
'     PROP_VALUE=
'     PROP_NAME=AllowDrop
'     PROP_VALUE=False
'     PROP_NAME=KeyPreview
'     PROP_VALUE=False
'     PROP_NAME=CancelButton
'     PROP_VALUE=
'     PROP_NAME=Icon
'     PROP_VALUE=
'     PROP_NAME=Locked
'     PROP_VALUE=False
'     PROP_NAME=MaximumHeight
'     PROP_VALUE=0
'     PROP_NAME=MaximumWidth
'     PROP_VALUE=0
'     PROP_NAME=MinimumHeight
'     PROP_VALUE=0
'     PROP_NAME=MinimumWidth
'     PROP_VALUE=0
'     PROP_NAME=Tag
'     PROP_VALUE=
'   PROPERTIES_END
'   EVENTS_START
'   EVENTS_END
' WINFBE CONTROL_END
' WINFBE CONTROL_START ListView
'   PROPERTIES_START
'     PROP_NAME=Name
'     PROP_VALUE=MiscListView
'     PROP_NAME=Left
'     PROP_VALUE=6
'     PROP_NAME=Top
'     PROP_VALUE=7
'     PROP_NAME=Width
'     PROP_VALUE=688
'     PROP_NAME=Height
'     PROP_VALUE=407
'     PROP_NAME=AllowDrop
'     PROP_VALUE=False
'     PROP_NAME=BackColor
'     PROP_VALUE=SYSTEM|Window
'     PROP_NAME=BorderStyle
'     PROP_VALUE=ControlBorderStyle.Fixed3D
'     PROP_NAME=CheckBoxes
'     PROP_VALUE=False
'     PROP_NAME=Enabled
'     PROP_VALUE=True
'     PROP_NAME=Font
'     PROP_VALUE=Segoe UI,9,400,0,0,0,1
'     PROP_NAME=ForeColor
'     PROP_VALUE=SYSTEM|WindowText
'     PROP_NAME=FullRowSelect
'     PROP_VALUE=True
'     PROP_NAME=GridLines
'     PROP_VALUE=True
'     PROP_NAME=HeaderStyle
'     PROP_VALUE=ColumnHeaderStyle.NonClickable
'     PROP_NAME=HeaderHeight
'     PROP_VALUE=20
'     PROP_NAME=HeaderForeColor
'     PROP_VALUE=SYSTEM|WindowText
'     PROP_NAME=HeaderBackColor
'     PROP_VALUE=SYSTEM|Control
'     PROP_NAME=HideSelection
'     PROP_VALUE=False
'     PROP_NAME=Locked
'     PROP_VALUE=False
'     PROP_NAME=MultiSelect
'     PROP_VALUE=False
'     PROP_NAME=OddRowColor
'     PROP_VALUE=SYSTEM|Window
'     PROP_NAME=OddRowColorEnabled
'     PROP_VALUE=False
'     PROP_NAME=TabIndex
'     PROP_VALUE=1
'     PROP_NAME=TabStop
'     PROP_VALUE=True
'     PROP_NAME=Tag
'     PROP_VALUE=
'     PROP_NAME=ToolTip
'     PROP_VALUE=
'     PROP_NAME=ToolTipBalloon
'     PROP_VALUE=False
'     PROP_NAME=Visible
'     PROP_VALUE=True
'   PROPERTIES_END
'   EVENTS_START
'   EVENTS_END
' WINFBE CONTROL_END
' WINFBE FORM_END
' WINFBE_CODEGEN_START
#if 0
type MiscType extends wfxForm
    private:
        temp as byte
    public:
        declare static function FormInitializeComponent( byval pForm as MiscType ptr ) as LRESULT
        declare constructor
        ' Controls
        MiscListView As wfxListView
end type


function MiscType.FormInitializeComponent( byval pForm as MiscType ptr ) as LRESULT
    dim as long nClientOffset

    pForm->Name = "Misc"
    pForm->ChildForm = True
    pForm->Text = "Form1"
    pForm->BorderStyle = FormBorderStyle.None
    pForm->SetBounds(10,10,699,420)
    pForm->MiscListView.Parent = pForm
    pForm->MiscListView.Name = "MiscListView"
    pForm->MiscListView.GridLines = True
    pForm->MiscListView.HeaderStyle = ColumnHeaderStyle.NonClickable
    pForm->MiscListView.HeaderBackColor = Colors.SystemControl
    pForm->MiscListView.SetBounds(6,7-nClientOffset,688,407)
    pForm->Controls.Add(ControlType.ListView, @(pForm->MiscListView))
    Application.Forms.Add(ControlType.Form, pForm)
    function = 0
end function

constructor MiscType
    InitializeComponent = cast( any ptr, @FormInitializeComponent )
    this.FormInitializeComponent( @this )
end constructor

dim shared Misc as MiscType
#endif
' WINFBE_CODEGEN_END
''
''  Remove the following Application.Run code if it used elsewhere in your application.
'Application.Run(Misc)

