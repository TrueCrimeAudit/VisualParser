#Requires AutoHotkey v2.1-alpha.14
#SingleInstance Force
#Include <GuiEnhancerKit>

#DllLoad "msftedit.dll"

!Esc:: ExitApp()

TP()

class TP {
    static GuiTitle := "!!!Parser.ahk"
    static Margin := 10
    static xPad := 15
    static yPad := 15
    static ControlHeight := 30
    static ControlSpacing := 5
    static inputHeight := 125
    static EditHeight := 150
    static SmallWidth := 100
    static MediumWidth := 150
    static LargeWidth := 300
    static Opt := " cF8F8F8 Background1F1F1F "
    static fWhite := " cF8F8F8 "
    static bBlack := " Background1F1F1F "

    static ParserWidth := 550
    static ParserHeight := 380
    static pWidth := "w" . this.ParserWidth
    static pHeight := "h" . this.ParserHeight

    static Row1 := this.yPad
    static Row2 := (this.yPad)*2 + this.InputHeight
    static Row3 := (this.yPad)*4 + this.InputHeight + this.ControlHeight
    static Row4 := (this.yPad)*6 + this.InputHeight + this.ControlHeight*2
    static Row5 := (this.yPad)*8 + this.InputHeight + this.ControlHeight*3
    static Row6 := 300

    static Column1 := this.xPad
    static Column2 := this.Column1 + this.SmallWidth + this.ControlSpacing
    static Column3 := this.Column2 + this.MediumWidth + this.ControlSpacing  
    static Column4 := this.Column3 + this.MediumWidth + this.ControlSpacing
    static Column5 := this.Column4 + this.SmallWidth + this.ControlSpacing

    static DelimiterMap := Map(
        "CSV", [",", "Comma"],
        "TSV", ["`t", "Tab"],
        "DSV", ["|", "Pipe"],
        "SSV", [" ", "Space"],
        "Custom", ["", "Custom"]
    )

    static ParsersMap := Map(
        "Array", ["Chunk", "Compact", "DepthOf", "Difference", "Drop", "DropRight", "Fill", "Flatten", "FlattenDeep", "FlattenDepth", "FromPairs", "Head", "IndexOf", "Initial", "Intersection", "Join", "Last", "LastIndexOf", "Nth", "Reverse", "Slice", "SortedIndex", "SortedUniq", "Tail", "Take", "TakeRight", "Union", "Uniq", "Unzip", "Without", "Zip", "ZipObject"],
        "Collection", ["Includes", "Map", "Sample", "Shuffle", "Size", "Some"],
        "Lang", ["Clone", "CloneDeep", "IsAlnum", "IsArray", "IsBoolean", "IsBuffer", "IsEmpty", "IsEqual", "IsFloat", "IsFunction", "IsInteger", "IsMap", "IsMatch", "IsNumber", "IsObject", "IsString", "IsUndefined", "ToString", "TypeOf"],
        "Math", ["Add", "Ceil", "Divide", "Floor", "Max", "Mean", "Min", "Multiply", "Round", "Subtract", "Sum"],
        "Number", ["Clamp", "InRange", "Random"],
        "String", ["EndsWith", "Escape", "LowerFirst", "Pad", "PadEnd", "PadStart", "ParseInt", "Repeat", "StartsWith", "ToLower", "ToUpper", "Trim", "TrimEnd", "TrimStart", "Truncate", "Unescape", "UpperFirst", "Words"],
        "Utility", ["Identity"]
    )

    static TestMap := Map(
        "Multi", "hello world, test string, 123, abc",
        "Names", "Neo, Morpheus, Trinity, Agent Smith",
        "Numbers", "10, 20, 30, 40, 50, 60",
        "Dashed", "some-words-with-dashes, more_words_with_underscores",
        "Tabs", "Initiative	Strategy	Initiative	Strategy"
    )

    __New() {
        this.CreateGui()
        this.setupHotkeys()
        this.gui.OnEvent("Size", this.GuiResize.Bind(this))
        this.gui.Show(TP.pHeight TP.pWidth)
    }

    static __New() {
        this.WM_SETTEXT := 0x000C
        this.EM_SETTEXTEX := 0x0461
        this.WM_CTLCOLORSCROLLBAR := 0x0137
    }

    CreateGui() {
        TP.gui := Gui("+Resize", "!!!Parser.ahk")
        this.gui := TP.gui
        this.gui.BackColor := "1F1F1F"
        SetDarkMode(this.gui)        
        this.gui.Opt("+Resize")
        this.gui.SetFont("s10 cWhite", "Segoe UI",)
        this.gui.SetDarkTitle()
        this.gui.SetDarkMenu()

    
        this.inputEdit := RichEdit.Create(this.gui, 
            Format("vInputText h{1} x{2} w{3}",
                TP.InputHeight,
                TP.Column1,
                TP.ParserWidth - (TP.xPad * 2)))
        this.inputEdit.Move(, TP.Row1)
        RichEdit.SetText(this.inputEdit, TP.TestMap["Multi"])
        this.inputEdit.Opt(TP.Opt)
        
        this.outputEdit := RichEdit.Create(this.gui, 
            Format("vOutputText h{1} x{2} w{3}",
                TP.EditHeight,
                TP.Column1,
                TP.ParserWidth - (TP.xPad * 2)))
        this.outputEdit.Move(, TP.Row3)
        RichEdit.SetText(this.outputEdit, "Tests.")
        this.outputEdit.Opt(TP.Opt)

        ; Labels
        ; this.gui.AddText("vModeLabel x" TP.Column1 " y" TP.Row2 + 2, "Mode: ")
        ; this.gui.AddText("vCategoryLabel x" TP.Column2 " y" TP.Row2 + 2, "Category: ")
        ; this.gui.AddText("x" TP.Column3 " y" TP.Row2 + 2, "Method: ")
        ; this.gui.AddText("x" TP.Column4 " y" TP.Row2 + 2, "Type: ")
        ; this.gui.AddText("x" TP.Column5 " y" TP.Row2 + 2, "Delimiter: ")
        

        this.modeDDL := this.gui.AddDropDownList(TP.Opt "vMode Choose1 w100 x" TP.Column1 " y" TP.Row2, ["Table", "Parsers"])
        this.methDDL := this.gui.AddDropDownList(TP.Opt "vMthdSel Choose1 w150 x" TP.Column2 " y" TP.Row2, ["Type"])
        this.typeDDL := this.gui.AddDropDownList(TP.Opt "vParseType Choose1 w150 x" TP.Column3 " y" TP.Row2, ["CSV", "TSV", "DSV", "SSV", "Custom"])
        
        this.modeDDL.OnEvent("Change", this.UpdateMethodList.Bind(this))

        categories := []
        for category in TP.ParsersMap
            categories.Push(category)
        this.categoryDDL := this.gui.Add("DropDownList", 
            TP.Opt "vCategory w150 Hidden x" TP.Column2 " y" TP.Row2, 
            categories)
        
        this.highlightBtn := this.gui.AddButton(TP.Opt "vHighlight w120 x" TP.Column5 + 60 " y" TP.Row2, "Highlight")
        this.typeDDL.OnEvent("Change", this.UpdateDelimiter.Bind(this))
    }

    static Create(gui, options) {
        static WM_VSCROLL := 0x0115
        static COLOR_SCROLLBAR := 0
        static COLOR_BTNFACE := 15
        
        control := gui.AddCustom("ClassRichEdit50W +0x5031b1c4 +E0x20000 " options)
        SendMessage(0x0443, 0, 0x202020, control)
        DllCall("uxtheme\SetWindowTheme", "Ptr", control.hwnd, "Str", "DarkMode_Explorer", "Ptr", 0)
        static WM_SYSCOLORCHANGE := 0x0015
        control.OnMessage(WM_SYSCOLORCHANGE, (ctrl, *) => (
            DllCall("SetSysColors", "Int", 1, "Int*", COLOR_SCROLLBAR, "UInt*", 0x383838),
            DllCall("SetSysColors", "Int", 1, "Int*", COLOR_BTNFACE, "UInt*", 0x383838)
        ))
        PostMessage(WM_SYSCOLORCHANGE, 0, 0, control)
        return control
    }

    UpdateMethodList(*) {
        mode := this.gui["Mode"].Text
        if (mode = "Table") {
            ; Show Table mode controls
            this.methDDL.Opt("Hidden")
            this.typeDDL.Delete()
            this.typeDDL.Add(["CSV", "TSV", "DSV", "SSV", "Custom"])
            this.typeDDL.Choose(1)
            this.typeDDL.Opt("-Hidden")
        } else {
            ; Show Parser mode controls
            this.typeDDL.Opt("Hidden")
            categories := ["Type"]
            for category in TP.ParsersMap
                categories.Push(category)
            
            this.methDDL.Delete()
            this.methDDL.Add(categories)
            this.methDDL.Choose(1)
            this.methDDL.Opt("-Hidden")
            this.methDDL.OnEvent("Change", this.UpdateMethodSelection.Bind(this))
        }
    }
    
    UpdateMethodSelection(*) {
        selectedCategory := this.methDDL.Text
        if (selectedCategory = "Type")
            return
            
        if selectedCategory && TP.ParsersMap.Has(selectedCategory) {
            methods := TP.ParsersMap[selectedCategory]
            this.typeDDL.Delete()
            this.typeDDL.Add(methods)
            this.typeDDL.Choose(1)
            this.typeDDL.Opt("-Hidden")
        }
    }

    GuiResize(thisGui, MinMax, Width, Height) {
        if MinMax = -1
            return
            
        inputHeight := Height * 0.3  ; Adjust ratio as needed
        outputHeight := Height * 0.4  ; Adjust ratio as needed
        contentWidth := Width - (TP.xPad * 2)
        
        if thisGui.HasProp("InputText") && thisGui.HasProp("OutputText") {
            thisGui["InputText"].Move(TP.xPad, TP.yPad, contentWidth, inputHeight)
            thisGui["OutputText"].Move(TP.xPad, inputHeight + TP.yPad * 2, contentWidth, outputHeight)
        }
    }

    UpdateDelimiter(*) {
        parseType := this.gui["ParseType"].Text
        this.gui["Delimiter"].Value := TP.DelimiterMap[parseType][1]
        parseType := this.gui["ParseType"].Text
        if !TP.DelimiterMap.Has(parseType)
            return

        if parseType = "Custom"
            this.gui["Delimiter"].Value := TP.DelimiterMap[parseType][1]
        else
            this.gui["Delimiter"].Opt(parseType = "Custom" ? "-ReadOnly" : "+ReadOnly")
    }

    ProcessHandler(*) {
        try {
            text := RichEdit.GetText(this.gui["InputText"])
            mode := this.gui["Mode"].Text
            
            if (mode = "Table") {
                this.ProcessTableMode(text)
            } else {
                this.ProcessParserMode(text)
            }
        } catch Error as err {
            MsgBox "Error processing text: " err.Message, "Error", "Icon!"
        }
    }

    ProcessTableMode(text) {
        removeHeader := this.gui["RemoveHeader"].Value
        trimLines := this.gui["TrimLines"].Value
        inputDelimiter := this.gui["Delimiter"].Value

        if text = ""
            return

        lines := StrSplit(RegExReplace(text, "^\s+|\s+$"), "`n", "`r")
        if removeHeader && lines.Length
            lines.RemoveAt(1)

        result := ["<table>"]
        for line in lines {
            if line = "" || (trimLines && Trim(line) = "")
                continue

            if trimLines
                line := Trim(line)

            fields := StrSplit(line, inputDelimiter)
            if trimLines
                fields := fields.Map(field => Trim(field))

            row := ["<tr>"]
            for field in fields
                row.Push("<td>" field "</td>")
            row.Push("</tr>")
            result.Push(row.Join(""))
        }
        result.Push("</table>")

        finalOutput := result.Join("`n")
        RichEdit.SetText(this.gui["OutputText"], finalOutput)
        A_Clipboard := finalOutput
    }

    HighlightDelimiters(*) {
        text := RichEdit.GetText(this.gui["InputText"])
        delimiter := this.gui["Delimiter"].Value
        
        if (text = "" || delimiter = "")
            return

        displayDelimiter := delimiter = "`t" ? "→" : delimiter
        RichEdit.SetRTF(this.gui["InputText"], RichEdit.FormatText(text, delimiter))
        
        ToolTip("Highlighted delimiter: " displayDelimiter)
        SetTimer () => ToolTip(), -1000
    }

    ProcessParserMode(text) {
        category := this.gui["Category"].Text
        method := this.gui["Method"].Text
        result := "Processing with " category "." method "`nInput: " text
        
        RichEdit.SetText(this.gui["OutputText"], result)
        A_Clipboard := result
    }
    setupHotkeys() {
        if !this.gui
            return
        HotIfWinActive("ahk_id " this.gui.Hwnd)
        Hotkey("^w", this.closeGui.Bind(this))
        Hotkey("Esc", this.closeGui.Bind(this)) 
        HotIfWinActive()
    }
    closeGui(*) => this.gui.Hide()
}
;#EndRegion

;#Region RichEdit
class RichEdit {
    static __New() {
        this.WM_SETTEXT := 0x000C
        this.EM_SETTEXTEX := 0x0461
        this.WM_CTLCOLORSCROLLBAR := 0x0137
    }
    static Create(gui, options) {
        static WM_VSCROLL := 0x0115
        static COLOR_SCROLLBAR := 0
        static COLOR_BTNFACE := 15
        
        control := gui.AddCustom("ClassRichEdit50W +0x5031b1c4 +E0x20000 " options)  ; Add options here
        SendMessage(0x0443, 0, 0x202020, control)
        DllCall("uxtheme\SetWindowTheme", "Ptr", control.hwnd, "Str", "DarkMode_Explorer", "Ptr", 0)
    
        static WM_SYSCOLORCHANGE := 0x0015
        control.OnMessage(WM_SYSCOLORCHANGE, (ctrl, *) => (
            this.SetSysColor(COLOR_SCROLLBAR, 0x383838),
            this.SetSysColor(COLOR_BTNFACE, 0x383838)
        ))
        PostMessage(WM_SYSCOLORCHANGE, 0, 0, control)
        return control
    }
    static SetText(control, text) {
        this.SetRTF(control, this.FormatText(text))
    }
    static GetText(control) {
        static WM_GETTEXT := 0x000D
        length := control.SendMsg(WM_GETTEXT, 0, 0)
        buf := Buffer(length * 2 + 2)
        control.SendMsg(WM_GETTEXT, length + 1, buf)
        return StrGet(buf)
    }
    static SetRTF(control, rtf) {
        buf := Buffer(StrPut(rtf, "UTF-8"))
        StrPut(rtf, buf, "UTF-8")
        settextex := Buffer(8, 0)
        NumPut("UInt", 0, "UInt", 1200, settextex)
        SendMessage(0x461, settextex, buf, control)
    }
    static FormatText(text, delimiter := "") {
        rtf := "{\rtf{\colortbl;"
            . "\red225\green225\blue225;"
            . "\red150\green150\blue255;"
            . "\red255\green0\blue0;"
            . "}"
        loop parse text, "`n", "`r" {
            line := A_LoopField
            if A_Index = 1
                coloredLine := "\cf2 " this.FormatLineWithDelimiters(line, delimiter)
            else
                coloredLine := "\cf1 " this.FormatLineWithDelimiters(line, delimiter)
            rtf .= coloredLine "\line"
        }
        rtf .= "}"
        return rtf
    }
    static FormatLineWithDelimiters(line, delimiter) {
        if (delimiter = "")
            return this.EscapeRTF(line)
        formattedLine := ""
        remaining := line
        while (pos := InStr(remaining, delimiter)) {
            formattedLine .= this.EscapeRTF(SubStr(remaining, 1, pos - 1))
            formattedLine .= "\cf3" this.EscapeRTF(delimiter) "\cf1"
            remaining := SubStr(remaining, pos + StrLen(delimiter))
        }
        formattedLine .= this.EscapeRTF(remaining)
        return formattedLine
    }

    static EscapeRTF(text) {
        return RegExReplace(text, "([{}])", "\\$1")
    }

    static SetSysColor(index, color) {
        DllCall("SetSysColors", "Int", 1, "Int*", index, "UInt*", color)
        return true
    }
}
;#EndRegion

SetDarkMode(_obj) {
    For v in [135, 136]
        DllCall(DllCall("GetProcAddress", "ptr", DllCall("GetModuleHandle", "str", "uxtheme", "ptr"), "ptr", v, "ptr"), "int", 2)
    if !(attr := VerCompare(A_OSVersion, "10.0.18985") >= 0 ? 20 : VerCompare(A_OSVersion, "10.0.17763") >= 0 ? 19 : 0)
        return false
    DllCall("dwmapi\DwmSetWindowAttribute", "ptr", _obj.hwnd, "int", attr, "int*", true, "int", 4)
}
