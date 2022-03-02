#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

;To convert HEX values to DEC
HexToDec(Hex)
{
	if (InStr(Hex, "0x") != 1)
		Hex := "0x" Hex
	return, Hex + 0
}

maxTsByOd := []

;We check all files once to find which one is the latest
Loop Files, *.PDF
{   
    ;We extract useful values from filename
    PdfFile := A_LoopFileName
    SplitPath, A_LoopFileName, name, dir, ext, FileNoExt, drive
    regexmatch(FileNoExt, "(od_10_)(.*?)(?=_)", results)
    od := results2
    ts := HexToDec(SubStr(FileNoExt, -12))
    
    ;We store the Max timestamp value found so far for this od
    If (maxTsByOd.HasKey(od)) 
        maxTsByOd[od]:=Max(maxTsByOd[od],ts)
    Else 
        maxTsByOd[od]:= ts
}

;We check all files again and rename oldest to keep only one
Loop Files, *.PDF
{
    ;We extract useful values from filename
    PdfFile := A_LoopFileName
    SplitPath, A_LoopFileName, name, dir, ext, FileNoExt, drive
    regexmatch(FileNoExt, "(od_10_)(.*?)(?=_)", results)
    od := results2
    ts := HexToDec(SubStr(FileNoExt, -12))

    If (maxTsByOd[od]!=ts) ;only one file (the max TS) per OD will not match that condition
        FileMove,  %A_LoopFileName%, %A_LoopFileName%.bak
}
MsgBox, Dédoublonnage terminé