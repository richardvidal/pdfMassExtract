﻿#Include pdf2text.ahk
SetWorkingDir %A_ScriptDir%
Gui, New,,"Extracteur et Anonimiseur de Rem:New"

SalaryExtractFromPDF(pdfFile)
{
    ResultStriplineBreaks := true
    TextFromPDF := PDF2Text(PdfFile, ResultStriplineBreaks)    
    regexPos := RegExMatch(TextFromPDF,"(\d*) €", outVar)
    return outVar1
}


Folder:=""
Folder := RegExReplace(Folder, "\\$")  ; Removes the trailing backslash, if present.
SelectedFile:="listeSalaires.txt"
SalaryMin := 28000
SalaryMax := 120000
RoundMlt := 50
nbPdfs:= 0

StartProcess()
{
    global nbPdfs
    global Folder
    global SelectedFile
    global SalaryMin
    global SalaryMax
    global RoundMlt

    if (nbPdfs < 1)
    {
        MsgBox Vous devez choisir un dossier avec des fichiers justificatifs de rémuneration
        return
    }
        
    ListSalary := ""
    StartTime := A_TickCount
    ListErreurs:=""

    Loop Files, %Folder%\*.PDF, R  ; Recurse into subfolders.
    {   
        ProgressRatio:= A_Index / nbPdfs * 100
        GuiControl,, MyProgress, %ProgressRatio%
        GuiControl,, StatusBarText, %A_Index% / %nbPdfs% --- %A_LoopFileName%
        PdfFile := A_LoopFileFullPath
        Salary := SalaryExtractFromPDF(PdfFile)
        if (Salary < 1)
        {
            ListErreurs.=A_LoopFileFullPath "`n"
        }
        else
        {
            NewSalary := Floor(Salary)
            NewSalary := NewSalary - Mod(NewSalary,RoundMlt)
            NewSalary := Min(SalaryMax,NewSalary)    
            NewSalary := Max(SalaryMin,NewSalary)
            ListSalary.= NewSalary "`n"
        }
        
    }
    sort, ListSalary, ND`n	
    if FileExist(SelectedFile)
        FileAppend, Processus lancé sur un fichier déjà existant, %SelectedFile%-erreurs.txt

    FileAppend, %ListSalary%, %SelectedFile%
    EndTime := A_TickCount
    ElapsedSeconds := (EndTime - StartTime)/1000.0

    
    MsgBox, C'est fait en  %ElapsedSeconds% secondes
    if (ListErreurs!="")
    {
        FileAppend, %ListErreurs%, %SelectedFile%-erreurs.txt
        MsgBox, Des erreurs ont été detectées, regardez le fichier  %SelectedFile%-erreurs.txt
    }
    SelectedFile :=""
    
}

choixDossier(){
    global Folder
    global nbPdfs
    FileSelectFolder, Folder    
    Folder := RegExReplace(Folder, "\\$")  ; Removes the trailing backslash, if present.

    ; We count file
    Loop Files, %Folder%\*.PDF, R  ; Recurse into subfolders.
    {    
        nbPdfs +=1
    }
    MsgBox, %nbPdfs% fichiers detectés.
    GuiControl, Text, btnDossier, Dossier source : %Folder% (%nbPdfs% fichiers)
    
}

choixFichier(){
    global SelectedFile

    FileSelectFile, SelectedFile, S
    if (SelectedFile = "")    
        SelectedFile.="listeSalaires.txt"
    if FileExist(SelectedFile)
        MsgBox, Attention ce fichier existe déjà, il y a risque de doublon lors de l'extraction
    GuiControl, Text, btnFichier, Fichier cible : %SelectedFile%
}

; Generated by Auto-GUI 3.0.1
#SingleInstance Force
#NoEnv

SetBatchLines -1


Gui Add, StatusBar, vStatusBarText, ...
Gui Font, s12
Gui Add, Text, x12 y7 w597 h27 +0x200, Extracteur de rémuneration
Gui Font, s9 c0x0080C0
Gui Font, c0xA72D7A, Arial
Gui Add, Text, x12 y39 w597 h23 +0x200, CSE ACCENTURE SAS
Gui Font, s9 c0x0080C0
Gui Add, Text, x12 y69 w620 h23 +0x200, 1. Conversion du document
Gui Add, Text, x12 y96 w620 h23 +0x200, 2. Extraction du montant
Gui Add, Text, x12 y123 w620 h23 +0x200, 3. Arrondis, Plancher et Plafond
Gui Add, Text, x12 y150 w620 h23 +0x200, 4. Listing et Tri
Gui Add, Text, x12 y177 w620 h23 +0x200, 5. Enregistrement
Gui Add, Button, x16 y241 w588 h35 VbtnDossier GchoixDossier, Dossier source :
Gui Add, Button, x16 y275 w588 h35 VbtnFichier GchoixFichier, Fichier cible :
Gui Add, GroupBox, x12 y221 w598 h151, Réglages
Gui Add, Edit, x25 y339 w120 h21 VSalaryMin, %SalaryMin%
Gui Add, Text, x25 y309 w120 h23 +0x200, Plancher :
Gui Add, Progress, x12 y419 w588 h20 -Smooth vMyProgress, 0
Gui Add, Text, x242 y309 w120 h23 +0x200, Arrondi :
Gui Add, Text, x422 y309 w120 h23 +0x200, Plafond :
Gui Add, Edit, x242 y339 w120 h21 VRoundMlt, %RoundMlt%
Gui Add, Edit, x422 y339 w120 h21 VSalaryMax, %SalaryMax%

Gui Font, s9 c0x0080C0
Gui Add, Button, x12 y383 w588 h35 GstartProcess, Débuter le processus


Gui Show, x644 y310 w617 h466, Extracteur Rémuneration CSE
Return

GuiEscape:
GuiClose:
    ExitApp


