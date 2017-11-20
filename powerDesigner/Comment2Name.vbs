
'******************************************************************************
'* File:     ModelComment2Name.vbs
'* Title:    将模型注释复制到名字
'* Author:   lsj qq:273364475
'* Created:  2017-11-09
'* Mod By:   
'* Modified: 
'* Version:  1.0
'* Comment:  
'*  v1.0 
'******************************************************************************

Option   Explicit 
ValidationMode   =   True 
InteractiveMode   =   im_Batch
Dim   mdl   '   the   current   model
'   get   the   current   active   model 
Set   mdl   =   ActiveModel 
If   (mdl   Is   Nothing)   Then 
      MsgBox   "There   is   no   current   Model " 
ElseIf   Not   mdl.IsKindOf(PdPDM.cls_Model)   Then 
      MsgBox   "The   current   model   is   not   an   Physical   Data   model. " 
Else 
      ProcessFolder   mdl 
End   If
Private   sub   ProcessFolder(folder) 
On Error Resume Next
      Dim   Tab   'running     table 
      for   each   Tab   in   folder.tables 
            if   not   tab.isShortcut   then 
                  'tab.name   =   tab.comment
                  Dim   col   '   running   column 
                  for   each   col   in   tab.columns 
                  if col.comment="" then
                  else
                        col.name=   col.comment 
                  end if
                  next 
            end   if 
      next
      Dim   view   'running   view 
      for   each   view   in   folder.Views 
            if   not   view.isShortcut   then 
                  view.name   =   view.comment 
            end   if 
      next
      '   go   into   the   sub-packages 
      Dim   f   '   running   folder 
      For   Each   f   In   folder.Packages 
            if   not   f.IsShortcut   then 
                  ProcessFolder   f 
            end   if 
      Next 
end   sub