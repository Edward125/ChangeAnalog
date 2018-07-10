dim analogpath,fs,oFolder,oFiles,sLine,bPanel,sLineStr
analogpath = InputBox("Input the analog path","PATH")
if analogpath<> "" then 
	 i =0 
    if IsExitAFolder(analogpath) then 
       Set fs = CreateObject("Scripting.FileSystemObject")
       Set oFolder = fs.GetFolder(analogpath)    
       Set oFiles = oFolder.Files 
       For each file In oFiles
         sExt = fs.GetExtensionName(file)    
         sExt = LCase(sExt)
         if (sExt = "") and Right(file.name,1)<>"~" and InStr(file.name,"r")>0 or InStr(file.name,"pr")>0 then 
           if  InStr(Left(file.name,2),"%") > 0 then 
              bPanel = True
              sLine = "i to ""#%GND"""
           Else 
              bPanel = False  
              sLine = "i to ""GND"""  
           end If
           Set f = fs.OpenTextFile(file)
           read = f.readall
           'f.close
           
          
           If InStr(read,sLine) >0 Then     
             On Error Resume Next         	
             fs.CopyFile file,file &".vb",False
             fs.delect file
             Set f = fs.OpenTextFile(file,2,true)
             i = i+1
             arr= Split(read,vbCrLf)
             sBus = False
             iBus = False
             for each substr in arr
            
             bConnect = False
             if InStr(substr,"connect s") >0 and bConnect = False then 'and sBus = True and iBus = False then 
             	 substr = replace(substr,"connect s","connect i")
             	 bConnect = True
             end if 

             if InStr(SubStr,"connect i") > 0 and bConnect = False then 'and iBus = True	and sBus = False then
               substr = replace(substr,"connect i","connect s")
               bConnect = True
             end if
             
             f.WriteLine(substr)
             
             next
             'f.close
           end if 
         End If 
       Next
    MsgBox "Complete."& "Total rewrite " & i & " resistance files." ,0,"Complete"
    else
        msgbox "The path doesn't exist...",0,"Not Exist"
    end if 
end if



Function IsExitAFile(filespec)
        Dim fso
        Set fso=CreateObject("Scripting.FileSystemObject")        
        If fso.fileExists(filespec) Then         
        IsExitAFile=True        
        Else IsExitAFile=False        
        End If
End Function 

Function IsExitAFolder(folderspec)
        Dim fso
        Set fso=CreateObject("Scripting.FileSystemObject")        
        If fso.folderExists(folderspec) Then         
        IsExitAFolder=True        
        Else IsExitAFolder=False        
        End If
End Function 