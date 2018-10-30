'################################################################### 
'##           Script to check the status of machines              ## 
'##           Author: Unknown                                    ## 
'##           Date: 03-30-2012                                 ## 
'##           modified by: Vikas Sukhija                          ## 
'################################################################### 
 
'# call excel applicationin visible mode 
 
Set objExcel = CreateObject("Excel.Application") 
  
objExcel.Visible = True 
  
objExcel.Workbooks.Add 
  
intRow = 2 
  
'# Define Labels  
  
objExcel.Cells(1, 1).Value = "Machine Name" 
  
objExcel.Cells(1, 2).Value = "Results" 
  
  
'# Create file system object for reading the hosts from text file 
 
 
Set Fso = CreateObject("Scripting.FileSystemObject") 
  
Set InputFile = fso.OpenTextFile("HOSTS.Txt") 
  
'# Loop thru the text file till the end  
  
Do While Not (InputFile.atEndOfStream) 
  
HostName = InputFile.ReadLine 
   
'# Create shell object for Pinging the host machines 
 
  
Set WshShell = WScript.CreateObject("WScript.Shell") 
  
Ping = WshShell.Run("ping -n 1 " & HostName, 0, True) 
  
  
objExcel.Cells(intRow, 1).Value = HostName