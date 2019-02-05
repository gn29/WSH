' =============================================================================================
' Script gets record about User from ActiveDirectory
' To specify concrete user use filter params in filter var:
' @param DisplayName = user name (like 'Ivanov Ivan Ivanovich')
' @param Company = Company name (set, exclude or use *)
'
' Additional feature: script takes every name from file source c:\Temp\fios.txt (line by line)
' and puts the result to c:\Temp\out.txt file
' so specify user name in fios.txt and run it on Windows machine.
'
' 05-02-2019, 29gena@gmail.com
' =============================================================================================

MsgBox "RUN"
set fso = CreateObject("Scripting.FileSystemObject")
set dataSource = fso.OpenTextFile("c:\Temp\fios.txt",1)
set out = fso.OpenTextFile("c:\Temp\out.txt",2,1,0)

set ldap = GetObject("LDAP://RootDSE")
root = ldap.Get("DefaultNamingContext")
dim filter
attrs = "co,sAMAccountName,DisplayName,mail,Company"
scope = "subtree"
set cn = CreateObject("ADODB.Connection")
set cmd = CreateObject("ADODB.Command")
cn.Provider = "ADsDSOObject"
cn.Open "Active Directory Provider"
cmd.ActiveConnection = cn

dim fio, outString
While not dataSource.AtEndOfStream
                fio = dataSource.ReadLine()
                               filter = "(&(DisplayName=" & fio & ")(Company=company_name))"
                               cmd.CommandText = "<LDAP://" & root & ">;" & filter & ";" & attrs & ";subtree"
                               'MsgBox cmd.CommandText
                               set recset = cmd.Execute
                               if not recset.EOF then
                                               recset.Movefirst
                                               do until recset.EOF
                                                               'MsgBox recset.Fields("DisplayName").Value & ";" & recset.Fields("sAMAccountName").Value & ";" & recset.Fields("Company").Value & ";" & recset.Fields("mail").Value
                                                               outString = recset.Fields("DisplayName").Value & ";" & recset.Fields("sAMAccountName").Value & ";" & recset.Fields("Company").Value & ";" & recset.Fields("mail").Value
                                                               out.WriteLine(outString)
                                                               recset.Movenext
                                               loop
                               else 
                                               out.WriteLine(fio & ";NOT_FOUND")
                               end if
                               recset.Close
                
                wend

dataSource.Close
out.Close
cn.Close
MsgBox "DONE"
