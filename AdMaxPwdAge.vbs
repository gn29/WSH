' Find out how many days the Ad User in corporate network can you his password
' so you can predict expiration date and show it to him
' Works on Windows Script Host: vbscript
' 08-02-2019
' @author 29gena@gmail.com

set ldap = GetObject("LDAP://RootDSE")
root = ldap.Get("DefaultNamingContext")
set adUser = GetObject("LDAP://"&root)
set a = adUser.MaxPwdAge
b = CCur((a.HighPart * 2^32) + a.LowPart) / CCur(-864000000000)
MsgBox b
