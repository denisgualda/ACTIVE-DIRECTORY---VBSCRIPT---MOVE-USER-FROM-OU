'EXEMPLE SCRIPT INICIAL
'ES NECESITA AQUESTA INFORMACIO PER EXECUTAR L'ORDRE
'----------------------------------------------------------------------------------------------------------------
'strNewParentDN = "LDAP://OU=Usuaris,DC=domini,DC=lab"    'NOVA OU
'strObjectDN    = "LDAP://cn=usuari02,OU=Usuaris02,DC=domini,DC=lab"     'DISTINGUESHED NAME ACTUAL D'USUARI
'strObjectRDN   = "cn=usauri02"         'NOM D'USUARI
'----------------------------------------------------------------------------------------------------------------
'ORIGINAL--------------------------------------
'strNewParentDN = "LDAP://<NewParentDN>"
'strObjectDN    = "LDAP://cn=jsmith,<OldParentDN>"
'strObjectRDN   = "cn=jsmith"
'set objCont = GetObject(strNewParentDN)
'objCont.MoveHere strObjectDN, strObjectRDN
'---------------------------------------------


'-------------------------------------------------
'(FET)
'OBTENIR DISTINGUESHED NAME OU D'USUARI MIRALL (ON ANIRÀ A PARAR USUARI2)
strUsuari1 = InputBox("Entra usuari 1 (Usuari Mirall)")          '--> strNewParentDN = "LDAP://OU=Usuaris,DC=domini,DC=lab"  
Set objSystemInfo = CreateObject("ADSystemInfo") 
strDomain = objSystemInfo.DomainShortName
'obenim Distingueshed name de usuari 1
strDNUSuari01 = GetUserDN(strUsuari1,StrDomain)
'************************************
	'Separem CN del distinguehsed name obtingut per obtenir el distingueshed name de la OU de l'usuari.
	retallastrDNUsuari01 = Split(strDNUSuari01,",",2)
		for each x in retallastrDNUsuari01
			i = i +1 
				if i = 2 Then
					strDNOUUsuari01 = x
				end if
		next
	'************************************

strDNOUUsuari01 = "LDAP://" & strDNOUUsuari01
wscript.echo strDNOUUsuari01
'-------------------------------------------------

'-------------------------------------------------
'(FET)
'OBTENIR DISTINGUESHED NAME USUARI NOU strUsuari2 (DESTÍ)         '--> strObjectDN    = "LDAP://cn=usuari02,OU=Usuaris02,DC=domini,DC=lab"
strUsuari2 = InputBox("Entra usuari 2 (Alta Usuari")
strDNUsuari02 = GetUserDN(strUSuari2,StrDomain)

strDNUsuari02 = "LDAP://" & strDNUsuari02
wscript.echo strDNUsuari02
'-------------------------------------------------

'-------------------------------------------------
'(FET)
'OBTENIR CN DE DISTINGUESHED NAME D'USUARI NOU strUsuari2 (DESTÍ)    '' --> strObjectRDN   = "cn=usauri02"         'NOM D'USUARI
strCNUsuari2  = "cn=" & strUsuari2
wscript.echo strCNUsuari2                       
'-------------------------------------------------

'-------------------------------------------------
'EXECUTA PER MOURE USUARI Ou

set objCont = GetObject(strDNOUUsuari01)
objCont.MoveHere strDNUsuari02, strCNUsuari2
'-------------------------------------------------


Function GetUserDN(byval strUserName,byval strDomain)
	set objTrans = CreateObject("NameTranslate")
	objTrans.Init 1, strDomain
	objTrans.Set 3, strDomain & "\" & strUserName 
	strUserDN = objTrans.Get(1) 
	GetUserDN = strUserDN

End function