' Bohack
' 01/09/18
' Firewall alias parser to find orphaned aliases.
' Before deleting any aliases grep the config to be 100% sure

Option Explicit
Dim objXMLDoc, Item, aryList
Dim objAliasMasterList, Alias, ObjAliasConfigured
Dim aryXMLList, XMLPath, objAliasCollection
Dim objFSO, objFile, colConfigFile, strLine, Count
Const ForReading = 1

If WScript.Arguments.Count < 1 Then
   Wscript.Echo "Usage:" & Wscript.ScriptName & " {name of the pfsense XML files}"
   WScript.Quit
End If

Set objXMLDoc = CreateObject("Microsoft.XMLDOM")
Set objAliasMasterList = CreateObject("Scripting.Dictionary")

objXMLDoc.Async = "False"
objXMLDoc.Load(Wscript.Arguments(0))

'List of XML paths to check for aliases in use
aryXMLList = Array("//pfsense/filter/rule/source/address","//pfsense/filter/rule/destination/address", _
"//pfsense/filter/rule/source/port","//pfsense/filter/rule/destination/port","//pfsense/nat/outbound/rule/source/network", _
"//pfsense/nat/outbound/rule/sourceport","//pfsense/nat/outbound/rule/source/network","//pfsense/nat/rule/outbound/source/port", _
"//pfsense/nat/rule/outbound/destination/network","//pfsense/nat/rule/outbound/destination/port","//pfsense/nat/rule/target", _
"//pfsense/nat/rule/local-port")

'Read master list into objAlias
For Each XMLPath in aryXMLList
  Set ObjAliasCollection=objXMLDoc.selectNodes(XMLPath)
  For Each Item in ObjAliasCollection
   If Not objAliasMasterList.Exists(Item.Text) Then
     objAliasMasterList.Add Item.Text, XMLPath
   End If
  Next
  Set ObjAliasCollection = Nothing
Next

'Read nested aliases into objAlias that is not an FQDN or IP address
Set ObjAliasCollection=objXMLDoc.selectNodes ("//pfsense/aliases/alias/address")
For Each Alias in ObjAliasCollection
  aryList = Split(Alias.Text," ")
  For Each Item in aryList
    If (not objAliasMasterList.Exists(item)) AND Instr(Item,".")<1 AND (Instr(Item,":")<1) AND NOT IsNumeric(Item) Then
      objAliasMasterList.Add Item, "//pfsense/aliases/alias/address"
    End If
  Next
Next

'Print out any orphaned aliases
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile(Wscript.Arguments(0), ForReading)
colConfigFile = Split(objFile.ReadAll, vbLf)

Wscript.Echo "List of Orphaned Aliases"
Wscript.Echo "-----------------------------"
Set ObjAliasConfigured=objXMLDoc.selectNodes ("//pfsense/aliases/alias/name")
For Each Alias in ObjAliasConfigured
  If Not objAliasMasterList.Exists(Alias.Text) Then
    Count = 0
    For Each strLine in colConfigFile
      Count = Count + 1
      If Instr(strLine, Alias.Text) > 0 Then
        If Instr(strLine,"<name>") < 1 Then
          Wscript.Echo Alias.Text & " - was found in config at line " & Count & "," & Instr(strLine, Alias.Text)
          Wscript.Echo String(len(Alias.Text)+3," ") & Mid(strLine,Instr(strLine,"<"),Len(strLine))
        Else
          Wscript.Echo Alias.Text & " - was NOT found anywhere in the config outside of the alias tags"
        End If
      End If
    Next
  End If
Next

Set objFile = Nothing
Set objFSO = Nothing
Set ObjAliasConfigured = Nothing
Set ObjAliasCollection = Nothing
