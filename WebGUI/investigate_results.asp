<%
'Dimension variables
dim sName, sDept, sLocation, sStatus, sReason
dim URLRed

'Read in the record number to be updated
sName = Request.Form("Name")
sFolder = Request.Form("Folder")
sEGrp = Request.Form("Group")
sInjectGrp = Request.Form("InjectGrp")
sPermission = Request.Form("Permission")

URLRed = "investigate.asp?filter=Search"

if not sName = "" then URLRed = URLRed & "&name=" & sName
if not sFolder = "" then URLRed = URLRed & "&folder=" & sFolder
if not sEGrp = "" then URLRed = URLRed & "&EGrp=" & sEGrp
if not sInjectGrp = "" then URLRed = URLRed & "&injectgrp=" & sInjectGrp
if not sPermission = "Any" then URLRed = URLRed & "&permission=" & sPermission

if URLRed = "investigate.asp?filter=Search" then URLRed = "investigate.asp"

Response.Redirect URLRed
%>