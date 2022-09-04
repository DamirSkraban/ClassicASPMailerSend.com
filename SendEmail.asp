<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="aspJSON1.19.asp" -->
<%


function SendEmail (toMail, fromMail, fromName,  subject, body)

    dim data

    data=GetMailerSendJson(toMail, fromMail, fromName,  subject, body)

    dim apiKey

    apiKey="API KEY from MailerSend.com"

	Set objXmlHttpMain = CreateObject("Msxml2.ServerXMLHTTP")
    objXmlHttpMain.open "POST", "https://api.mailersend.com/v1/email", false

    objXmlHttpMain.setRequestHeader "Authorization", "Bearer " + apiKey
    objXmlHttpMain.setRequestHeader "Content-Type", "application/json"

    objXmlHttpMain.setRequestHeader "CharSet", "charset=UTF-8"
    objXmlHttpMain.setRequestHeader "Accept", "application/json"


    objXmlHttpMain.send data
    
    set objXmlHttpMain=nothing

end function




function GetMailerSendJson(byval toMail,byval fromMail,byval fromName, byval subject, byval body)


    Set oJSON = New aspJSON

    With oJSON.data

        .Add "from", oJSON.Collection()       
        With .item("from")
            .Add "email", fromMail
            .Add "name", fromName
        End With           

        .Add "to", oJSON.Collection()

        With oJSON.data("to")

            .Add 0, oJSON.Collection()                 
            With .item(0)
                .Add "email", toMail
                
            End With



        End With

        .Add "subject", subject      
        .Add "html", body     
    End With

    GetMailerSendJson = oJSON.JSONoutput()  

    
end function



call SendEmail("to@gmail.com", "from@gmail.com", "From Name", "Subject", "Body")
%>