Imports System
Imports System.Data
Imports System.Web
Imports System.Web.Mail
Imports System.Collections.Specialized

Namespace DLL
  Public Class GestError

        Public Function GetHTMLError(ByVal Ex As Exception, ByVal style As String, ByVal ChiudiClickEvent As String) As String

            'Returns HTML an formatted error message.

            Dim Heading As String
            Dim MyHTML As String
            Dim Error_Info As New NameValueCollection

            Dim HTMLDesign As String

            HTMLDesign = "<HTML><HEAD><link rel=""stylesheet"" type=""text/css"" href=" & """" & style & """" & "  ></HEAD>" & vbCrLf
            HTMLDesign = HTMLDesign & "<body bottomMargin=""0"" leftMargin=""0"" topMargin=""0"" marginwidth=""0"" marginheight=""0"" class=""BODYPAGE"" >" & vbCrLf
            HTMLDesign = HTMLDesign & "<table cellSpacing=""0"" cellPadding=""1"" width=""90%"" align=""center"" border=""0"">" & vbCrLf
            HTMLDesign = HTMLDesign & "<tr><td  class=""TITLEBOLD"" width=""25%"">" & vbCrLf
            HTMLDesign = HTMLDesign & "<br>Siamo Spiacenti, ma si è verificato un errore durante il processo dell'ultima richiesta.<br><br>Questo Potrebbe essere il risultato di un input" & vbCrLf
            HTMLDesign = HTMLDesign & "di un valore errato, oppure di una anomalia nel nostro codice.<br><br>Scusateci per l'inconveniente.</td></tr>" & vbCrLf
            HTMLDesign = HTMLDesign & "<tr><td align=""center"">" & vbCrLf
            HTMLDesign = HTMLDesign & "<INPUT type=""button"" class=""BOTTONE"" value=""Chiudi"" onclick=" & """" & ChiudiClickEvent & """" & " > " & vbCrLf
            HTMLDesign = HTMLDesign & "</td></tr>" & vbCrLf
            HTMLDesign = HTMLDesign & "<tr><td><INPUT type=""checkbox"" id=""Errore"" onclick=""if(this.checked==false){document.getElementById('Error2').style.display='none';}else{document.getElementById('Error2').style.display='';}""><span class=""TITLEBOLD""> Visualizza dettaglio Errore </span></td></tr>" & vbCrLf
            HTMLDesign = HTMLDesign & "<tr><td class=""BODYDIV"" width=""855""><div id=""Error2""  style=""display:none"">" & vbCrLf
            HTMLDesign = HTMLDesign & "<TABLE BORDER=""0"" WIDTH=""95%"" CELLPADDING=""1"" CELLSPACING=""0""><TR><TD class=""LABELBOLD""><!--HEADER--></TD></TR></TABLE>" & vbCrLf
            HTMLDesign = HTMLDesign & "<span class=""ERRORSTYLE""> Error - " & Ex.Message & "</span><BR><BR>" & vbCrLf
            Error_Info.Add("Message", CleanHTML(Ex.Message))
            Error_Info.Add("Source", CleanHTML(Ex.Source))
            Error_Info.Add("TargetSite", CleanHTML(Ex.TargetSite.ToString()))
            Error_Info.Add("StackTrace", CleanHTML(Ex.StackTrace))
            HTMLDesign = HTMLDesign & CollectionToHtmlTable(Error_Info)
            HTMLDesign = HTMLDesign & "</div>" & vbCrLf
            HTMLDesign = HTMLDesign & "</td></tr>"
            HTMLDesign = HTMLDesign & "</table>"
            HTMLDesign = HTMLDesign & "</BODY>"
            HTMLDesign = HTMLDesign & "</HTML>"

            Return HTMLDesign

        End Function

        Public Function SendEmailError(ByVal ex As Exception, ByVal mailTo As String, ByVal mailSubject As String, ByVal mailFrom As String, ByVal NameSmtpServer As String)

            Dim mail As New MailMessage

            Dim ErrorMessage = "Descrizione Errore : " & ex.Message & ex.StackTrace
            mail.To = mailTo
            mail.Subject = mailSubject
            mail.Priority = MailPriority.High
            mail.BodyFormat = MailFormat.Text
            mail.Body = ErrorMessage
            mail.From = mailFrom
            SmtpMail.SmtpServer = SmtpMail.SmtpServer.Insert(0, NameSmtpServer)
            SmtpMail.Send(mail)

        End Function

        Public Function CollectionToHtmlTable(ByVal Collection As NameValueCollection) As String
            Dim TD As String
            Dim MyHTML As String
            Dim i As Integer
            TD = "<TD class=""LABELBOLD""><!--VALUE--></TD>"
            MyHTML = "<TABLE width=""95%"">" & _
            " <TR>" & _
            TD.Replace("<!--VALUE-->", " <B>Name</B>") & _
            " " & TD.Replace("<!--VALUE-->", " <B>Value</B>") & "</TR>"
            'No Body? -> N/A
            If (Collection.Count <= 0) Then
                Collection = New NameValueCollection
                Collection.Add("N/A", "")
            Else
                'Table Body
                For i = 0 To Collection.Count - 1
                    MyHTML += "<TR valign=""top"">" & _
                    TD.Replace("<!--VALUE-->", Collection.Keys(i)) & " " & _
                    TD.Replace("<!--VALUE-->", Collection(i)) & "</TR> "
                Next i
            End If
            'Table Footer
            Return MyHTML & "</TABLE>"
        End Function

        Private Function CollectionToHtmlTable(ByVal Collection As HttpCookieCollection) As String
            'Converts HttpCookieCollection to NameValueCollection
            Dim NVC = New NameValueCollection
            Dim i As Integer
            Dim Value As String
            Try
                If Collection.Count > 0 Then
                    For i = 0 To Collection.Count - 1
                        NVC.Add(i, Collection(i).Value)
                    Next i
                End If
                Value = CollectionToHtmlTable(NVC)
                Return Value
            Catch MyError As Exception
                MyError.ToString()
            End Try
        End Function

        Private Function CollectionToHtmlTable(ByVal Collection As System.Web.SessionState.HttpSessionState) As String
            'Converts HttpSessionState to NameValueCollection
            Dim NVC = New NameValueCollection
            Dim i As Integer
            Dim Value As String
            If Collection.Count > 0 Then
                For i = 0 To Collection.Count - 1
                    NVC.Add(i, Collection(i).ToString())
                Next i
            End If
            Value = CollectionToHtmlTable(NVC)
            Return Value
        End Function

        Private Function CleanHTML(ByVal HTML As String) As String

            If HTML.Length <> 0 Then
                HTML.Replace("<", "<").Replace("\r\n", "<BR>").Replace("&", "&").Replace(" ", " ")
            Else
                HTML = ""
            End If

            Return HTML

        End Function


    End Class
End Namespace