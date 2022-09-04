#Classic ASP MailerSend.com Implementation

MeilerSend.com REST API implementation in Classic ASP ( VB Script ) for send emails

There are problems with using the MailerSend SMTP service in Classic ASP (CDO). I think the problem is with the required TLS connection.

To solution was to use the REST API. After some research I found a working JSON library from Gerrit van Kuipers and combined it with code for making http requests to get it working.

#Classic ASP REST API Implementation

The same code can be used for any kind of REST requests in Classic ASP.

