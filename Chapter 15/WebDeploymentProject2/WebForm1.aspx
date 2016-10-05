<%@ Page Language="vb" AutoEventWireup="false" Codebehind="WebForm1.aspx.vb" Inherits="WebDeploymentProject.WebForm1" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
   <head>
      <title>WebDeploymentProject</title>
      <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
      <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
      <meta name="vs_defaultClientScript" content="JavaScript">
      <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
      <script language="jscript">
   function ChangeLanguage(language) {
      // Save the culture in hidden form field
      Form1.languageInput.value = language
   }
      </script>
   </head>
   <body ms_positioning="GridLayout">
      <form id="Form1" method="post" runat="server">
         <table>
            <tr>
               <td>
                  <input id="languageInput" type="hidden" size="1" name="languageInput" runat="server" style="WIDTH: 8px; HEIGHT: 25px">
               </td>
            </tr>
         </table>
         <table width="100%" align="center" border="0">
            <tr width="100%">
               <td valign="top" align="center" width="50%">
                  <asp:image id="staticImage" runat="server" imageurl="1590590422.jpg"></asp:image>
               </td>
               <td valign="top" align="center" width="50%">
                  <asp:image id="dynamicImage" runat="server"></asp:image>
               </td>
            </tr>
            <tr width="100%">
               <td valign="middle" align="center" width="50%">
                  <asp:label id="invokeWebServiceTimeLabel" runat="server" width="408px">Web Service invoked at </asp:label>
               </td>
               <td valign="middle" align="center" width="50%">
                  <asp:label id="initAssemblyTimeLabel" runat="server" width="408px">.NET component initialized at </asp:label>
               </td>
            </tr>
            <tr width="100%">
               <td valign="middle" align="center" width="100%" colspan="2">
                  <asp:button id="clickMeButton" runat="server" text="Click Me"></asp:button>
               </td>
            </tr>
         </table>
         <table cellspacing="3" cellpadding="0" width="100%" align="center" border="0">
            <tr width="100%">
               <td valign="middle" align="center" width="100%">
                  <a id="daLangAnchor" onclick="ChangeLanguage('da')" href="WebForm1.aspx" runat="server">
                     <img alt='<%=m_pageResourceManager.GetString("DANISH")%>' src="images\FLGDEN.GIF" border="0"></a>
                  <a id="enLangAnchor" onclick="ChangeLanguage('en')" href="WebForm1.aspx" runat="server">
                     <img alt='<%=m_pageResourceManager.GetString("ENGLISH")%>' src="images\FLGUK.GIF" border="0"></a>
               </td>
            </tr>
         </table>
      </form>
   </body>
</html>
