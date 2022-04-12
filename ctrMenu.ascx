<%@ Control Language="VB" AutoEventWireup="false" CodeFile="ctrMenu.ascx.vb" Inherits="ctrMenu" %>

<!-- sidebar: style can be found in sidebar.less -->
<section class="sidebar">
    <!-- Sidebar user panel -->
    <div class="user-panel" id="divGoverno" runat="server" visible="true">
        <center>
            <img src="Imagens/rh.png" alt="" /></center>
    </div>
    <div class="user-panel" id="divUsuario" runat="server">
        <div class="pull-left image">
            <%--<asp:Image ID="imgFoto" runat="server" CssClass="img-circle" alt="User Image" ImageUrl="img/Fotos/user7-128x128.jpg" />--%>
            <asp:Image ID="imgFoto" runat="server" CssClass="img-circle" alt="User Image" />
        </div>
        <div class="pull-left info">
            <p>
                Olá,
                <asp:Label ID="lblUsuario" runat="server" Text="teste" />
            </p>
            <a href="#"><i class="fa fa-circle text-success"></i>Online</a>
        </div>
    </div>
    <!-- sidebar menu: : style can be found in sidebar.less -->
    <ul class="sidebar-menu">
        <li class="header">MENU</li>

        <li id="liHeaderAreaEnsino" runat="server" class="treeview">

                    <asp:LinkButton ID="lnkServidorCargaHoraria" runat="server"><i class="fa fa-folder-open-o"></i>Servidor Carga Horária</asp:LinkButton>
      
        </li>        


</section>

