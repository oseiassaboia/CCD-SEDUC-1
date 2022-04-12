<%@ Control Language="VB" AutoEventWireup="false" CodeFile="ctrTopo.ascx.vb" Inherits="ctrTopo" %>


            <a href="" class="logo">
                <!-- Add the class icon to your logo image or logo icon to add the margining -->
                <img src="Imagens/ico_rh_30px.png" alt="" /> <asp:Label ID="lblSistema" runat="server" Text="Recursos Humanos" />
            </a>
            <!-- Header Navbar: style can be found in header.less -->
            <nav class="navbar navbar-static-top" >
                <!-- Sidebar toggle button-->
                <a href="#" class="sidebar-toggle" data-toggle="offcanvas" role="button">
                    <span class="sr-only">Toggle navigation</span>
                </a>
                
                <div class="navbar-custom-menu">
                    <ul class="nav navbar-nav">
                        <!-- User Account: style can be found in dropdown.less -->
                        <li class="dropdown user user-menu">
                            <a href="#" class="dropdown-toggle" data-toggle="dropdown">
                                <asp:Image ID="imgUsuario" runat="server" class="user-image" alt="User Image"/>
                                <span class="hidden-xs"><asp:Label ID="lblUsuario" runat="server" Text="Jane da Silva" /></span>
                            </a>
                            <ul class="dropdown-menu" id="divUsuario" runat="server">
                              <!-- User image -->
                              <li class="user-header">
                                <asp:Image id="imgFoto" runat="server" cssclass="img-circle" alt="User Image" ImageUrl="img/Fotos/user7-128x128.jpg" />

                                <p>
                                  <asp:Label ID="lblNomeUsuario" runat="server" Text="Rh" />
                                  <small></small>
                                </p>
                              </li>
                              <!-- Menu Body -->
<%--                              <li class="user-body">
                                <div class="row">
                                  <div class="col-xs-4 text-center">
                                    <a href="#">Turmas</a>
                                  </div>
                                  <div class="col-xs-4 text-center">
                                    <a href="#">Disciplinas</a>
                                  </div>
                                  <div class="col-xs-4 text-center">
                                    <a href="#">Frequência</a>
                                  </div>
                                </div>
                              </li>--%>
                              <!-- Menu Footer-->
                                <li id="liSair" runat="server"  class="user-footer">
                                <div class="pull-left">
                                    <asp:LinkButton ID="lnkControleAcesso" runat="server" class="btn btn-lg btn-block btn-danger" Text="Controle de Acesso"/>
                                </div>
                                <div class="pull-right">
                                    <asp:LinkButton ID="lnkSair" runat="server" class="btn btn-lg btn-block btn-danger" Text="Sair"/>
                                </div>
                            </ul>
                        </li>
                    </ul>
                </div>
                <div class="navbar-custom-menu">
                    <ul class="nav navbar-nav">
                        <li>
                            <a href="frmSelecaoEscola.aspx" >
                                <span><b><asp:Label ID="lblEscola" runat="server" /></b></span>
                            </a>
                        </li>
                    </ul>
                </div>
            </nav>