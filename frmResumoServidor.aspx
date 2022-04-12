<%@ Page Title="" Language="VB" MasterPageFile="~/MasterPageCargaHoraria.master" AutoEventWireup="false" CodeFile="frmResumoServidor.aspx.vb" Inherits="frmResumoServidor" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="cc1" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
 <asp:UpdatePanel ID="UpdatePanel1" runat="server">
        <ContentTemplate>
            <section id="Cadastro" runat="server" class="content">
                <!-- Small boxes (Stat box) -->
                <div class="box">

                    <div class='row'>
                        <div class='col-sm-3'>
                            <div class="box-body">
                                <div class="box box-primary">
                                    <div class="box-body box-profile">
                                        <asp:Repeater ID="dtlProfessor" runat="server">
                                            <ItemTemplate>
                                                <%--<a href="frmImagemPerfil.aspx">--%>
                                                    <img class="profile-user-img img-responsive img-circle" src='frmPrincipalFotos.aspx?idPessoa=<%# Eval("RH01_ID_PESSOA") %>' alt="User profile picture"></a>
                                            </ItemTemplate>
                                        </asp:Repeater>

                                        <h3 class="profile-username text-center">
                                            <asp:Label ID="lblNome" runat="server" />
                                        </h3>

                                        <p class="text-muted text-center">
                                            <asp:Label ID="lblCargo" runat="server" />
                                        </p>

                                        <ul class="list-group list-group-unbordered">
                                            <li class="list-group-item">
                                                <b>CPF</b> <a class="pull-right">
                                                    <asp:Label ID="lblCPF" runat="server" /></a>
                                            </li>
                                            <li class="list-group-item">
                                                <b>Matrícula</b> <a class="pull-right">
                                                    <asp:Label ID="lblMatricula" runat="server" /></a>
                                            </li>
                                            <li class="list-group-item">
                                                <b>C.H.</b> <a class="pull-right">
                                                    <asp:Label ID="lblCargaHoraria" runat="server" /></a>
                                            </li>
                                            <li class="list-group-item">
                                                <b>Perfil</b> <a class="pull-right">
                                                    <asp:Label ID="lblPerfil" runat="server" /></a>
                                            </li>
                                            <li class="list-group-item">
                                                <b>Situação</b> <a class="pull-right">
                                                    <asp:Label ID="lblSituação" runat="server" /></a>
                                            </li>
                                            <li class="list-group-item">
                                                <b>Vínculo:</b> <a class="pull-right">
                                                    <asp:Label ID="lblVinculo" runat="server" /></a>
                                            </li>
                                        </ul>

                                        <%-- <a href="#" class="btn btn-primary btn-block"><b>Follow</b></a>--%>
                                    </div>
                                </div>
                            </div>
                        </div>

                        <div class='col-sm-9'>
                            <div class="box-body">

                                <div class="box box-primary">
                                    <div class="box-header">
                                        <div class='row'>
                                            <div class='col-sm-10'>
                                                <h3 class="box-title">Lotações</h3>
                                            </div>
                                        </div>
                                    </div>
                                    <asp:GridView ID="grdPerfilLotacoes" runat="server" CssClass="table table-bordered" PagerStyle-CssClass="paginacao" AllowSorting="True" AllowPaging="True" PageSize="20" AutoGenerateColumns="False" DataKeyNames="">
                                        <HeaderStyle CssClass="bg-aqua" ForeColor="White" />
                                        <Columns>
                                             <asp:BoundField DataField="TG05_NM_REGIONAL" SortExpression="TG05_NM_REGIONAL" HeaderText="Regional" />
                                             <asp:BoundField DataField="TG03_NM_MUNICIPIO" SortExpression="TG03_NM_MUNICIPIO" HeaderText="Cidade" />
                                            <asp:BoundField DataField="RH36_NM_LOTACAO" SortExpression="RH36_NM_LOTACAO" HeaderText="Lotação" />
                                            <asp:BoundField DataField="RH36_CD_INEP_LOTACAO" SortExpression="RH36_CD_INEP_LOTACAO" HeaderText="Inep" />
                                            <asp:BoundField DataField="RH02_CD_MATRICULA" SortExpression="RH02_CD_MATRICULA" HeaderText="Matrícula" />
                                        </Columns>
                                    </asp:GridView>
                                </div>
                            </div>




                            <div class="box-body">

                                <div class="box box-primary">
                                    <div class="box-header">
                                        <div class='row'>
                                            <div class='col-sm-10'>
                                                <h3 class="box-title">Habilidades</h3>
                                            </div>
                                        </div>
                                    </div>
                                    <asp:GridView ID="grdPerfilHabilidades" runat="server" CssClass="table table-bordered" PagerStyle-CssClass="paginacao" AllowSorting="True" AllowPaging="True" PageSize="20" AutoGenerateColumns="False" DataKeyNames="">
                                        <HeaderStyle CssClass="bg-aqua" ForeColor="White" />
                                        <Columns>
                                            <asp:BoundField DataField="DE09_NM_DISCIPLINA" SortExpression="DE09_NM_DISCIPLINA" HeaderText="Componente Curricular" />
                                            <asp:BoundField DataField="TP_HABILIDADE" SortExpression="TP_HABILIDADE" HeaderText="Tipo Habilidade" />
                                            <asp:BoundField DataField="RH02_CD_MATRICULA" SortExpression="RH02_CD_MATRICULA" HeaderText="Matrícula" />
                                        </Columns>
                                    </asp:GridView>
                                </div>
                            </div>
                        </div>
                    </div>

                        <div class='row'>
                        <div class='col-sm-12'>
                            <div class="box-body">

                                <div class="box box-primary">
                                    <div class="box-header">
                                        <div class='row'>
                                            <div class='col-sm-10'>
                                                <h3 class="box-title">Alocação Carga Horária</h3>
                                            </div>
                                        </div>
                                    </div>
                                    <div class="box box-primary box-solid">
                                <div class="box-header">
                                    <div class="row">
                                        <div class="col-sm-2">
                                            <h3 class="box-title">Matrícula</h3>
                                        </div>
                                        <div class="col-sm-5" style="text-align: center">
                                            <h3 class="box-title">Lotação</h3>
                                        </div>
                                        <div class="col-sm-2" style="text-align: center">
                                            <h3 class="box-title">Turno</h3>
                                        </div>
                                        <div class="col-sm-3" style="text-align: center">
                                            <h3 class="box-title">Carga Horária</h3>
                                        </div>
                                    </div>
                                </div>
                                <cc1:Accordion ID="Accordion1" runat="server" CssClass="" AutoSize="None"
                                    FadeTransitions="true" TransitionDuration="250" FramesPerSecond="40" RequireOpenedPane="false" SuppressHeaderPostbacks="true"
                                    ContentCssClass="box-body" HeaderCssClass="box-header with-border pointer" HeaderSelectedCssClass="box-header with-border accordion-select">
                                    <HeaderTemplate>
                                        <div class="row">
                                            <div class="col-sm-2">
                                                <asp:Label ID="Label1" runat="server" Text='<%#DataBinder.Eval(Container.DataItem, "RH02_CD_MATRICULA") %>' />
                                            </div>
                                            <div class="col-sm-5" style="text-align: center">
                                                <asp:Label ID="lblLotacao" runat="server" Text='<%#DataBinder.Eval(Container.DataItem, "RH36_NM_LOTACAO") %>' />
                                            </div>
                                            <div class="col-sm-2" style="text-align: center">
                                                <asp:Label ID="lblTurno" runat="server" Text='<%#DataBinder.Eval(Container.DataItem, "TG06_NM_TURNO") %>' />
                                            </div>
                                            <div class="col-sm-3" style="text-align: center">
                                                <asp:Label ID="Label2" runat="server" Text='<%#DataBinder.Eval(Container.DataItem, "CARGA_HORARIA") %>' />
                                            </div>
                                        </div>
                                    </HeaderTemplate>
                                    <ContentTemplate>
                                        <asp:HiddenField ID="hdnCodigoAlocacao" runat="server" Value='<%#DataBinder.Eval(Container.DataItem, "RH80_ID_ALOCACAO_CARGA_HORARIA") %>' />
                                        <asp:GridView ID="grdhabilitacao" runat="server" CssClass="table table-bordered" PagerStyle-CssClass="paginacao" AutoGenerateColumns="false" DataKeyNames="RH74_ID_HABILITACAO">
                                            <HeaderStyle CssClass="bg-blue" ForeColor="White" BorderColor="#3c8dbc" />
                                            <Columns>
                                                <asp:BoundField DataField="DE09_NM_DISCIPLINA" HeaderText="Componente Curricular" ItemStyle-BorderColor="#3c8dbc" />
                                                <asp:BoundField DataField="RH74_QT_HORA_ALOCADA" HeaderText="Quantidade (h)" HeaderStyle-HorizontalAlign="Center" ItemStyle-HorizontalAlign="Center" ItemStyle-BorderColor="#3c8dbc" />
                                            </Columns>
                                        </asp:GridView>
                                    </ContentTemplate>
                                </cc1:Accordion>
                            </div>
                                </div>
                            </div>
                        </div>
                    </div>
                    <div class="box-footer">
                        <div class="btn-group">
                            <asp:LinkButton ID="btnVoltar" runat="server" class="btn btn-warning"><i class="fa fa-mail-reply"></i> Voltar</asp:LinkButton>
                        </div>
                    </div>

                </div>


            </section>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>

