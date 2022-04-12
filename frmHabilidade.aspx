﻿<%@ Page Title="" Language="VB" MasterPageFile="~/MasterPageCargaHoraria.master" AutoEventWireup="false" CodeFile="frmHabilidade.aspx.vb" Inherits="Habilidade" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="cc1" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
     <asp:UpdatePanel ID="UpdatePanel1" runat="server">
        <ContentTemplate>

            ELIELSON TOPPPP
            
            <section id="Listagem" runat="server"  class="content">
                <div class='row'>
                    <div class='col-sm-12'>
                        <div class='box box-blue'>
                            <div class="box-footer">
                                <asp:GridView ID="grdHabilidade" runat="server" CssClass="table table-bordered table-hover" PagerStyle-CssClass="paginacao" AllowSorting="True" AllowPaging="True" PageSize="20" AutoGenerateColumns="False" DataKeyNames="RH72_ID_HABILIDADE, Desativado, RH72_TP_HABILIDADE">
                                    <HeaderStyle CssClass="bg-aqua" ForeColor="White" />
                                    <Columns>
                                         <asp:ButtonField DataTextField="RH02_CD_MATRICULA" SortExpression="RH02_CD_MATRICULA" HeaderText="Matricula" />
                                         <asp:ButtonField DataTextField="TIPO_HABILIDADE" SortExpression="TIPO_HABILIDADE" HeaderText="Tipo" />
                                         <asp:ButtonField DataTextField="DE09_NM_DISCIPLINA" SortExpression="DE09_NM_DISCIPLINA" HeaderText="Disciplina" />
                                         <asp:TemplateField>
                                            <ItemStyle Width="120" HorizontalAlign="Right" />
                                                <ItemTemplate>
                                                    <div class="row">
                                                        <div class="col-sm-12">
                                                            <asp:LinkButton ID="lnkCancelar" runat="server" class="">
                                                                <i id="iCancelar" runat="server" class=""></i>
                                                                <asp:Label ID="lblCancelar" runat="server"></asp:Label>
                                                            </asp:LinkButton>
                                                        </div>
                                                    </div>
                                            </ItemTemplate>
                                         </asp:TemplateField>
                                    </Columns>
                                </asp:GridView>
                                <asp:Label ID="lblRegistros" runat="server" CssClass="badge bg-aqua" />
                            </div>
                        </div>
                    </div>
                </div>
            </section>
        </ContentTemplate>
     </asp:UpdatePanel>

    <%--Modal de confirmacao padrão--%>
    <div class="modal fade" id="modalSucesso">
        <div class="modal-dialog">
            <div class="modal-content">
                <div class="modal-header">
                    <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                        <span aria-hidden="true">×</span></button>
                    <h4 class="modal-title" id="txtHeaderModal"></h4>
                </div>
                <div class="modal-body">
                    <p class="h4">
                        <asp:Label runat="server" ID="lblMensagem"></asp:Label></p>
                    <p class="h5">
                        <asp:Label runat="server" ID="lblMensagem2"></asp:Label></p>
                </div>
                <div class="modal-footer">
                    <asp:LinkButton id="btnHabilidade" runat="server" visible="false" type="button"  class="btn btn-primary"> Ok</asp:LinkButton>
                    <asp:LinkButton id="btnEntendi" runat="server" class="btn btn-primary" data-dismiss="modal"> Entendi</asp:LinkButton>
                </div>
            </div>
        </div>
    </div>

</asp:Content>

