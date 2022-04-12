<%@ Control Language="VB" AutoEventWireup="false" CodeFile="ctrRelatorio.ascx.vb" Inherits="ctrRelatorio" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="cc1" %>

<asp:Label ID="lblModal" runat="server" Text="" />

<cc1:ModalPopupExtender ID="mpeModal" runat="server" PopupControlID="pnlRelatorio" TargetControlID="lblModal" BackgroundCssClass="modalBackground" />                

<asp:Panel ID="pnlRelatorio" runat="server" CssClass="modalPanel" style="display:none;">
    <div class="modal-dialog modal-md">
        <div class="modal-content teste">
            <div class="modal-header">
                <h4 class="modal-title"><asp:Label ID="lblTitulo" runat="server" /></h4>
            </div>
            <div class="modal-body">
                <div class="row">
                    <div class="form-group col-sm-12 text-center">
                        <br />
                        <asp:LinkButton ID="lbtVisualizar" runat="server" CssClass="btn btn-primary"><span class="fa fa-eye"></span> Visualizar </asp:LinkButton>
                        <asp:LinkButton ID="lbtDownload" runat="Server" CssClass="btn btn-primary"><span class="fa fa-save"></span> Salvar </asp:LinkButton>
                        <asp:LinkButton ID="lbtFechar" runat="Server" CssClass="btn btn-primary"><span class="fa fa-mail-reply"></span> Voltar </asp:LinkButton>
                        <br />
                    </div>
                </div>
            </div>
            <div class="modal-footer">
                <div class="row">
                    <div class="form-group col-sm-12 text-left">
                        � necess�rio ter instalado no dispositivo um aplicativo para leitura de arquivos pdf
                    </div>
                </div>
            </div>
        </div>
    </div>
</asp:Panel>
