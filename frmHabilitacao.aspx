<%@ Page Title="" Language="VB" MasterPageFile="~/MasterPageCargaHoraria.master" AutoEventWireup="false" CodeFile="frmHabilitacao.aspx.vb" Inherits="frmHabilitacao" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="cc1" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
	<asp:UpdatePanel ID="UpdatePanel1" runat="server">
		<ContentTemplate>
			<section id="Cadastro" runat="server" class="content">
				<asp:Panel ID="pnlCadastro" runat="server">
					<div class='row'>
						 <div class="col-sm-4">
							 Período<br; />
							 <asp:DropDownList ID="drpPeriodo" runat="server"  AutoPostBack="true" class="form-control" />
						 </div>
					 </div>
					<div class="row">
						<div class="form-group  col-sm-6">
							Servidor<br />
							<asp:DropDownList ID="drpServidor" runat="server" Enabled="false" CssClass="form-control" AutoPostBack="true"/>
						</div>
					    <div class="form-group  col-sm-6">
					        Lotação<br />
					        <asp:DropDownList ID="drpLotacao" runat="server" Enabled="false" CssClass="form-control" AutoPostBack="true"/>
					    </div>
					</div>
					<div class="row">
						<div class="form-group  col-sm-12">
							Carga Horária<br />
							<asp:DropDownList runat="server" ID="drpCargaHoraria"  Enabled="false"  CssClass="form-control"  AutoPostBack="true" />
						</div>
					</div>
					<div class="row">
						<div class="form-group  col-sm-12">
							Alocação<br />
							<asp:DropDownList runat="server" ID="drpAlocacao" Enabled="false"  CssClass="form-control" AutoPostBack="true" />
						</div>
					</div>
					<div class="row">
						<div class="form-group  col-sm-10">
							Habilidade<br />
							<asp:DropDownList runat="server" ID="drphabilidade" Enabled="false"  CssClass="form-control" AutoPostBack="true" />
						</div>
						<div class="form-group  col-sm-2">
							Horas<br />
							<asp:TextBox runat="server" MaxLength="2" ID="txtHoraAlocada" CssClass="form-control"  onkeyup="formataInteiro(this,event);"  />
						</div>
					</div>
					<div class="box-footer">
						<div class="btn-group">
							<asp:LinkButton ID="btnNovo" runat="server" class="btn btn-info"><i class="fa fa-mail-reply"></i>&nbsp;Novo</asp:LinkButton>
						</div>
						<div class="btn-group">
							<asp:LinkButton ID="btnSalvar" runat="server" class="btn btn-success"><i class="fa fa-save"></i>&nbsp; Salvar</asp:LinkButton>
						</div>
						<div class="btn-group">
							<asp:LinkButton ID="btnLocalizar" runat="server" class="btn btn-default"><i class="fa fa-search"></i>&nbsp; Localizar</asp:LinkButton>
						</div>
					</div>
					<div class="box-footer">
						<asp:Label runat="server" ID="lblHoraAula" CssClass="badge bg-aqua-gradient" />
						<asp:Label runat="server" ID="lblHoraAulaRestante" CssClass="badge bg-red-gradient" />
					</div>
				</asp:Panel>
			</section>  
			<section id="Listagem" runat="server" class="content">
				<!-- Small boxes (Stat box) -->
				<div class='row'>
					<div class='col-sm-12'>
						<div class='box box-blue'>
							<div class="box-footer">
								<asp:GridView ID="grdhabilitacao" runat="server" CssClass="table table-bordered table-hover" PagerStyle-CssClass="paginacao" AllowSorting="True" AllowPaging="True" PageSize="20" AutoGenerateColumns="False" DataKeyNames="RH74_ID_HABILITACAO,Desativado">
									<HeaderStyle CssClass="bg-aqua" ForeColor="White" />
								    <FooterStyle CssClass="bg-aqua" ForeColor="White" />
									<Columns>
                                        <asp:ButtonField DataTextField="RH80_ID_ALOCACAO_CARGA_HORARIA" SortExpression="RH80_ID_ALOCACAO_CARGA_HORARIA" HeaderText="Código" />
										<asp:ButtonField DataTextField="RH02_CD_MATRICULA" SortExpression="RH02_CD_MATRICULA" HeaderText="Matricula" />
										<asp:ButtonField DataTextField="TG06_NM_TURNO" SortExpression="TG06_NM_TURNO" HeaderText="Turno" />
										<asp:ButtonField DataTextField="TIPO_HABILIDADE" SortExpression="TIPO_HABILIDADE" HeaderText="Tipo" />
										<asp:ButtonField DataTextField="DE09_NM_DISCIPLINA" SortExpression="DE09_NM_DISCIPLINA" HeaderText="Disciplina" />
										<asp:ButtonField DataTextField="RH36_NM_LOTACAO" SortExpression="RH36_NM_LOTACAO" HeaderText="Lotação" />
										<asp:ButtonField DataTextField="RH80_QT_HORA_ALOCADA" SortExpression="RH80_QT_HORA_ALOCADA" HeaderText="C.H Alocada" />
										<asp:ButtonField DataTextField="RH74_QT_HORA_ALOCADA" SortExpression="RH74_QT_HORA_ALOCADA" HeaderText="C.H Distribuida" />
										<asp:ButtonField DataTextField="RH79_NM_TIPO_CARGA_HORARIA" SortExpression="RH79_NM_TIPO_CARGA_HORARIA" HeaderText="Carga Horária" />
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
<%--										<asp:ButtonField Text="DESATIVAR" CommandName="Desativar" />--%>
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

</asp:Content>

