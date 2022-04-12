



/*----------- JAVASCRIPT EXTRA--------------*/

$(function () {
    //Initialize Select2 Elements
    $(".select2").select2();
});

function CarregarSelect2() {
    $(".select2").select2();
    return false;
}


function openModal(tipo = "") {
    if (tipo == "sucesso") {
        txtHeaderModal.innerHTML = "<i class='fa fa-check'></i> Sucesso!";
        $('#modalSucesso').modal('show');
    } else if (tipo == "erro") {
        txtHeaderModal.innerHTML = "<i class='fa fa-close'></i> Erro!";
        $('#modalSucesso').modal('show');
    } else if (tipo == "aviso") {
        txtHeaderModal.innerHTML = "<i class='fa fa-exclamation-circle'></i> Atenção!"
        $('#modalSucesso').modal('show');
    }
}


function OpenFile(file, extension, tipoDocumento) {
    let pdfWindow = window.open("")

    let contentType

    if (extension == ".pdf") {
        contentType = "data:application/pdf;"
    } else if (extension == ".jpg" || extension == ".jpeg") {
        contentType = "data:image/jpeg;"
    }
    pdfWindow.document.write =
        pdfWindow.document.write(
            "<title>" + tipoDocumento + "</title>" +
            "<iframe width='100%' height='100%' title=" + tipoDocumento + " src='" + contentType + "base64, " +
            encodeURI(file) + "'></iframe>"
        )
}

function getText(txtbox, e) {
    var maxlength = 500;
    var keyCode;
    if (window.event)
        keyCode = window.event.keyCode;
    else
        keyCode = e.which;

    switch (keyCode) {
        case 8:
            return true;
        default:
            if (txtbox.value.length == maxlength) {
                return false;
            }
            else {
                setText(txtbox);
            }
    }
    return true;
}