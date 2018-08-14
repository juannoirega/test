$(document).ready(function(){

    alterarLayout();
    validacionCampos();

});
let valoresEmail = [];

function validacionCampos(){
    

}

function alterarLayout(){

    criacionColum();
    ciracionGridRight();
    $('.tipo_de_linea, .numero_de_poliza, .tipo_proceso').addClass('col-md-12');
    $('.tipo_de_linea, .numero_de_poliza, .tipo_proceso').find('label').addClass('divTitle col-md-12');
    $('.email_dirigido, .cuerpo_de_email, .fecha_hora_de_email, .email_en_copia, .asunto_de_email, .email_solicitante').hide();

}
function ciracionGridRight(){
    getValoresEmail();
    $('[name=ColumRight]').append('<div class="row" name="GridRight"></div>');
    $('[name=GridRight]').append(`<div class="text col-md-12">De : ${ valoresEmail [0]} </div>`);
    $('[name=GridRight]').append(`<div class="text col-md-12">Para :  ${valoresEmail[1]}</div>`);
    $('[name=GridRight]').append(`<div class="text col-md-12">Co : ${valoresEmail[2]}</div>`);
    $('[name=GridRight]').append(`<div class="text col-md-12">Fecha : ${valoresEmail[3]}</div>`);
    $('[name=GridRight]').append(`<div class="text col-md-12">Titulo : ${valoresEmail[4]}</div>`);
    $('[name=GridRight]').append(`<div class="text col-md-12">Documentos Adjuntos :</div>`);
    $('[name=GridRight]').append(`<div class="text col-md-12">Contenido : ${valoresEmail[5]} </div>`);

}
function getValoresEmail(){
  
    valoresEmail [0] = $('[name=email_solicitante]').val();
    valoresEmail [1] = $('[name=email_dirigido]').val();
    valoresEmail [2] = $('[name=email_en_copia]').val();
    valoresEmail [3] = $('[name=fecha_hora_de_email]').val();
    valoresEmail [4] = $('[name=asunto_de_email]').val();
    valoresEmail [5] = $('[name=cuerpo_de_email]').val();

}
function criacionColum(){

    $('.tipo_de_linea, .numero_de_poliza, .tipo_proceso').wrapAll('<div class="col-md-6"  name="ColumLeft"></div>');
    $('[name=ColumLeft]').wrapAll('<div name="Grid" class="row"></div>');
    $('[name=Grid]').append('<div class=" divContainer col-md-6"  name="ColumRight"></div>');
    $('[name=ColumRight]').append('<div class="divTitle col-md-12" name="divTitle">Datos del correo</div>');
    
}