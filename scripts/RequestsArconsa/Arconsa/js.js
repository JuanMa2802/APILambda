function ConsultaAjax(metodo, type, callback, parametros) {
    var _url = metodo;
    var XMLHttpRequest = require('xhr2');
    var req = new XMLHttpRequest();
    if (type == 'GET' || type == 'DELETE') {if (parametros != undefined) {
        if (parametros.length != undefined) {
            _url += '/' + encodeURIComponent(parametros);
        } else {
            _url += '/' + encodeURIComponent(parametros);
        }}}
    req.open(type, _url, true);
    //req.open(type, _url, async == undefined ? true : async);

    req.onreadystatechange = function (aEvt) {
        if (req.readyState == 4) {
            if (req.status == 200 || req.status == 204) {
                var _response = req.response;
                if (_response != '') {
                    _response = JSON.parse(req.response);
                }
                //   $('.progress-bar').hide();
                if (callback != undefined)
                    callback(_response);
            } 
        }
    };

    req.setRequestHeader("Authorization", process.argv[4]);
    if (type != 'GET') {
        req.setRequestHeader("Content-type", "application/json");

        if (parametros != undefined)
            parametros = JSON.stringify(parametros);
        req.send(parametros);
    } else {
        req.send(null);
        if (progressBar != undefined || progressBar != null) {
            $('.progress-bar').css('width', '10%');
        }
    }
}
function GenerarOC(idpedidos, obra) {
    let data = {};
    data.ObjProvSug = {};
    idpedidos = JSON.parse(idpedidos).toString()
    data.ObjProvSug.idpedidos = idpedidos;
    data.ObjProvSug.obra = JSON.parse(obra);
    ConsultaAjax("https://www1.sincoerp.com/sincoarconsa/V3/ADPRO/api/ComprarPedidos/GenerarOCPedSugerido", "PUT", function (source) {
            console.log(JSON.stringify({"status":source.ObjMensajes.codigo,"mensaje":source.ObjMensajes.mensaje}));
    }, data);
}

GenerarOC(process.argv[2],process.argv[3])