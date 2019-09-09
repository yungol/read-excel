const express = require('express');
const XLSX = require("xlsx");
const app = express();
const fs = require('fs');


function agrupar() {
    const buf = fs.readFileSync("ventas.xlsx");

    const wb = XLSX.read(buf, {
        type: 'buffer'
    });
    // Toma la primera hoja
    const wsname = wb.SheetNames[0];
    const ws = wb.Sheets[wsname];
    /* toma la primera fila y genera los nombres de las columnas */
    const aoa = XLSX.utils.sheet_to_json(ws, {
        header: 1,
        raw: false
    });
    let cols = [];
    for (var i = 0; i < aoa[0].length; ++i)
        cols[i] = {
            field: aoa[0][i]
        };
    /* se genera el resto de los datos */
    let data = [];
    for (let r = 1; r < aoa.length; ++r) {
        data[r - 1] = {};
        for (i = 0; i < aoa[r].length; ++i) {
            if (aoa[r][i] === null) continue;
            data[r - 1][aoa[0][i]] = aoa[r][i];
        }
    }
    // console.log(data);

    const columnasexcel = [];
    const variables = [];
    Object.keys(data[0]).forEach(function (k, v) {
        const d = {
            codigo: k
        };
        if (isNaN(parseInt(data[0][k]))) {
            d.clasif = "grupo";
        } else {
            d.clasif = "variable";
            variables.push(k);
        }
        columnasexcel.push(d);
    });

    // ========= aqui empieza

    const gruposkey = [];
    const datoskey = [];
    const columnas = Object.keys(data[0]);
    columnas.forEach(function (k, v) {
        if (variables.indexOf(k) === -1) {
            gruposkey.push(k);
        } else {
            datoskey.push(k);
        }
    });
    const grupos = {};
    const groups1 = {};
    gruposkey.forEach(function (k, v) {
        groups1[k] = [];
        grupos[k] = {
            subgrupos: [],
            sub: {}
        }
    });
    data.forEach(function (k, v) {
        gruposkey.forEach(function (k1, v1) {
            if (grupos[k1].subgrupos.indexOf(k[k1]) === -1) {
                grupos[k1].subgrupos.push(k[k1]);
                grupos[k1].sub[k[k1]] = [];
                grupos[k1].sub[k[k1]].push(k);
                grupos[k1].total = {};
            } else {
                grupos[k1].sub[k[k1]].push(k);
            }
        });
    });
    gruposkey.forEach(function (k, v) {
        grupos[k].subgrupos.forEach(function (k1, v1) {
            grupos[k].total[k1] = {
                var: {},
                vars: []
            }
            grupos[k].sub[k1].forEach(function (k2, v2) {
                datoskey.forEach(function (k3, v3) {
                    if (grupos[k].total[k1].vars.indexOf(k3) === -1) {
                        grupos[k].total[k1].vars.push(k3);
                        grupos[k].total[k1].var[k3] = 0;
                        grupos[k].total[k1].var[k3] += parseFloat(k2[k3]);
                    } else {
                        grupos[k].total[k1].var[k3] += parseFloat(k2[k3]);
                    }
                });
            });
        });
    });
    gruposkey.forEach(function (k, v) {
        grupos[k].subgrupos.forEach(function (k1, v1) {
            groups1[k][v1] = {};
            groups1[k][v1].Grupo = k;
            groups1[k][v1].SubGrupo = k1;
            groups1[k][v1].check = false;
            groups1[k][v1].check2 = false;
            grupos[k].total[k1].vars.forEach(function (k2, v2) {
                groups1[k][v1][k2] = {
                    'sum': grupos[k].total[k1].var[k2]
                };
            });
        });
    });
    groups1.grupos = gruposkey;
    groups1.variables = [];
    variables.forEach(function (k) {
        const d = {
            cod: k,
            check: false
        };
        groups1.variables.push(d);
    });

    console.log(groups1);
    return groups1;
}

app.get('/', function (req, res) {
    const data = agrupar();
    res.send(data);
});

app.listen(3000, function () {
    console.log('Servidor servido en el puerto 3000!');
});