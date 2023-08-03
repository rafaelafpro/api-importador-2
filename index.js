const express = require("express");
const fileUpload = require("express-fileupload");
const ExcelJS = require("exceljs");
const app = express();

app.get("/teste", (req, res) => {
    res.send("Bem-vindo à minha API!");
});

app.use(express.static("public"));

app.use(
    fileUpload({
        createParentPath: true,
    })
);

app.post("/upload", async (req, res) => {
    if (!req.files || Object.keys(req.files).length === 0) {
        return res.status(400).send("Nenhum arquivo foi enviado.");
    }

    // Leitura do arquivo
    let excelFile = req.files.excelFile;
    let workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(excelFile.data);

    // Processamento dos dados
    let acomodacoes = getAcomodacoes(workbook);
    let reservas = getReservas(workbook);
    let despesas = getDespesas(workbook);
    let movimentacoes = getMovimentacoes(workbook);

    // Resposta
    res.json({
        acomodacoes: acomodacoes,
        reservas: reservas,
        despesas: despesas,
        movimentacoes: movimentacoes,
    });
});

app.listen(3000, () => {
    console.log("Servidor escutando na porta 3000");
});

function getReservas(workbook) {
    var wsOrigem = workbook.getWorksheet("Dados Reservas");

    // Prepara o array de dados
    var dataArray = [];

    // Copia os dados da aba original para o array
    var lastRow = wsOrigem.lastRow.number;
    for (var i = 5; i <= lastRow; i++) {
        var row = wsOrigem.getRow(i);
        var rowData = row.values.slice(2, 11);

        // Verifica se a linha está vazia
        var isEmptyRow = rowData.every(function (value) {
            return value === undefined || value === "";
        });

        // Se a linha não estiver vazia, adiciona o objeto de linha ao array
        if (!isEmptyRow) {
            dataArray.push({
                RESERVA: `${rowData[0]}`,
                ACOMODACAO: `${rowData[1]}`,
                CANAL: rowData[2],
                QTDHOSPEDES: rowData[3],
                HOSPEDE: rowData[4],
                CHECKIN: rowData[5],
                CHECKOUT: rowData[6],
                VALORTOTAL: rowData[7],
                COMISSAO: rowData[8],
                VALORLIQUIDO: rowData[7] - rowData[8],
            });
        }
    }

    return dataArray;
}

function getAcomodacoes(workbook) {
    var wsOrigem = workbook.getWorksheet("INPUTS");

    var dataObject = {
        Importar_Acomodacoes_PROHOST: [],
    };

    var lastRow = wsOrigem.lastRow.number;
    for (var i = 3; i <= lastRow; i++) {
        var row = wsOrigem.getRow(i);
        var acomodacao = row.getCell(3).value;
        var status = row.getCell(4).value;
        var quantidade = row.getCell(5).value;
        var inicio = row.getCell(6).value;
        var valor = row.getCell(7).value;
        var comissao = row.getCell(23).value || "1";

        // Ignora linhas que não contêm dados
        if (
            acomodacao != null &&
            status != null &&
            quantidade != null &&
            inicio != null &&
            valor != null
        ) {
            dataObject["Importar_Acomodacoes_PROHOST"].push({
                ACOMODACAO: acomodacao,
                STATUS: status,
                "QUANTIDADE DE HOSPEDES": quantidade,
                "INICIO DA OPERACAO": inicio,
                "VALOR DO IMOVEL": valor,
                COMISSAO: comissao,
            });
        }
    }

    return dataObject;
}

function getDespesas(workbook) {
    var wsOrigemNegocio = workbook.getWorksheet("Despesas Negócio");
    var wsOrigemFixas = workbook.getWorksheet("Despesas Fixas");
    var wsOrigemVariaveis = workbook.getWorksheet("Despesas Variáveis");

    var dataObject = {
        Importar_Despesas_PROHOST: [],
    };

    // Extrai dados de Despesas Negócio
    var lastRowNegocio = wsOrigemNegocio.lastRow.number;
    for (var i = 5; i <= lastRowNegocio; i++) {
        // Ajuste '5' conforme necessário
        var row = wsOrigemNegocio.getRow(i);
        var dataDespesa = row.getCell(2).value;
        var tipoDespesa = row.getCell(3).value;
        var valorDespesa = row.getCell(4).value;

        // Ignora linhas que não contêm dados
        if (dataDespesa != null && tipoDespesa != null) {
            dataObject["Importar_Despesas_PROHOST"].push({
                DATADESPESA: dataDespesa,
                TIPODESPESA: tipoDespesa,
                VALORDESPESA: valorDespesa,
                ACOMODACAODESPESA: "NEGOCIO",
                CATDESPESA: "NEGOCIO",
            });
        }
    }

    // Extrai dados de Despesas Fixas
    var lastRowFixas = wsOrigemFixas.lastRow.number;
    for (var i = 5; i <= lastRowFixas; i++) {
        // Ajuste '5' conforme necessário
        var row = wsOrigemFixas.getRow(i);

        var dataDespesa = row.getCell(2).value;
        var acomodacaoDespesa = row.getCell(3).value;
        var tipoDespesa = row.getCell(4).value;
        var valorDespesa = row.getCell(5).value;

        // Ignora linhas que não contêm dados
        if (
            tipoDespesa != null &&
            valorDespesa != null &&
            acomodacaoDespesa != null
        ) {
            dataObject["Importar_Despesas_PROHOST"].push({
                DATADESPESA: dataDespesa,
                TIPODESPESA: tipoDespesa,
                VALORDESPESA: valorDespesa,
                ACOMODACAODESPESA: acomodacaoDespesa,
                CATDESPESA: "FIXA",
            });
        }
    }

    // Extrai dados de Despesas Variáveis
    var lastRowVariaveis = wsOrigemVariaveis.lastRow.number;
    for (var i = 5; i <= lastRowVariaveis; i++) {
        // Ajuste '5' conforme necessário
        var row = wsOrigemVariaveis.getRow(i);
        var dataDespesa = row.getCell(3).value;
        var acomodacaoDespesa = row.getCell(4).value;
        var tipoDespesa = row.getCell(5).value;
        var valorDespesa = row.getCell(7).value;

        // Ignora linhas que não contêm dados
        if (
            dataDespesa != null &&
            acomodacaoDespesa != null &&
            valorDespesa != null
        ) {
            dataObject["Importar_Despesas_PROHOST"].push({
                DATADESPESA: dataDespesa,
                TIPODESPESA: tipoDespesa,
                VALORDESPESA: valorDespesa,
                ACOMODACAODESPESA: acomodacaoDespesa,
                CATDESPESA: "VARIAVEL",
            });
        }
    }

    return dataObject;
}

function getMovimentacoes(workbook) {
    var wsOrigem = workbook.getWorksheet("Dados Reservas");

    var dataObject = {
        Importar_Movimentacoes_PROHOST: [],
    };

    var lastRow = wsOrigem.lastRow.number;
    for (var i = 5; i <= lastRow; i++) {
        var row = wsOrigem.getRow(i);
        var rowData = row.values.slice(2, 11);

        var reserva = rowData[0];
        var acomodacao = rowData[1];
        var canal = rowData[2];
        var checkin = new Date(rowData[5]);
        var checkout = new Date(rowData[6]);
        var valorTotal = rowData[7];
        var comissao = rowData[8];
        var valorLiquido = rowData[7] - rowData[8];
        var nome = rowData[4];

        var diarias = Math.round((checkout - checkin) / (1000 * 60 * 60 * 24));

        for (var j = 0; j < diarias; j++) {
            var data = new Date(checkin.getTime() + j * (1000 * 60 * 60 * 24));
            var valorDiaria = valorTotal / diarias;
            var valorLiq = valorLiquido / diarias;
            var com = comissao / diarias;

            dataObject["Importar_Movimentacoes_PROHOST"].push({
                ACOMODACAO: acomodacao,
                CANAL: canal,
                DATA: data,
                RESERVA: reserva,
                VALORTOTAL: valorDiaria,
                VALORLIQUIDO: valorLiq,
                COMISSAOOTA: com,
                NOME:
                    "Diária " +
                    (j + 1) +
                    "/" +
                    diarias +
                    " da Reserva " +
                    reserva,
            });
        }
    }

    return dataObject;
}
