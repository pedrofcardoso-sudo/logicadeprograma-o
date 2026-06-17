const XLSX = require("xlsx");
const fs = require("fs");


function calcularSalarioFinal(salarioBase, horasExtras) {
    return salarioBase + (horasExtras * 30);
}


function classificarFuncionario(salarioFinal) {
    if (salarioFinal < 3000) return "Júnior";
    else if (salarioFinal < 4000) return "Pleno";
    else return "Sênior";
}


const workbook = XLSX.readFile("funcionarios.xlsx");
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const funcionarios = XLSX.utils.sheet_to_json(sheet);


let resultados = [];
let totalFolha = 0;
let maiorSalario = 0;
let menorSalario = Infinity;
let contagem = { Júnior: 0, Pleno: 0, Sênior: 0 };

funcionarios.forEach(f => {
    const salarioFinal = calcularSalarioFinal(f.Salário, f.HorasExtras);
    const classificacao = classificarFuncionario(salarioFinal);

    resultados.push({
        Nome: f.Nome,
        SalarioFinal: salarioFinal,
        Classificacao: classificacao
    });

    totalFolha += salarioFinal;
    if (salarioFinal > maiorSalario) maiorSalario = salarioFinal;
    if (salarioFinal < menorSalario) menorSalario = salarioFinal;
    contagem[classificacao]++;
});


console.table(resultados);


console.log("Folha de pagamento total: R$", totalFolha.toFixed(2));
console.log("Maior salário final: R$", maiorSalario.toFixed(2));
console.log("Menor salário final: R$", menorSalario.toFixed(2));
console.log("Média salarial: R$", (totalFolha / funcionarios.length).toFixed(2));
console.log("Quantidade por classificação:", contagem);


let relatorio = "Relatório de Funcionários\n\n";
resultados.forEach(r => {
    relatorio += `${r.Nome} - R$ ${r.SalarioFinal.toFixed(2)} - ${r.Classificacao}\n`;
});
relatorio += `\nFolha total: R$ ${totalFolha.toFixed(2)}\n`;
relatorio += `Maior salário: R$ ${maiorSalario.toFixed(2)}\n`;
relatorio += `Menor salário: R$ ${menorSalario.toFixed(2)}\n`;
relatorio += `Média salarial: R$ ${(totalFolha / funcionarios.length).toFixed(2)}\n`;
relatorio += `Classificações: ${JSON.stringify(contagem)}\n`;

fs.writeFileSync("relatorio.txt", relatorio);
