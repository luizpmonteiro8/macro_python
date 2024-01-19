import * as XLSX from "xlsx";

async function findNameByColumn(
  sheet: XLSX.WorkSheet,
  colunaProcurada: string,
  valorProcurado: string
): Promise<number | null> {
  let linhaEncontrada = null;

  // Verifica se a propriedade '!ref' existe antes de acessá-la
  if (sheet["!ref"]) {
    // Obtém o número da última linha
    const matchResult = sheet["!ref"].split(":")[1].match(/\d+/);
    const ultimaLinha = matchResult ? parseInt(matchResult[0], 10) : 0;

    // Itera sobre as linhas da coluna A
    for (let linha = 1; linha <= ultimaLinha; linha++) {
      const cellAddress = `${colunaProcurada}${linha}`;
      const cell = sheet[cellAddress];

      // Verifica se a célula contém um valor
      if (cell && cell.v !== undefined) {
        const textValue = cell.v.toString();

        if (textValue.includes(valorProcurado)) {
          linhaEncontrada = linha;
          break;
        }
      }
    }
  }

  return linhaEncontrada;
}

export default findNameByColumn;
