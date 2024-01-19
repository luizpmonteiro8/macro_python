import * as XLSX from "xlsx";

/**
 * Copia os valores de uma coluna para outra em um intervalo específico e salva o arquivo.
 * @param workbook O livro de trabalho (workbook) contendo a planilha.
 * @param worksheetName O nome da planilha dentro do workbook.
 * @param sourceColumn A letra da coluna de origem.
 * @param targetColumn A letra da coluna de destino.
 * @param initialRow A primeira linha do intervalo.
 * @param finalRow A última linha do intervalo.
 * @param filePath O caminho do arquivo onde o workbook será salvo.
 */

export async function copyColumnToAnother(
  workbook: XLSX.WorkBook,
  worksheetName: string,
  sourceColumn: string,
  targetColumn: string,
  initialRow: number,
  finalRow: number,
  filePath: string
): Promise<void> {
  try {
    // Verifica se a planilha existe no workbook
    const worksheet: XLSX.WorkSheet | undefined =
      workbook.Sheets[worksheetName];

    if (!worksheet) {
      throw new Error(`A planilha "${worksheetName}" não foi encontrada.`);
    }

    // Copia os valores da coluna de origem para a coluna de destino no intervalo especificado
    for (let i = initialRow; i <= finalRow; i++) {
      const sourceCell = worksheet[`${sourceColumn}${i}`];

      // Certifique-se de que a célula de origem exista e tenha um valor
      if (sourceCell && sourceCell.v !== undefined) {
        worksheet[`${targetColumn}${i}`] = { ...sourceCell };
      }
    }

    // Salva o workbook utilizando a função customizada
    await XLSX.writeFile(workbook, filePath, {
      compression: true,
      cellStyles: true,
      cellDates: true,
      bookSST: true,
      bookType: "xlsx",
    });
  } catch (error) {
    console.error(`Erro ao salvar o arquivo: ${error}`);
    throw new Error(`Erro ao salvar o arquivo: ${error}`);
  }
}
