import * as ExcelJS from "excel4node";
import findNameByColumn from "./findNameByColumn";

export const findStartEndRow = async (
  workbook: ExcelJS.Workbook,
  config: {
    planilha: string;
    colunaInicial: string;
    valorInicial: string;
    colunaFinal: string;
    valorFinal: string;
  }
): Promise<{
  initialRow: number;
  finalRow: number;
}> => {
  const sheetName = config.planilha;

  // Encontra a planilha no workbook
  const sheet = workbook.getWorksheet(sheetName);
  if (!sheet) {
    throw new Error(`A planilha "${sheetName}" não foi encontrada.`);
  }

  // Use os valores do arquivo de configuração
  let initialRow = await findNameByColumn(
    sheet,
    config.colunaInicial,
    config.valorInicial
  );

  if (initialRow == null) {
    throw new Error(
      `Valor inicial - ${config.valorInicial} não encontrado na planilha.`
    );
  } else {
    initialRow = initialRow + 1;
  }

  let finalRow = await findNameByColumn(
    sheet,
    config.colunaFinal,
    config.valorFinal
  );

  if (finalRow == null) {
    throw new Error(
      `Valor final - ${config.valorFinal} não encontrado na planilha.`
    );
  } else {
    finalRow = finalRow - 1;
  }

  return { initialRow, finalRow };
};
