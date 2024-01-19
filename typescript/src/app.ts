import * as ExcelJS from "excel4node";
import express, { Express, Request, Response } from "express";
import fs from "fs";
import multer, { Multer } from "multer";
import path from "path";
import { copyColumnToAnother } from "./functions/copyColumnToAnother";
import { findStartEndRow } from "./functions/findStartEndRow";

const app: Express = express();
const PORT: number = 3000;
const rootDirectory = path.resolve(__dirname, "..");

// Configuração do Multer para lidar com uploads de arquivos
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, path.join(rootDirectory, "./excel-files")); // Pasta de destino para os arquivos Excel
  },
  filename: (req, file, cb) => {
    cb(null, file.originalname);
  },
});

const upload: Multer = multer({
  storage: storage,
  limits: { fileSize: 30 * 1024 * 1024 }, // 30MB em bytes
});

app.use(express.static(path.join(rootDirectory, "../public")));

// Configuração CORS
app.use((req: Request, res: Response, next) => {
  res.header("Access-Control-Allow-Origin", "*");
  res.header("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.header("Access-Control-Allow-Headers", "Content-Type");
  next();
});

app.get("/", (req: Request, res: Response) => {
  res.sendFile("index.html", { root: path.join(rootDirectory, "/public") });
});
app.get("/script.js", (req: Request, res: Response) => {
  res.sendFile("script.js", { root: path.join(rootDirectory, "/public") });
});

app.get("/config-files", (req: Request, res: Response) => {
  const configDirectory: string = path.join(__dirname, "../config");

  try {
    const files: string[] = fs.readdirSync(configDirectory);
    res.send({
      success: true,
      files,
    });
  } catch (error) {
    console.error(
      `Erro ao obter a lista de arquivos de configuração: ${error}`
    );
    res.status(500).send({
      success: false,
      error: "Erro ao obter a lista de arquivos de configuração.",
    });
  }
});

// Rota para processar o arquivo Excel usando Multer
app.post(
  "/process-excel",
  upload.single("excelFile"),
  async (req: Request, res: Response) => {
    try {
      const excelFile: Express.Multer.File = req.file as Express.Multer.File;
      const configFileName: string = req.headers["configfile"] as string;

      if (excelFile && configFileName) {
        const storedFilePath: string = path.join(
          rootDirectory,
          `./excel-files/${excelFile.originalname}`
        );

        // Lê o arquivo de configuração
        const configPath = path.join(rootDirectory, "config", configFileName);

        const configData = fs.readFileSync(configPath, "utf-8");
        const config = JSON.parse(configData);

        // Stream do arquivo
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(storedFilePath);

        const { initialRow, finalRow } = await findStartEndRow(
          workbook,
          config
        );

        await copyColumnToAnother(
          workbook,
          config.planilha,
          "G",
          "K",
          initialRow,
          finalRow,
          storedFilePath
        );

        res.status(200).send({
          success: true,
          initialRow,
          finalRow,
        });
      } else {
        res.status(400).send({
          success: false,
          error: "Arquivo Excel ou nome do arquivo de configuração ausente.",
        });
      }
    } catch (error: any) {
      console.error(`Erro ao processar o arquivo: ${error.message}`);
      res.status(400).send({
        success: false,
        error: error.message,
      });
    }
  }
);

app.listen(PORT, () => {
  console.log(`Servidor rodando em http://localhost:${PORT}`);
});
