import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const inputPath = path.resolve("inputs/planilha_sem_titulo.xlsx");
const outputDir = path.resolve("outputs/guest_conversion");
const outputPath = path.join(outputDir, "convidados_formatados.xlsx");

function normalizeText(value) {
  if (value === null || value === undefined) return "";
  return String(value).replace(/\s+/g, " ").trim();
}

function cleanGuestName(value) {
  return normalizeText(value).replace(/\s*\(prima beth\)\s*/i, " ").replace(/\s+/g, " ").trim();
}

function looksLikeHeader(row) {
  const first = normalizeText(row[0]).toLowerCase();
  const second = normalizeText(row[1]).toLowerCase();
  return (
    first.includes("convite") ||
    first.includes("nome") ||
    second.includes("convidado") ||
    second.includes("nome")
  );
}

function pushGroup(groups, review, group) {
  if (!group) return;

  const uniqueGuests = [];
  for (const name of group.guests) {
    if (name && !uniqueGuests.some((existing) => existing.toLowerCase() === name.toLowerCase())) {
      uniqueGuests.push(name);
    }
  }

  if (uniqueGuests.length === 0) {
    review.push([group.invitationName, "", "Convite sem nomes de convidados na coluna B"]);
    return;
  }

  const representative = uniqueGuests[0];
  const companions = uniqueGuests.slice(1);
  groups.push({
    invitationName: group.invitationName,
    representative,
    companions,
  });

  if (normalizeText(group.invitationName) && representative) {
    const firstInvitationToken = normalizeText(group.invitationName).split(/\s+|,|;|&| e /i)[0]?.toLowerCase();
    const firstRepresentativeToken = representative.split(/\s+/)[0]?.toLowerCase();
    if (firstInvitationToken && firstRepresentativeToken && firstInvitationToken !== firstRepresentativeToken) {
      review.push([
        group.invitationName,
        representative,
        "Representante difere do primeiro nome do convite; conferir se esta correto",
      ]);
    }
  }
}

const input = await FileBlob.load(inputPath);
const sourceWorkbook = await SpreadsheetFile.importXlsx(input);
const summary = await sourceWorkbook.inspect({
  kind: "workbook,sheet,table",
  maxChars: 8000,
  tableMaxRows: 20,
  tableMaxCols: 6,
  tableMaxCellChars: 120,
});
console.log(summary.ndjson);

const sourceSheet = sourceWorkbook.worksheets.getItemAt(0);
const usedRange = sourceSheet.getUsedRange(true);
const rows = usedRange.values;

const groups = [];
const review = [];
let current = null;

for (let index = 0; index < rows.length; index += 1) {
  const row = rows[index];
  const invitationName = normalizeText(row[0]);
  const guestName = cleanGuestName(row[1]);

  if (index === 0 && looksLikeHeader(row)) {
    continue;
  }

  if (invitationName) {
    pushGroup(groups, review, current);
    current = { invitationName, guests: [] };
  }

  if (guestName) {
    if (!current) {
      current = { invitationName: "(sem nome de convite)", guests: [] };
      review.push(["", guestName, "Nome encontrado antes de qualquer nome de convite"]);
    }
    current.guests.push(guestName);
  }
}
pushGroup(groups, review, current);

const outputWorkbook = Workbook.create();
const formatted = outputWorkbook.worksheets.add("Convidados_Formatados");
const audit = outputWorkbook.worksheets.add("Revisao");

const formattedRows = [
  ["Nome_Completo", "Acompanhantes_Permitidos", "Nomes_Dos_Acompanhantes"],
  ...groups.map((group) => [
    group.representative,
    group.companions.length,
    group.companions.join("; "),
  ]),
];

formatted.getRangeByIndexes(0, 0, formattedRows.length, 3).values = formattedRows;
formatted.getRange("A1:C1").format = {
  fill: "#1F4E79",
  font: { bold: true, color: "#FFFFFF" },
};
formatted.getRangeByIndexes(0, 0, formattedRows.length, 3).format.borders = {
  preset: "inside",
  style: "thin",
  color: "#D9E2F3",
};
formatted.getRangeByIndexes(0, 0, formattedRows.length, 3).format.autofitColumns();
formatted.freezePanes.freezeRows(1);
formatted.showGridLines = false;

const reviewRows = [
  ["Nome_Do_Convite", "Nome_Representante_Ou_Convidado", "Observacao"],
  ...(review.length ? review : [["", "", "Nenhum ponto de revisao encontrado"]]),
];
audit.getRangeByIndexes(0, 0, reviewRows.length, 3).values = reviewRows;
audit.getRange("A1:C1").format = {
  fill: "#7F6000",
  font: { bold: true, color: "#FFFFFF" },
};
audit.getRangeByIndexes(0, 0, reviewRows.length, 3).format.autofitColumns();
audit.freezePanes.freezeRows(1);
audit.showGridLines = false;

const preview = await outputWorkbook.render({
  sheetName: "Convidados_Formatados",
  autoCrop: "all",
  scale: 1,
  format: "png",
});

await fs.mkdir(outputDir, { recursive: true });
await fs.writeFile(path.join(outputDir, "preview_convidados_formatados.png"), new Uint8Array(await preview.arrayBuffer()));

const check = await outputWorkbook.inspect({
  kind: "table",
  sheetId: "Convidados_Formatados",
  range: `A1:C${Math.min(formattedRows.length, 20)}`,
  include: "values",
  tableMaxRows: 20,
  tableMaxCols: 3,
  maxChars: 6000,
});
console.log(check.ndjson);

const output = await SpreadsheetFile.exportXlsx(outputWorkbook);
await output.save(outputPath);
console.log(JSON.stringify({ outputPath, groups: groups.length, reviewItems: review.length }));
