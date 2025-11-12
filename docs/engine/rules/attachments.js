import { STRINGS } from "../messages.js";
import { getCell, getCellByIndex, asNumber } from "../parseUtils.js";
import { format } from "../format.js";

const DEG_TO_RAD = Math.PI / 180;

export function runAttachmentChecks(workbook) {
  const feedback = [];
  let delta = 0;
  let disconnected = 0;

  const main = workbook.sheets.main;
  const geom = workbook.sheets.geom;

  const fuselageLength = asNumber(getCell(main, "B32"));
  const pcsArea = asNumber(getCell(main, "C18"));
  const pcsX = asNumber(getCell(main, "C23"));
  const pcsRootChord = asNumber(getCell(geom, "C8"));
  if (
    pcsArea != null &&
    pcsArea >= 1 &&
    pcsX != null &&
    pcsRootChord != null &&
    fuselageLength != null &&
    pcsX > fuselageLength - 0.25 * pcsRootChord
  ) {
    feedback.push(STRINGS.attachment.pcsX);
    disconnected += 1;
  }

  const vtArea = asNumber(getCell(main, "H18"));
  const vtX = asNumber(getCell(main, "H23"));
  const vtRootChord = asNumber(getCell(geom, "C10"));
  if (
    vtArea != null &&
    vtArea >= 1 &&
    vtX != null &&
    vtRootChord != null &&
    fuselageLength != null &&
    vtX > fuselageLength - 0.25 * vtRootChord
  ) {
    feedback.push(STRINGS.attachment.vtX);
    disconnected += 1;
  }

  const pcsZ = asNumber(getCell(main, "C25"));
  const fuseZCenter = asNumber(getCell(main, "D52"));
  const fuseZHeight = asNumber(getCell(main, "F52"));
  if (
    pcsArea != null &&
    pcsArea >= 1 &&
    pcsZ != null &&
    fuseZCenter != null &&
    fuseZHeight != null &&
    (pcsZ < fuseZCenter - fuseZHeight / 2 || pcsZ > fuseZCenter + fuseZHeight / 2)
  ) {
    feedback.push(STRINGS.attachment.pcsZ);
    disconnected += 1;
  }

  const vtY = asNumber(getCell(main, "H24"));
  const fuseWidth = asNumber(getCell(main, "E52"));
  if (
    vtArea != null &&
    vtArea >= 1 &&
    vtY != null &&
    fuseWidth != null &&
    vtY > fuseWidth / 2
  ) {
    feedback.push(STRINGS.attachment.vtY);
    disconnected += 1;
  }

  const strakeArea = asNumber(getCell(main, "D18"));
  if (strakeArea != null && strakeArea >= 1) {
    const sweep = asNumber(getCell(geom, "K15"));
    const y = asNumber(getCell(geom, "M152"));
    const strake = asNumber(getCell(geom, "L155"));
    const apex = asNumber(getCell(geom, "L38"));
    if (
      sweep != null &&
      y != null &&
      strake != null &&
      apex != null
    ) {
      const wing = y / Math.tan((90 - sweep) * DEG_TO_RAD) + apex;
      if (!(wing < strake + 0.5)) {
        feedback.push(STRINGS.attachment.strake);
        disconnected += 1;
      }
    }
  }

  if (fuselageLength != null) {
    const activeComponentPositions = [];
    for (let col = 2; col <= 8; col += 1) {
      const area = asNumber(getCellByIndex(main, 18, col));
      const position = asNumber(getCellByIndex(main, 23, col));
      if (area != null && area >= 1 && position != null) {
        activeComponentPositions.push(position);
      }
    }
    if (activeComponentPositions.length > 0) {
      const hasBehind = activeComponentPositions.some((value) => value >= fuselageLength);
      if (hasBehind) {
        feedback.push(
          format(STRINGS.attachment.fuselage, fuselageLength)
        );
        delta -= 1;
      }
    }
  }

  if (disconnected > 0) {
    feedback.push(STRINGS.attachment.deduction);
    delta -= 1;
  }

  return { delta, feedback };
}
