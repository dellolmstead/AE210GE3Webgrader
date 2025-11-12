import { STRINGS } from "../messages.js";
import { getCell, asNumber } from "../parseUtils.js";
import { format } from "../format.js";

export function runLandingGearChecks(workbook) {
  const feedback = [];
  let good = true;

  const gear = workbook.sheets.gear;

  const noseRule = asNumber(getCell(gear, "J19"));
  if (Number.isFinite(noseRule) && (noseRule < 9.5 || noseRule > 20)) {
    feedback.push(format(STRINGS.gear.nose, noseRule));
    good = false;
  }

  // Tip-back: MATLAB requires Gear(19,12) < Gear(20,12) (upper angle less than lower) so it won't tail sit.
  // On the template, the numeric values live at L20 (upper) and L21 (lower).
  const tipbackUpper = asNumber(getCell(gear, "L20"));
  const tipbackLower = asNumber(getCell(gear, "L21"));
  if (Number.isFinite(tipbackUpper) && Number.isFinite(tipbackLower) && !(tipbackUpper < tipbackLower)) {
    feedback.push(format(STRINGS.gear.tipback, tipbackUpper, tipbackLower));
    good = false;
  }

  // Rollover: MATLAB requires Gear(19,13) < Gear(20,13) to avoid rollover. Values sit at M20 and M21.
  const rolloverUpper = asNumber(getCell(gear, "M20"));
  const rolloverLower = asNumber(getCell(gear, "M21"));
  if (Number.isFinite(rolloverUpper) && Number.isFinite(rolloverLower) && !(rolloverUpper < rolloverLower)) {
    feedback.push(format(STRINGS.gear.rollover, rolloverUpper, rolloverLower));
    good = false;
  }

  const rotationAuthority = asNumber(getCell(gear, "N20"));
  const takeoffSpeed = asNumber(getCell(gear, "N21"));
  if (!Number.isFinite(rotationAuthority) || !Number.isFinite(takeoffSpeed)) {
    feedback.push(STRINGS.gear.rotationData);
    good = false;
  } else {
    if (!(rotationAuthority < takeoffSpeed)) {
      feedback.push(format(STRINGS.gear.rotationAuthority, rotationAuthority, takeoffSpeed));
      good = false;
    }
    if (takeoffSpeed > 200) {
      feedback.push(format(STRINGS.gear.takeoffSpeed, takeoffSpeed));
      good = false;
    }
  }

  let delta = 0;
  if (!good) {
    feedback.push(STRINGS.gear.deduction);
    delta -= 1;
  }

  return { delta, feedback };
}
