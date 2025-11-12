import { STRINGS } from "../messages.js";
import { getCell, asNumber } from "../parseUtils.js";

export function runStabilityChecks(workbook) {
  const feedback = [];
  let delta = 0;
  let stable = true;

  const main = workbook.sheets.main;

  const sm = asNumber(getCell(main, "M10"));
  const clb = asNumber(getCell(main, "O10"));
  const cnb = asNumber(getCell(main, "P10"));
  const ratio = asNumber(getCell(main, "Q10"));

  if (!(sm >= -0.1 && sm <= 0.11)) {
    feedback.push(STRINGS.stability.sm);
    stable = false;
  } else if (sm < 0) {
    feedback.push(STRINGS.stability.smWarn);
  }

  if (!(clb < -0.001)) {
    feedback.push(STRINGS.stability.clb);
    stable = false;
  }

  if (!(cnb > 0.002)) {
    feedback.push(STRINGS.stability.cnb);
    stable = false;
  }

  if (!(ratio >= -1 && ratio <= -0.3)) {
    feedback.push(STRINGS.stability.ratio);
    stable = false;
  }

  if (!stable) {
    feedback.push(STRINGS.stability.deduction);
    delta -= 1;
  }

  return { delta, feedback };
}
