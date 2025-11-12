import { STRINGS } from "../messages.js";
import { getCell, asNumber } from "../parseUtils.js";

const THRUST_CELLS = [
  ["C48", "C49"],
  ["D48", "D49"],
  ["E48", "E49"],
  ["F48", "F49"],
  ["G48", "G49"],
  ["H48", "H49"],
  ["I48", "I49"],
  ["J48", "J49"],
  ["K48", "K49"],
  ["L48", "L49"],
  ["M48", "M49"],
  ["N48", "N49"],
];

export function runThrustAndTakeoff(workbook) {
  const feedback = [];
  let delta = 0;

  const miss = workbook.sheets.miss;
  const main = workbook.sheets.main;

  const thrustFail = THRUST_CELLS.some(([topRef, bottomRef]) => {
    const thrust = asNumber(getCell(miss, topRef));
    const drag = asNumber(getCell(miss, bottomRef));
    if (thrust == null || drag == null) {
      return false;
    }
    return thrust > drag;
  });

  if (thrustFail) {
    delta -= 1;
    feedback.push(STRINGS.thrustLeg);
  } else {
    const takeoffDistance = asNumber(getCell(main, "K38"));
    const takeoffRequired = asNumber(getCell(main, "X12"));
    if (
      takeoffDistance != null &&
      takeoffRequired != null &&
      takeoffDistance > takeoffRequired
    ) {
      delta -= 1;
      feedback.push(STRINGS.takeoffRoll);
    }
  }

  return { delta, feedback };
}
