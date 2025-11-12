import { STRINGS } from "../messages.js";
import { getCell, getCellByIndex, asNumber } from "../parseUtils.js";

const LEG_COLUMNS = [11, 12, 13, 14, 16, 18, 19, 22, 23]; // 1-based column indices
const ROWS = {
  altitude: 33,
  mach: 35,
  afterburner: 36,
  distance: 38,
  time: 39,
};

export function runMissionChecks(workbook) {
  const feedback = [];
  let missionFailed = false;

  const main = workbook.sheets.main;
  const constraintsMach = asNumber(getCell(main, "U4"));

  const readRowValues = (rowIndex) =>
    LEG_COLUMNS.map((col) => asNumber(getCellByIndex(main, rowIndex, col)));

  const altitude = readRowValues(ROWS.altitude);
  const mach = readRowValues(ROWS.mach);
  const afterburner = readRowValues(ROWS.afterburner);
  const distance = readRowValues(ROWS.distance);
  const time = readRowValues(ROWS.time);

  if (altitude[0] !== 0 || afterburner[0] !== 100) {
    feedback.push(STRINGS.missionLegs[0]);
    missionFailed = true;
  }

  if (!(altitude[1] >= altitude[0] && altitude[1] <= altitude[2])) {
    feedback.push(STRINGS.missionLegs[1]);
    missionFailed = true;
  }
  if (!(mach[1] >= mach[0] && mach[1] <= mach[2])) {
    feedback.push(STRINGS.missionLegs[2]);
    missionFailed = true;
  }
  if (afterburner[1] !== 0) {
    feedback.push(STRINGS.missionLegs[3]);
    missionFailed = true;
  }

  if (altitude[2] < 35000 || mach[2] !== 0.9 || afterburner[2] !== 0) {
    feedback.push(STRINGS.missionLegs[4]);
    missionFailed = true;
  }

  if (altitude[3] < 35000 || mach[3] !== 0.9 || afterburner[3] !== 0) {
    feedback.push(STRINGS.missionLegs[5]);
    missionFailed = true;
  }

  if (
    altitude[4] < 35000 ||
    constraintsMach == null ||
    Math.abs(mach[4] - constraintsMach) > 0.01 ||
    afterburner[4] !== 0 ||
    distance[4] < 150
  ) {
    feedback.push(STRINGS.missionLegs[6]);
    missionFailed = true;
  }

  if (altitude[5] < 30000 || mach[5] < 1.2 || afterburner[5] !== 100 || time[5] < 2) {
    feedback.push(STRINGS.missionLegs[7]);
    missionFailed = true;
  }

  if (
    altitude[6] < 35000 ||
    constraintsMach == null ||
    Math.abs(mach[6] - constraintsMach) > 0.01 ||
    afterburner[6] !== 0 ||
    distance[6] < 150
  ) {
    feedback.push(STRINGS.missionLegs[8]);
    missionFailed = true;
  }

  if (altitude[7] < 35000 || mach[7] !== 0.9 || afterburner[7] !== 0) {
    feedback.push(STRINGS.missionLegs[9]);
    missionFailed = true;
  }

  if (altitude[8] !== 10000 || mach[8] !== 0.4 || afterburner[8] !== 0 || time[8] !== 20) {
    feedback.push(STRINGS.missionLegs[10]);
    missionFailed = true;
  }

  if (missionFailed) {
    feedback.push(STRINGS.missionSummary);
  }

  return { delta: 0, feedback };
}
