import { STRINGS } from "./messages.js";
import { runMissionChecks } from "./rules/mission.js";
import { runAeroChecks } from "./rules/aero.js";
import { runThrustAndTakeoff } from "./rules/thrust.js";
import { runConstraintChecks } from "./rules/constraints.js";
import { runAttachmentChecks } from "./rules/attachments.js";
import { runStabilityChecks } from "./rules/stability.js";
import { runFuelVolumeChecks } from "./rules/fuel.js";
import { runRecurringCostChecks } from "./rules/cost.js";
import { runLandingGearChecks } from "./rules/landingGear.js";

export function gradeWorkbook(workbook, rules) {
  const feedback = [];
  let score = 10;

  if (workbook.fileName) {
    feedback.push(workbook.fileName);
  }

  // Aero tab deductions
  const aeroResult = runAeroChecks(workbook, rules);
  score += aeroResult.delta;
  feedback.push(...aeroResult.feedback);

  // Mission advisory checks (no point change)
  const missionResult = runMissionChecks(workbook, rules);
  feedback.push(...missionResult.feedback);

  // Thrust / takeoff roll
  const thrustResult = runThrustAndTakeoff(workbook, rules);
  score += thrustResult.delta;
  feedback.push(...thrustResult.feedback);

  // Constraint table
  const constraintResult = runConstraintChecks(workbook, rules);
  score += constraintResult.delta;
  feedback.push(...constraintResult.feedback);

  // Control surface attachment
  const attachmentResult = runAttachmentChecks(workbook, rules);
  score += attachmentResult.delta;
  feedback.push(...attachmentResult.feedback);

  // Stability
  const stabilityResult = runStabilityChecks(workbook, rules);
  score += stabilityResult.delta;
  feedback.push(...stabilityResult.feedback);

  // Fuel & volume
  const fuelResult = runFuelVolumeChecks(workbook, rules);
  score += fuelResult.delta;
  feedback.push(...fuelResult.feedback);

  // Recurring cost
  const costResult = runRecurringCostChecks(workbook, rules);
  score += costResult.delta;
  feedback.push(...costResult.feedback);

  // Landing gear
  const gearResult = runLandingGearChecks(workbook, rules);
  score += gearResult.delta;
  feedback.push(...gearResult.feedback);

  const clampedScore = Math.max(0, score);

  const scoreLine = STRINGS.summary.score.replace("%d", clampedScore);
  const cutoutLine = STRINGS.summary.cutout;
  feedback.push(scoreLine);
  feedback.push(cutoutLine);

  return {
    score: clampedScore,
    maxScore: 10,
    scoreLine,
    cutoutLine,
    feedbackLog: feedback.join("\n"),
  };
}
