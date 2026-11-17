import type { Evaluation, Stage } from "./evaluations";
import { metrics } from "./evaluations";

function cell(value: string | number) {
  let text = String(value);
  if (/^[\s]*[=+\-@\t\r]/.test(text)) text = "'" + text;
  return /[",\r\n]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
}
export function toCsv(stage: Stage, evaluations: Evaluation[]): string {
  const headers =
    stage === "alpha"
      ? [
          "Jogo avaliado",
          ...metrics.alpha,
          "Você jogaria este jogo?",
          "Críticas e Sugestões",
          "Quanto pagaria por esta ideia?",
        ]
      : [
          "Jogo avaliado",
          ...metrics.beta,
          "Dificuldade",
          "Ponto forte",
          "Ponto fraco",
          "Idade",
          "Sexo",
        ];
  const rows = evaluations
    .filter((item) => item.stage === stage)
    .map((item) =>
      item.stage === "alpha"
        ? [item.game, ...item.scores, item.play, item.comments[0], item.pay]
        : [
            item.game,
            ...item.scores,
            item.difficulty,
            ...item.comments,
            item.age,
            item.gender,
          ],
    );
  return (
    "\uFEFF" +
    [headers, ...rows].map((row) => row.map(cell).join(",")).join("\r\n") +
    "\r\n"
  );
}
