import { metrics, type Stage } from "./evaluations";
function field(data: FormData, name: string): string {
  return data.has(name) ? String(data.get(name)) : "";
}
function numericField(data: FormData, name: string): number {
  const value = field(data, name);
  return value.trim().length === 0 ? NaN : Number(value);
}
export function responsePayload(
  data: FormData,
  stage: Stage,
  id: string,
): object {
  const common = {
    id,
    stage,
    game: field(data, "game"),
    scores: metrics[stage].map((_, i) => numericField(data, `score-${i}`)),
  };
  return stage === "alpha"
    ? {
        ...common,
        play: field(data, "play"),
        pay: field(data, "pay"),
        comments: [field(data, "comment-0")],
      }
    : {
        ...common,
        difficulty: numericField(data, "difficulty"),
        age: numericField(data, "age"),
        gender: field(data, "gender"),
        comments: [field(data, "comment-0"), field(data, "comment-1")],
      };
}
