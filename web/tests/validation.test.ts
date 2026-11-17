import { describe, expect, it } from "vitest";
import { parseResponse } from "../lib/evaluations";
import { toCsv } from "../lib/export";

export const alpha = {
  id: "e550419b-d01c-432e-91c8-8a605d44a998",
  stage: "alpha",
  game: "Órbita",
  scores: [8, 6, 10],
  play: "Sim",
  pay: "Menos R$ 10,00",
  comments: ["Ótima ação"],
};

describe("evaluation boundary", () => {
  it("preserves Portuguese answers", () =>
    expect(parseResponse(alpha).game).toBe("Órbita"));
  it.each([-1, 11, "8", null, true, Infinity])(
    "rejects bad score %s",
    (score) => {
      expect(() =>
        parseResponse({ ...alpha, scores: [score, 6, 10] }),
      ).toThrow();
    },
  );
  it("rejects missing, extra, and long responses", () => {
    for (const change of [
      { game: "" },
      { scores: [8, 6] },
      { comments: ["a".repeat(2001)] },
      { play: "Talvez" },
      { pay: "" },
      { id: "not-a-uuid" },
    ]) {
      expect(() => parseResponse({ ...alpha, ...change })).toThrow();
    }
  });
  it("validates beta age and difficulty", () => {
    const beta = {
      ...alpha,
      stage: "beta",
      scores: [8, 7, 9, 6, 10],
      difficulty: 3,
      age: 22,
      gender: "Prefiro não informar",
      comments: ["", ""],
    };
    expect(parseResponse(beta).stage).toBe("beta");
    expect(() => parseResponse({ ...beta, difficulty: 2.5 })).toThrow();
    expect(() => parseResponse({ ...beta, age: 121 })).toThrow();
  });
});

describe("CSV compatibility and spreadsheet safety", () => {
  it("exports exact Portuguese headers, numeric scores and quoted multiline comments", () => {
    const csv = toCsv("alpha", [
      parseResponse({ ...alpha, comments: ["Ação, música\nótima"] }),
    ]);
    expect(csv).toContain("Jogo avaliado,Originalidade,Diversão,Mercado");
    expect(csv).toContain('Órbita,8,6,10,Sim,"Ação, música\nótima",');
  });
  it("neutralizes formulas in text cells and retains leading apostrophe explicitly", () => {
    const csv = toCsv("alpha", [parseResponse({ ...alpha, game: "=1+1" })]);
    expect(csv).toContain("'=1+1");
    expect(toCsv("alpha", [])).toContain("Críticas e Sugestões");
  });
});
