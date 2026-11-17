import { expect, it } from "vitest";
import { responsePayload } from "../lib/form";
import { parseResponse } from "../lib/evaluations";
const id = "e550419b-d01c-432e-91c8-8a605d44a998";
it("serializes the actual alpha form fields into a validated evaluation", () => {
  const data = new FormData();
  for (const [name, value] of Object.entries({
    game: "Órbita",
    "score-0": "8",
    "score-1": "6",
    "score-2": "10",
    play: "Sim",
    pay: "Menos R$ 10,00",
    "comment-0": "Ação",
  }))
    data.set(name, value);
  expect(parseResponse(responsePayload(data, "alpha", id)).scores).toEqual([
    8, 6, 10,
  ]);
  expect(parseResponse(responsePayload(data, "alpha", id)).comments).toEqual([
    "Ação",
  ]);
});
it("keeps empty beta comments and rejects omitted scores instead of inventing zero", () => {
  const data = new FormData();
  data.set("game", "Órbita");
  data.set("age", "22");
  data.set("gender", "Outro");
  data.set("difficulty", "3");
  for (let i = 0; i < 5; i++) data.set(`score-${i}`, "7");
  expect(parseResponse(responsePayload(data, "beta", id)).comments).toEqual([
    "",
    "",
  ]);
  data.delete("score-0");
  expect(() => parseResponse(responsePayload(data, "beta", id))).toThrow();
});
it("rejects blank beta numbers and preserves a genuine zero score", () => {
  const data = new FormData();
  for (const [name, value] of Object.entries({
    game: "Órbita",
    age: "22",
    gender: "Outro",
    difficulty: "3",
    "score-0": "0",
    "score-1": "7",
    "score-2": "7",
    "score-3": "7",
    "score-4": "7",
  }))
    data.set(name, value);
  expect(parseResponse(responsePayload(data, "beta", id)).scores[0]).toBe(0);
  for (const name of ["age", "difficulty", "score-0"]) {
    const old = data.get(name) as string;
    data.set(name, "");
    expect(() => parseResponse(responsePayload(data, "beta", id))).toThrow();
    data.set(name, old);
  }
});
