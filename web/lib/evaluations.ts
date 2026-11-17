import { z } from "zod";

export const metrics = {
  alpha: ["Originalidade", "Diversão", "Mercado"],
  beta: ["Gráficos", "Música e SFX", "HUD", "Jogabilidade", "História"],
} as const;
export const prices = [
  "Menos R$ 10,00",
  "Entre R$ 10,00 e R$ 20,00",
  "Entre R$ 20,00 e R$ 30,00",
  "Entre R$ 30,00 e R$ 50,00",
  "Entre R$ 50,00 e R$ 75,00",
  "Entre R$ 75,00 e R$ 100,00",
  "Mais R$ 100,00",
] as const;
export const difficulties = [
  "Muito fácil",
  "Fácil",
  "Normal",
  "Difícil",
  "Muito difícil",
];
export const genders = [
  "Feminino",
  "Masculino",
  "Outro",
  "Prefiro não informar",
] as const;
export type Stage = "alpha" | "beta";
const safeText = z
  .string()
  .trim()
  .refine((s) => !/[\x00-\x08\x0b\x0c\x0e-\x1f]/.test(s));
const common = {
  id: z.uuid(),
  game: safeText.pipe(z.string().min(1).max(120)),
};
const score = z.number().min(0).max(10);
const comment = safeText.pipe(z.string().max(2000));
const responseSchema = z.discriminatedUnion("stage", [
  z.object({
    ...common,
    stage: z.literal("alpha"),
    scores: z.tuple([score, score, score]),
    play: z.enum(["Sim", "Não"]),
    pay: z.enum(prices),
    comments: z.tuple([comment]),
  }),
  z.object({
    ...common,
    stage: z.literal("beta"),
    scores: z.tuple([score, score, score, score, score]),
    difficulty: z.number().int().min(1).max(5),
    age: z.number().int().min(0).max(120),
    gender: z.enum(genders),
    comments: z.tuple([comment, comment]),
  }),
]);
export type Evaluation = z.infer<typeof responseSchema>;
export function parseResponse(input: unknown): Evaluation {
  return responseSchema.parse(input);
}

export function parseStage(input: unknown): Stage {
  return z.enum(["alpha", "beta"]).parse(input);
}
