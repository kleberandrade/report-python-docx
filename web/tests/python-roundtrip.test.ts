import { expect, it } from "vitest";
import { execFileSync } from "node:child_process";
import { mkdtempSync, readdirSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { toCsv } from "../lib/export";
import { parseResponse, type Stage } from "../lib/evaluations";
const python = process.env.CLASS_REPORT_PYTHON;
for (const stage of ["alpha", "beta"] as Stage[]) {
  it.skipIf(!python)(
    `converts the real web ${stage} export into an actual DOCX`,
    () => {
      const directory = mkdtempSync(join(tmpdir(), "classreport-web-"));
      try {
        const common = {
          id: "e550419b-d01c-432e-91c8-8a605d44a998",
          game: "Órbita",
          stage,
        };
        const item = parseResponse(
          stage === "alpha"
            ? {
                ...common,
                scores: [8, 6, 10],
                play: "Sim",
                pay: "Menos R$ 10,00",
                comments: ["Ação, música\nótima"],
              }
            : {
                ...common,
                scores: [8, 9, 7, 6, 10],
                age: 22,
                gender: "Outro",
                difficulty: 3,
                comments: ["Ação, música\nótima", ""],
              },
        );
        const input = join(directory, `${stage}.csv`),
          output = join(directory, "reports");
        writeFileSync(input, toCsv(stage, [item]));
        const result = execFileSync(
          python!,
          ["-m", "class_report", input, "--stage", stage, "-o", output],
          { cwd: resolve(".."), encoding: "utf8" },
        );
        expect(result).toContain("1 report(s)");
        const files = readdirSync(output);
        expect(files).toHaveLength(1);
        expect(files[0]).toMatch(/\.docx$/);
        execFileSync(python!, [
          "-c",
          "import sys; from docx import Document; d=Document(sys.argv[1]); text='\\n'.join(p.text for p in d.paragraphs); assert 'Órbita' in text and 'Ação, música' in text and 'ótima' in text; assert d.tables[0].cell(1,1).text == '8.00'",
          join(output, files[0]),
        ]);
      } finally {
        rmSync(directory, { recursive: true, force: true });
      }
    },
  );
}
