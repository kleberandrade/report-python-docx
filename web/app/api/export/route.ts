import { authorized, readResponses } from "../../../lib/database";
import { boundedJson, sameOrigin } from "../../../lib/http";
import { parseStage, type Stage } from "../../../lib/evaluations";
import { toCsv } from "../../../lib/export";
export async function POST(request: Request) {
  if (!sameOrigin(request))
    return Response.json({ error: "Origem não permitida." }, { status: 403 });
  if (
    !authorized(
      request.headers.get("authorization")?.replace(/^Bearer /, "") ?? null,
    )
  ) {
    return Response.json(
      { error: "Chave de acesso inválida ou não configurada." },
      { status: 401 },
    );
  }
  let stage: Stage;
  try {
    const input = (await boundedJson(request)) as { stage?: unknown };
    stage = parseStage(input?.stage);
  } catch {
    return Response.json({ error: "Escolha alfa ou beta." }, { status: 400 });
  }
  try {
    const items = await readResponses(stage);
    if (!items.length)
      return Response.json(
        { error: "Ainda não há respostas nesta etapa." },
        { status: 404 },
      );
    return new Response(toCsv(stage, items), {
      headers: {
        "Content-Type": "text/csv; charset=utf-8",
        "Cache-Control": "no-store",
        "Content-Disposition": `attachment; filename="${stage}-evaluation.csv"`,
        "X-Response-Count": String(items.length),
      },
    });
  } catch {
    return Response.json(
      {
        error:
          "Exportação indisponível. Verifique a configuração do banco ou o limite de 10 mil respostas.",
      },
      { status: 503 },
    );
  }
}
