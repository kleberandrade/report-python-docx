import { parseResponse } from "../../../lib/evaluations";
import { saveResponse } from "../../../lib/database";
import { boundedJson, sameOrigin } from "../../../lib/http";
export async function POST(request: Request) {
  if (!sameOrigin(request))
    return Response.json({ error: "Origem não permitida." }, { status: 403 });
  let item;
  try {
    item = parseResponse(await boundedJson(request));
  } catch {
    return Response.json(
      { error: "Confira o jogo, as notas e as respostas obrigatórias." },
      { status: 400 },
    );
  }
  if (!process.env.DATABASE_URL) {
    return Response.json(
      {
        error:
          "A coleta ainda não está configurada. Nenhuma resposta foi salva.",
      },
      { status: 503 },
    );
  }
  try {
    await saveResponse(item);
    return Response.json(
      { saved: true },
      { status: 201, headers: { "Cache-Control": "no-store" } },
    );
  } catch {
    return Response.json(
      {
        error:
          "Não foi possível salvar agora. Suas respostas foram mantidas; tente novamente.",
      },
      { status: 503 },
    );
  }
}
