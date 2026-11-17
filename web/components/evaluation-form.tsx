"use client";
import Link from "next/link";
import type { Stage } from "../lib/evaluations";
import {
  GameSection,
  RatingsSection,
  ChoicesSection,
  CommentsSection,
} from "./form-sections";
import { useSubmission } from "./use-submission";

export function EvaluationForm({ stage }: { stage: Stage }) {
  const { status, error, submit, reset } = useSubmission(stage);
  if (status === "saved")
    return (
      <section className="success" role="status">
        <span className="success-icon" aria-hidden="true">
          ✓
        </span>
        <div className="eyebrow">AVALIAÇÃO REGISTRADA</div>
        <h2>Seu olhar faz diferença.</h2>
        <p>
          As respostas foram salvas. Obrigado por ajudar a construir a próxima
          versão.
        </p>
        <Link className="button" href="/">
          Voltar ao início
        </Link>
        <button className="plain-button" onClick={reset}>
          Avaliar outro jogo
        </button>
      </section>
    );
  return (
    <form onSubmit={submit} className="evaluation-form">
      <fieldset disabled={status === "saving"}>
        <GameSection stage={stage} />
        <RatingsSection stage={stage} />
        <ChoicesSection stage={stage} />
        <CommentsSection stage={stage} />
      </fieldset>
      <div className="submit-area">
        {error && (
          <p className="error-message" role="alert">
            {error}
          </p>
        )}
        <div className="submit-row">
          <p>
            Sem nome, sem email.
            <br />
            Não inclua dados pessoais nos comentários.
          </p>
          <button
            className="button"
            disabled={status === "saving"}
            type="submit"
          >
            {status === "saving" ? "Salvando avaliação…" : "Enviar avaliação →"}
          </button>
        </div>
        <p className="status-text" role="status">
          {status === "saving"
            ? "Aguarde a confirmação antes de fechar esta página."
            : ""}
        </p>
      </div>
    </form>
  );
}
