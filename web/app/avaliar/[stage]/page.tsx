import { notFound } from "next/navigation";
import Link from "next/link";
import { EvaluationForm } from "../../../components/evaluation-form";
export function generateStaticParams() {
  return [{ stage: "alpha" }, { stage: "beta" }];
}
export async function generateMetadata({
  params,
}: {
  params: Promise<{ stage: string }>;
}) {
  const { stage } = await params;
  return { title: stage === "alpha" ? "Avaliação alfa" : "Avaliação beta" };
}
export default async function EvaluationPage({
  params,
}: {
  params: Promise<{ stage: string }>;
}) {
  const { stage } = await params;
  if (stage !== "alpha" && stage !== "beta") notFound();
  return (
    <div className="form-wrap">
      <Link className="back-link" href="/">
        ← Todas as avaliações
      </Link>
      <header className="form-header">
        <span className="eyebrow">
          {stage === "alpha" ? "01 / CONCEITO" : "02 / EXPERIÊNCIA"}
        </span>
        <h1>
          Avaliação {stage === "alpha" ? "alfa" : "beta"}
          <span className="title-dot">.</span>
        </h1>
        <p>
          {stage === "alpha"
            ? "Antes de tudo, uma boa ideia. Conte o que você achou."
            : "Depois de jogar, é hora de compartilhar sua experiência."}
        </p>
        <div className="form-meta">
          <span>{stage === "alpha" ? "3" : "5"} critérios de 0 a 10</span>
          <span>{stage === "alpha" ? "Cerca de 3" : "Cerca de 5"} minutos</span>
          <span>Comentários opcionais</span>
        </div>
      </header>
      <EvaluationForm stage={stage} />
    </div>
  );
}
