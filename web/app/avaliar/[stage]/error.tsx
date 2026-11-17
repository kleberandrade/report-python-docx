"use client";
export default function Error({ reset }: { reset: () => void }) {
  return (
    <div className="wrap empty" role="alert">
      <h1>Não foi possível abrir a avaliação.</h1>
      <button className="button" onClick={reset}>
        Tentar novamente
      </button>
    </div>
  );
}
