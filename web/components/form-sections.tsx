import {
  difficulties,
  genders,
  metrics,
  prices,
  type Stage,
} from "../lib/evaluations";

export function GameSection({ stage }: { stage: Stage }) {
  return (
    <section className="form-section">
      <div className="section-number">01</div>
      <div className="section-body">
        <h2>Qual jogo você avaliou?</h2>
        <p>Use o nome informado pela equipe responsável.</p>
        <label htmlFor="game">
          Nome do jogo <span className="required">obrigatório</span>
        </label>
        <input
          id="game"
          name="game"
          autoComplete="off"
          required
          maxLength={120}
          placeholder="Ex.: Órbita"
        />
        {stage === "beta" && (
          <div className="two-fields">
            <div>
              <label htmlFor="age">Idade em anos</label>
              <input
                id="age"
                name="age"
                type="number"
                inputMode="numeric"
                min={0}
                max={120}
                step={1}
                required
              />
            </div>
            <div>
              <label htmlFor="gender">Gênero</label>
              <select id="gender" name="gender" required defaultValue="">
                <option value="" disabled>
                  Selecione
                </option>
                {genders.map((g) => (
                  <option key={g}>{g}</option>
                ))}
              </select>
            </div>
          </div>
        )}
      </div>
    </section>
  );
}

export function RatingsSection({ stage }: { stage: Stage }) {
  return (
    <section className="form-section">
      <div className="section-number">02</div>
      <div className="section-body">
        <h2>
          {stage === "alpha"
            ? "Como você percebe a ideia?"
            : "Como foi a experiência?"}
        </h2>
        <p id="scale-help">
          Escolha uma nota por critério. 0 é a menor avaliação; 10 é a maior.
        </p>
        {metrics[stage].map((metric, index) => (
          <fieldset
            className="rating"
            key={metric}
            aria-describedby="scale-help"
          >
            <legend>{metric}</legend>
            <div className="rating-options">
              {Array.from({ length: 11 }, (_, value) => (
                <label className="score" key={value}>
                  <input
                    required
                    type="radio"
                    name={`score-${index}`}
                    value={value}
                  />
                  <span>{value}</span>
                </label>
              ))}
            </div>
            <div className="scale-captions">
              <span>Menor avaliação</span>
              <span>Maior avaliação</span>
            </div>
          </fieldset>
        ))}
      </div>
    </section>
  );
}

export function ChoicesSection({ stage }: { stage: Stage }) {
  return (
    <section className="form-section">
      <div className="section-number">03</div>
      <div className="section-body">
        <h2>
          {stage === "alpha"
            ? "Você continuaria essa jornada?"
            : "O desafio ficou na medida?"}
        </h2>
        {stage === "alpha" ? (
          <>
            <fieldset className="choice">
              <legend>Você jogaria este jogo?</legend>
              <div className="choice-row">
                {["Sim", "Não"].map((answer) => (
                  <label key={answer}>
                    <input required type="radio" name="play" value={answer} />
                    {answer}
                  </label>
                ))}
              </div>
            </fieldset>
            <label htmlFor="pay">Quanto pagaria por esta ideia?</label>
            <select id="pay" name="pay" required defaultValue="">
              <option value="" disabled>
                Selecione uma faixa de preço
              </option>
              {prices.map((price) => (
                <option key={price}>{price}</option>
              ))}
            </select>
          </>
        ) : (
          <>
            <label htmlFor="difficulty">Qual foi a dificuldade do jogo?</label>
            <select id="difficulty" name="difficulty" required defaultValue="">
              <option value="" disabled>
                Selecione
              </option>
              {difficulties.map((difficulty, index) => (
                <option value={index + 1} key={difficulty}>
                  {difficulty}
                </option>
              ))}
            </select>
          </>
        )}
      </div>
    </section>
  );
}

export function CommentsSection({ stage }: { stage: Stage }) {
  return (
    <section className="form-section">
      <div className="section-number">04</div>
      <div className="section-body">
        <h2>O que a equipe precisa ouvir?</h2>
        <p>
          Seja específico e respeitoso. Os comentários são opcionais e serão
          incluídos no relatório.
        </p>
        {(stage === "alpha"
          ? ["Críticas e sugestões"]
          : ["Ponto forte", "Ponto fraco"]
        ).map((label, index) => (
          <div className="comment" key={label}>
            <label htmlFor={`comment-${index}`}>
              {label}
              <span className="optional">opcional</span>
            </label>
            <textarea
              id={`comment-${index}`}
              name={`comment-${index}`}
              maxLength={2000}
              rows={4}
              placeholder={
                index === 0
                  ? "O que funcionou? O que você mudaria?"
                  : "O que poderia melhorar na próxima versão?"
              }
            />
          </div>
        ))}
      </div>
    </section>
  );
}
