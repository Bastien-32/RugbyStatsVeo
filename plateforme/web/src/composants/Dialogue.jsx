import { useEffect, useRef, useState } from "react";

/**
 * La coquille commune à toutes les popups de saisie.
 *
 * Valider ne ferme que si la popup rend vrai : il manque
 * parfois un choix, et fermer sur rien ferait disparaître la
 * saisie sans rien dire.
 */
export function Dialogue({ titre, sous, children, surValidation, surFermeture }) {
  const ref = useRef(null);

  useEffect(() => {
    const d = ref.current;
    if (d && !d.open) d.showModal();
  }, []);

  function valider() {
    if (surValidation() === false) return;
    ref.current?.close();
    surFermeture();
  }

  return (
    <dialog
      ref={ref}
      onCancel={(e) => {
        e.preventDefault();
        surFermeture();
      }}
    >
      <div className="dlg-tete">
        <h3>{titre}</h3>
        {sous && <p>{sous}</p>}
      </div>
      <div className="dlg-corps">{children}</div>
      <div className="dlg-pied">
        <button type="button" className="btn-vide" onClick={surFermeture}>
          Annuler
        </button>
        <button type="button" className="btn-plein" onClick={valider}>
          Valider
        </button>
      </div>
    </dialog>
  );
}

/**
 * Un groupe de boutons dont un seul s'allume.
 *
 * Rien n'est présélectionné : une valeur ne part au journal que
 * si elle a été choisie. Une présélection silencieuse écrirait
 * dans les données une information que personne n'a vue.
 */
export function Choix({ libelle, options, valeur, surChoix, colonnes }) {
  return (
    <div className="champ">
      <label>{libelle}</label>
      <div className={"choix" + (colonnes ? " colonnes" : "")}>
        {options.map((o) => {
          const code = typeof o === "string" ? o : o.code;
          const texte = typeof o === "string" ? o : o.libelle;
          return (
            <button
              key={code}
              type="button"
              aria-pressed={valeur === code}
              onClick={() => surChoix(valeur === code ? null : code)}
            >
              {texte}
            </button>
          );
        })}
      </div>
    </div>
  );
}

/** Un champ de choix dont l'état vit dans le composant parent. */
export function useChoix(initial = null) {
  const [valeur, setValeur] = useState(initial);
  return [valeur, setValeur];
}
