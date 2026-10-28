import { useEffect, useState } from "react";
import { api } from "../api.js";
import { court, mmss } from "../etat/regles.js";

/**
 * Le journal d'actions.
 *
 * Chaque ligne garde le temps vidéo exact et y renvoie d'un
 * clic : c'est ce que le classeur ne pouvait pas faire.
 */
export function Journal({ actions, etat, surSaut, surSuppression }) {
  if (actions.length === 0) {
    return (
      <p className="vide">
        Le journal se remplit à chaque clic sur la palette.
        <br />
        Chaque ligne garde le temps vidéo exact et y renvoie.
      </p>
    );
  }

  return (
    <table>
      <thead>
        <tr>
          <th>Temps</th>
          <th>MT</th>
          <th>Joueur</th>
          <th>Action</th>
          <th>Détail</th>
          <th>Poss.</th>
          <th />
        </tr>
      </thead>
      <tbody>
        {actions.map((a) => (
          <tr key={a.id}>
            <td className="t-temps" onClick={() => surSaut(a.temps)}>
              {mmss(a.temps)}
            </td>
            <td>{a.mt}</td>
            <td>
              {a.numero != null
                ? etat.etiquette(a.numero)
                : a.groupe ?? "—"}
            </td>
            <td>{a.libelle}</td>
            <td>{detailLisible(a)}</td>
            <td>
              <span className={"puce " + (a.possession === "Nous" ? "nous" : "adv")}>
                {a.possession}
              </span>
            </td>
            <td>
              <button
                type="button"
                className="sup"
                aria-label="Supprimer cette ligne"
                onClick={() => surSuppression(a.id)}
              >
                ×
              </button>
            </td>
          </tr>
        ))}
      </tbody>
    </table>
  );
}

function detailLisible(a) {
  const morceaux = [];
  if (a.motif_libelle) morceaux.push(a.motif_libelle);
  const d = a.detail;
  if (d) {
    morceaux.push(
      [d.zone_saut, d.alignement, d.ballon, d.zone_longueur,
       d.zone_largeur, d.utilisation].filter(Boolean).join(" / ")
    );
  }
  return morceaux.filter(Boolean).join(" · ") || "—";
}

/**
 * Le temps de jeu, calculé par l'API depuis les remplacements.
 *
 * Rien n'est stocké : il se recalcule, et reste donc juste après
 * la moindre correction.
 */
export function TempsDeJeu({ matchId, rafraichir }) {
  const stats = useStatistiques(matchId, rafraichir);

  if (!stats) return <p className="vide">Chargement…</p>;
  if (stats.temps_de_jeu.length === 0) {
    return (
      <p className="vide">
        Le temps de jeu apparaîtra dès que la feuille de match sera posée
        et le journal entamé.
      </p>
    );
  }

  const max = Math.max(...stats.temps_de_jeu.map((t) => t.secondes), 1);

  return (
    <table>
      <thead>
        <tr>
          <th>N°</th>
          <th>Joueur</th>
          <th>Ligne</th>
          <th className="t-num">Temps de jeu</th>
          <th style={{ width: "30%" }} />
        </tr>
      </thead>
      <tbody>
        {stats.temps_de_jeu.map((t) => (
          <tr key={t.numero}>
            <td>{t.numero}</td>
            <td>{t.nom_complet}</td>
            <td>{t.ligne}</td>
            <td className="t-num">{court(t.secondes)}</td>
            <td>
              <span
                className="barre-jauge"
                style={{ width: `${Math.round((t.secondes * 100) / max)}%` }}
              />
            </td>
          </tr>
        ))}
      </tbody>
    </table>
  );
}

/** Le tableau des durées de possession, par tranches. */
export function Possession({ matchId, rafraichir }) {
  const stats = useStatistiques(matchId, rafraichir);
  if (!stats) return <p className="vide">Chargement…</p>;

  const p = stats.possession;
  const total = p.cumul_nous + p.cumul_adv;
  const part = (v) => (total > 0 ? `${Math.round((v * 100) / total)} %` : "—");
  const compte = (camp) =>
    (p.sequences ?? []).filter((s) => s.equipe === camp).length;

  return (
    <table>
      <thead>
        <tr>
          <th>Durée des séquences</th>
          <th className="t-num">Nous</th>
          <th className="t-num">Adversaire</th>
        </tr>
      </thead>
      <tbody>
        {p.tranches.map((nom, i) => (
          <tr key={nom}>
            <td>{nom}</td>
            <td className="t-num">{p.nous[i]}</td>
            <td className="t-num">{p.adv[i]}</td>
          </tr>
        ))}
        <tr>
          <td>
            <strong>Temps cumulé</strong>
          </td>
          <td className="t-num">
            <strong>
              {court(p.cumul_nous)} · {part(p.cumul_nous)}
            </strong>
          </td>
          <td className="t-num">
            <strong>
              {court(p.cumul_adv)} · {part(p.cumul_adv)}
            </strong>
          </td>
        </tr>
        <tr>
          <td>Séquences jouées</td>
          <td className="t-num">{compte("Nous")}</td>
          <td className="t-num">{compte("Adv")}</td>
        </tr>
      </tbody>
    </table>
  );
}

/** Ce que chaque joueur a produit, compté depuis le journal. */
export function ParJoueur({ actions, etat, referentiels }) {
  const familleDe = {};
  for (const t of referentiels?.types ?? []) familleDe[t.code] = t.famille;

  const parJoueur = {};
  for (const a of actions) {
    if (a.numero == null) continue;
    const nom = etat.etiquette(a.numero);
    const j = (parJoueur[nom] ||= { pl: 0, att: 0, poss: 0, faute: 0, autre: 0 });
    switch (familleDe[a.type_code]) {
      case "PLAQUAGE": j.pl++; break;
      case "ATTITUDE": j.att++; break;
      case "POSSESSION": j.poss++; break;
      case "PENALITE":
      case "COUP_FRANC": j.faute++; break;
      default: j.autre++;
    }
  }

  const noms = Object.keys(parJoueur).sort(
    (a, b) => parJoueur[b].pl + parJoueur[b].att - (parJoueur[a].pl + parJoueur[a].att)
  );

  if (noms.length === 0) {
    return <p className="vide">Les actions attribuées à un joueur apparaîtront ici.</p>;
  }

  return (
    <table>
      <thead>
        <tr>
          <th>Joueur</th>
          <th className="t-num">Plaquages</th>
          <th className="t-num">Au contact</th>
          <th className="t-num">Ballons</th>
          <th className="t-num">Fautes</th>
          <th className="t-num">Autres</th>
        </tr>
      </thead>
      <tbody>
        {noms.map((n) => (
          <tr key={n}>
            <td>{n}</td>
            <td className="t-num">{parJoueur[n].pl}</td>
            <td className="t-num">{parJoueur[n].att}</td>
            <td className="t-num">{parJoueur[n].poss}</td>
            <td className="t-num">{parJoueur[n].faute}</td>
            <td className="t-num">{parJoueur[n].autre}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );
}

/** Le détail des touches et mêlées saisies. */
export function TouchesMelees({ actions, etat, surSaut }) {
  const lignes = actions
    .filter((a) => a.detail)
    .sort((a, b) => a.temps - b.temps);

  if (lignes.length === 0) {
    return (
      <p className="vide">
        Les touches et mêlées saisies apparaîtront ici avec leur détail.
      </p>
    );
  }

  return (
    <table>
      <thead>
        <tr>
          <th>Temps</th>
          <th>MT</th>
          <th>Sauteur</th>
          <th>Issue</th>
          <th>Zone</th>
          <th>Utilisation</th>
        </tr>
      </thead>
      <tbody>
        {lignes.map((a) => (
          <tr key={a.id}>
            <td className="t-temps" onClick={() => surSaut(a.temps)}>
              {mmss(a.temps)}
            </td>
            <td>{a.mt}</td>
            <td>
              {a.detail.sauteur_numero
                ? etat.etiquette(a.detail.sauteur_numero)
                : "—"}
            </td>
            <td>{a.libelle}</td>
            <td>
              {[a.detail.zone_saut, a.detail.zone_longueur, a.detail.zone_largeur]
                .filter(Boolean)
                .join(" / ") || "—"}
            </td>
            <td>{a.detail.utilisation || "—"}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );
}

// Les statistiques viennent du serveur : les mêmes règles
// servent à l'interface et aux exports.
function useStatistiques(matchId, rafraichir) {
  const [stats, setStats] = useState(null);

  useEffect(() => {
    let vivant = true;
    api
      .statistiques(matchId)
      .then((s) => vivant && setStats(s))
      .catch(() => vivant && setStats(null));
    return () => {
      vivant = false;
    };
  }, [matchId, rafraichir]);

  return stats;
}
