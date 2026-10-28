import { useEffect, useRef, useState } from "react";

/**
 * Le lecteur et ses commandes.
 *
 * La vidéo reste sur le poste : elle n'est jamais envoyée au
 * serveur. Seul son horodatage part au journal, ce qui suffit à
 * y revenir.
 */
export function Video({ surPosition, enregistrerLecteur }) {
  const video = useRef(null);
  const [charge, setCharge] = useState(false);
  const [enLecture, setEnLecture] = useState(false);
  const arriere = useRef(null);

  useEffect(() => {
    enregistrerLecteur(video.current);
  }, [enregistrerLecteur]);

  // Le défilement arrière s'arrête si le composant disparaît.
  useEffect(() => () => clearInterval(arriere.current), []);

  function charger(fichier) {
    if (!fichier) return;
    video.current.src = URL.createObjectURL(fichier);
    setCharge(true);
    video.current.play().catch(() => {});
  }

  function reculer() {
    if (arriere.current) {
      clearInterval(arriere.current);
      arriere.current = null;
      return;
    }
    video.current.pause();
    // Aucun navigateur ne lit une vidéo à l'envers : on recule
    // par petits pas, ce qui donne un défilement continu.
    arriere.current = setInterval(() => {
      const v = video.current;
      if (!v) return;
      v.currentTime = Math.max(0, v.currentTime - 0.2);
      if (v.currentTime <= 0) {
        clearInterval(arriere.current);
        arriere.current = null;
      }
    }, 100);
  }

  return (
    <div className="carte bloc-video">
      <div
        className="zone-video"
        onDragOver={(e) => e.preventDefault()}
        onDrop={(e) => {
          e.preventDefault();
          const f = e.dataTransfer?.files?.[0];
          if (f?.type.startsWith("video/")) charger(f);
        }}
      >
        <video
          ref={video}
          playsInline
          onTimeUpdate={(e) => surPosition(e.target.currentTime)}
          onPlay={() => setEnLecture(true)}
          onPause={() => setEnLecture(false)}
        />
        {!charge && (
          <div className="accueil">
            <strong>Déposez la vidéo du match</strong>
            <p>
              Un fichier local, ou l'export Veo. Chaque action est horodatée
              sur la vidéo et cliquable pour y revenir.
            </p>
            <label className="bouton-fichier">
              Choisir un fichier
              <input
                id="fichier-video"
                type="file"
                accept="video/*"
                onChange={(e) => charger(e.target.files[0])}
              />
            </label>
          </div>
        )}
      </div>

      <div className="commandes">
        <button
          type="button"
          className="cmd primaire"
          onClick={() => {
            const v = video.current;
            if (!v.src) return document.getElementById("fichier-video")?.click();
            v.paused ? v.play() : v.pause();
          }}
        >
          {enLecture ? "Pause" : "Lecture"}
        </button>
        <button
          type="button"
          className="cmd"
          onClick={() => {
            video.current.currentTime = Math.max(0, video.current.currentTime - 5);
          }}
        >
          −5 s
        </button>
        <button
          type="button"
          className="cmd"
          onClick={() => {
            video.current.currentTime += 5;
          }}
        >
          +5 s
        </button>
        <button type="button" className="cmd" onClick={reculer}>
          {arriere.current ? "■ Stop" : "◀ Arrière"}
        </button>
        <span className="vitesse">
          <label htmlFor="vitesse">Vitesse</label>
          <select
            id="vitesse"
            defaultValue="1"
            onChange={(e) => {
              video.current.playbackRate = Number(e.target.value);
            }}
          >
            <option value="0.5">0,5×</option>
            <option value="1">1×</option>
            <option value="1.5">1,5×</option>
            <option value="2">2×</option>
          </select>
        </span>
      </div>
    </div>
  );
}

/**
 * Une poignée qu'on attrape pour redimensionner deux zones.
 *
 * Les flèches du clavier font la même chose, pour qui n'utilise
 * pas la souris.
 */
export function Separateur({ orientation, valeur, surChangement, min, max, libelle }) {
  const vertical = orientation === "vertical";
  const ref = useRef(null);

  function demarrer(e) {
    e.preventDefault();
    const poignee = ref.current;
    poignee.setPointerCapture(e.pointerId);
    poignee.classList.add("actif");
    document.body.classList.add("redimensionne");

    const parent = poignee.parentElement.getBoundingClientRect();

    const bouger = (ev) => {
      const pct = vertical
        ? ((ev.clientX - parent.left) / parent.width) * 100
        : ((ev.clientY - parent.top) / parent.height) * 100;
      surChangement(Math.min(max, Math.max(min, pct)));
    };
    const finir = () => {
      poignee.removeEventListener("pointermove", bouger);
      poignee.classList.remove("actif");
      document.body.classList.remove("redimensionne");
    };
    poignee.addEventListener("pointermove", bouger);
    poignee.addEventListener("pointerup", finir, { once: true });
    poignee.addEventListener("pointercancel", finir, { once: true });
  }

  function clavier(e) {
    const moins = vertical ? "ArrowLeft" : "ArrowUp";
    const plus = vertical ? "ArrowRight" : "ArrowDown";
    if (e.key === moins) surChangement(Math.max(min, valeur - 2));
    else if (e.key === plus) surChangement(Math.min(max, valeur + 2));
    else return;
    e.preventDefault();
  }

  return (
    <div
      ref={ref}
      className={`sep ${vertical ? "sep-v" : "sep-h"}`}
      role="separator"
      aria-orientation={orientation}
      aria-label={libelle}
      aria-valuenow={Math.round(valeur)}
      tabIndex={0}
      onPointerDown={demarrer}
      onKeyDown={clavier}
    />
  );
}
