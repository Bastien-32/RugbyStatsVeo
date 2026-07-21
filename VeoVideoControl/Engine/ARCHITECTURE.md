# Rugby Sync Engine

## Source de vérité

La source de vérité du système est toujours la vidéo.

Excel ne fait jamais d'hypothèse sur la position de la vidéo.

Le moteur lit toujours l'état réel de la vidéo puis synchronise Excel.

## Flux Veo → Excel

Lecture
Pause
Déplacement temporel
Retour au début
Position actuelle

=> Synchronisation Excel

## Flux Excel → Veo

Play / Pause
-5 secondes
+5 secondes
Reset

=> Commande vidéo

## Principe

La vidéo pilote toujours le chrono.

Excel ne fait que refléter son état réel.