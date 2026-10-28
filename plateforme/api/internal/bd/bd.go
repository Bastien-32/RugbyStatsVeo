// Package bd regroupe les accès à PostgreSQL.
//
// Le SQL est écrit à la main, en clair : il se lit, se copie
// dans psql et s'explique. Un ORM l'aurait caché, au moment
// précis où le modèle est ce qu'on veut montrer.
package bd

import (
	"context"
	"fmt"

	"github.com/jackc/pgx/v5"
	"github.com/jackc/pgx/v5/pgxpool"

	"rugbystats/api/internal/modele"
)

type Base struct {
	pool *pgxpool.Pool
}

func Ouvrir(ctx context.Context, url string) (*Base, error) {
	pool, err := pgxpool.New(ctx, url)
	if err != nil {
		return nil, fmt.Errorf("connexion : %w", err)
	}
	if err := pool.Ping(ctx); err != nil {
		return nil, fmt.Errorf("la base ne répond pas : %w", err)
	}
	return &Base{pool: pool}, nil
}

func (b *Base) Fermer() { b.pool.Close() }

func (b *Base) Verifier(ctx context.Context) error { return b.pool.Ping(ctx) }

// ---------------------------------------------------------
// Référentiels
// ---------------------------------------------------------

func (b *Base) Referentiels(ctx context.Context) (modele.Referentiels, error) {
	var r modele.Referentiels

	rows, err := b.pool.Query(ctx,
		`SELECT code, libelle, rang FROM famille_action ORDER BY rang`)
	if err != nil {
		return r, err
	}
	r.Familles, err = pgx.CollectRows(rows,
		pgx.RowToStructByPos[modele.FamilleAction])
	if err != nil {
		return r, err
	}

	rows, err = b.pool.Query(ctx,
		`SELECT code, famille, libelle, rang,
		        ouvre_jeu, ferme_jeu, arret_auto, bascule, attend_joueur
		 FROM type_action
		 ORDER BY famille, rang`)
	if err != nil {
		return r, err
	}
	r.Types, err = pgx.CollectRows(rows,
		pgx.RowToStructByPos[modele.TypeAction])
	if err != nil {
		return r, err
	}

	rows, err = b.pool.Query(ctx,
		`SELECT code, libelle, famille, penalite, coup_franc, rang
		 FROM motif_faute
		 ORDER BY famille, rang`)
	if err != nil {
		return r, err
	}
	r.Motifs, err = pgx.CollectRows(rows,
		pgx.RowToStructByPos[modele.MotifFaute])
	return r, err
}

// TypesParCode sert aux calculs : les règles ont besoin de
// savoir ce que chaque action fait au jeu.
func (b *Base) TypesParCode(ctx context.Context) (map[string]modele.TypeAction, error) {
	r, err := b.Referentiels(ctx)
	if err != nil {
		return nil, err
	}
	m := make(map[string]modele.TypeAction, len(r.Types))
	for _, t := range r.Types {
		m[t.Code] = t
	}
	return m, nil
}

// ---------------------------------------------------------
// Joueurs
// ---------------------------------------------------------

func (b *Base) Joueurs(ctx context.Context, clubID int) ([]modele.Joueur, error) {
	rows, err := b.pool.Query(ctx,
		`SELECT id, nom, prenom, ligne, actif
		 FROM joueur
		 WHERE club_id = $1 AND actif
		 ORDER BY nom, prenom`, clubID)
	if err != nil {
		return nil, err
	}
	return pgx.CollectRows(rows, pgx.RowToStructByPos[modele.Joueur])
}

func (b *Base) CreerJoueur(ctx context.Context, clubID int,
	j modele.Joueur) (modele.Joueur, error) {

	err := b.pool.QueryRow(ctx,
		`INSERT INTO joueur (club_id, nom, prenom, ligne)
		 VALUES ($1, $2, $3, $4)
		 RETURNING id, nom, prenom, ligne, actif`,
		clubID, j.Nom, j.Prenom, j.Ligne).
		Scan(&j.ID, &j.Nom, &j.Prenom, &j.Ligne, &j.Actif)
	return j, err
}

// ---------------------------------------------------------
// Matchs
// ---------------------------------------------------------

const colonnesMatch = `id, saison, to_char(date_match,'YYYY-MM-DD'),
	adversaire, categorie, lieu::text, phase, journee, sens_mt1, video_ref`

func (b *Base) Matchs(ctx context.Context, clubID int) ([]modele.Match, error) {
	rows, err := b.pool.Query(ctx,
		`SELECT `+colonnesMatch+`
		 FROM match WHERE club_id = $1
		 ORDER BY date_match DESC, id DESC`, clubID)
	if err != nil {
		return nil, err
	}
	return pgx.CollectRows(rows, pgx.RowToStructByPos[modele.Match])
}

func (b *Base) Match(ctx context.Context, id int) (modele.Match, error) {
	rows, err := b.pool.Query(ctx,
		`SELECT `+colonnesMatch+` FROM match WHERE id = $1`, id)
	if err != nil {
		return modele.Match{}, err
	}
	return pgx.CollectOneRow(rows, pgx.RowToStructByPos[modele.Match])
}

func (b *Base) CreerMatch(ctx context.Context, clubID int,
	m modele.Match) (modele.Match, error) {

	rows, err := b.pool.Query(ctx,
		`INSERT INTO match
		   (club_id, saison, date_match, adversaire, categorie,
		    lieu, phase, journee, sens_mt1, video_ref)
		 VALUES ($1,$2,$3::date,$4,$5,$6::lieu_match,$7,$8,$9,$10)
		 RETURNING `+colonnesMatch,
		clubID, m.Saison, m.DateMatch, m.Adversaire, m.Categorie,
		m.Lieu, m.Phase, m.Journee, m.SensMT1, m.VideoRef)
	if err != nil {
		return modele.Match{}, err
	}
	return pgx.CollectOneRow(rows, pgx.RowToStructByPos[modele.Match])
}

// ---------------------------------------------------------
// Feuille de match
// ---------------------------------------------------------

func (b *Base) Feuille(ctx context.Context, matchID int) ([]modele.LigneFeuille, error) {
	rows, err := b.pool.Query(ctx,
		`SELECT f.numero, f.joueur_id,
		        trim(j.nom || ' ' || j.prenom) AS nom_complet,
		        f.ligne, f.titulaire
		 FROM feuille_match f
		 JOIN joueur j ON j.id = f.joueur_id
		 WHERE f.match_id = $1
		 ORDER BY f.numero`, matchID)
	if err != nil {
		return nil, err
	}
	return pgx.CollectRows(rows, pgx.RowToStructByPos[modele.LigneFeuille])
}

// PoserFeuille remplace la feuille entière : la composition se
// corrige d'un bloc, et une ligne orpheline ne survit pas.
func (b *Base) PoserFeuille(ctx context.Context, matchID int,
	lignes []modele.LigneFeuille) error {

	tx, err := b.pool.Begin(ctx)
	if err != nil {
		return err
	}
	defer tx.Rollback(ctx)

	if _, err := tx.Exec(ctx,
		`DELETE FROM feuille_match WHERE match_id = $1`, matchID); err != nil {
		return err
	}
	for _, l := range lignes {
		if _, err := tx.Exec(ctx,
			`INSERT INTO feuille_match
			   (match_id, numero, joueur_id, ligne, titulaire)
			 VALUES ($1,$2,$3,$4::ligne_jeu,$5)`,
			matchID, l.Numero, l.JoueurID, l.Ligne, l.Titulaire); err != nil {
			return fmt.Errorf("numéro %d : %w", l.Numero, err)
		}
	}
	return tx.Commit(ctx)
}

// ---------------------------------------------------------
// Journal d'actions
// ---------------------------------------------------------

func (b *Base) Actions(ctx context.Context, matchID int) ([]modele.Action, error) {
	rows, err := b.pool.Query(ctx,
		`SELECT a.id, a.match_id, a.temps::float8, a.mt, a.type_code,
		        t.libelle, a.possession, a.numero, a.groupe,
		        a.motif_code, coalesce(m.libelle,''), a.commentaire
		 FROM action a
		 JOIN type_action t ON t.code = a.type_code
		 LEFT JOIN motif_faute m ON m.code = a.motif_code
		 WHERE a.match_id = $1
		 ORDER BY a.temps DESC, a.id DESC`, matchID)
	if err != nil {
		return nil, err
	}
	defer rows.Close()

	var liste []modele.Action
	for rows.Next() {
		var a modele.Action
		if err := rows.Scan(&a.ID, &a.MatchID, &a.Temps, &a.MT, &a.TypeCode,
			&a.Libelle, &a.Possession, &a.Numero, &a.Groupe,
			&a.MotifCode, &a.MotifLib, &a.Commentaire); err != nil {
			return nil, err
		}
		liste = append(liste, a)
	}
	if err := rows.Err(); err != nil {
		return nil, err
	}
	return b.attacherDetails(ctx, matchID, liste)
}

func (b *Base) attacherDetails(ctx context.Context, matchID int,
	liste []modele.Action) ([]modele.Action, error) {

	rows, err := b.pool.Query(ctx,
		`SELECT d.action_id, d.issue, d.sauteur_numero,
		        coalesce(d.zone_saut,''), coalesce(d.alignement,''),
		        coalesce(d.ballon,''), coalesce(d.zone_longueur,''),
		        coalesce(d.zone_largeur,''), coalesce(d.utilisation,''),
		        d.observation
		 FROM detail_phase d
		 JOIN action a ON a.id = d.action_id
		 WHERE a.match_id = $1`, matchID)
	if err != nil {
		return liste, err
	}
	defer rows.Close()

	details := map[int64]*modele.Detail{}
	for rows.Next() {
		var id int64
		d := &modele.Detail{}
		if err := rows.Scan(&id, &d.Issue, &d.SauteurNumero, &d.ZoneSaut,
			&d.Alignement, &d.Ballon, &d.ZoneLongueur, &d.ZoneLargeur,
			&d.Utilisation, &d.Observation); err != nil {
			return liste, err
		}
		details[id] = d
	}
	for i := range liste {
		if d, ok := details[liste[i].ID]; ok {
			liste[i].Detail = d
		}
	}
	return liste, rows.Err()
}

// AjouterAction écrit la ligne, et son détail quand il y en a
// un, dans la même transaction.
func (b *Base) AjouterAction(ctx context.Context,
	a modele.Action) (modele.Action, error) {

	tx, err := b.pool.Begin(ctx)
	if err != nil {
		return a, err
	}
	defer tx.Rollback(ctx)

	err = tx.QueryRow(ctx,
		`INSERT INTO action
		   (match_id, temps, mt, type_code, possession,
		    numero, groupe, motif_code, commentaire)
		 VALUES ($1,$2,$3::mi_temps,$4,$5::possession,$6,$7,$8,$9)
		 RETURNING id`,
		a.MatchID, a.Temps, a.MT, a.TypeCode, a.Possession,
		a.Numero, a.Groupe, a.MotifCode, a.Commentaire).Scan(&a.ID)
	if err != nil {
		return a, err
	}

	if a.Detail != nil {
		d := a.Detail
		if _, err := tx.Exec(ctx,
			`INSERT INTO detail_phase
			   (action_id, issue, sauteur_numero, zone_saut, alignement,
			    ballon, zone_longueur, zone_largeur, utilisation, observation)
			 VALUES ($1,$2::issue,$3,$4,$5,$6,$7,$8,$9,$10)`,
			a.ID, d.Issue, d.SauteurNumero, d.ZoneSaut, d.Alignement,
			d.Ballon, d.ZoneLongueur, d.ZoneLargeur, d.Utilisation,
			d.Observation); err != nil {
			return a, err
		}
	}
	return a, tx.Commit(ctx)
}

func (b *Base) SupprimerAction(ctx context.Context, id int64) error {
	ct, err := b.pool.Exec(ctx, `DELETE FROM action WHERE id = $1`, id)
	if err != nil {
		return err
	}
	if ct.RowsAffected() == 0 {
		return pgx.ErrNoRows
	}
	return nil
}

// ---------------------------------------------------------
// Remplacements
// ---------------------------------------------------------

func (b *Base) Remplacements(ctx context.Context, matchID int) ([]modele.Remplacement, error) {
	rows, err := b.pool.Query(ctx,
		`SELECT id, temps::float8, mt, sortant, entrant, temporaire
		 FROM remplacement WHERE match_id = $1 ORDER BY temps`, matchID)
	if err != nil {
		return nil, err
	}
	return pgx.CollectRows(rows, pgx.RowToStructByPos[modele.Remplacement])
}

func (b *Base) AjouterRemplacement(ctx context.Context, matchID int,
	r modele.Remplacement) (modele.Remplacement, error) {

	err := b.pool.QueryRow(ctx,
		`INSERT INTO remplacement (match_id, temps, mt, sortant, entrant, temporaire)
		 VALUES ($1,$2,$3::mi_temps,$4,$5,$6)
		 RETURNING id`,
		matchID, r.Temps, r.MT, r.Sortant, r.Entrant, r.Temporaire).Scan(&r.ID)
	return r, err
}

// Titulaires : les numéros qui commencent le match, point de
// départ du calcul des temps de jeu.
func (b *Base) Titulaires(ctx context.Context, matchID int) ([]int, error) {
	rows, err := b.pool.Query(ctx,
		`SELECT numero FROM feuille_match
		 WHERE match_id = $1 AND titulaire ORDER BY numero`, matchID)
	if err != nil {
		return nil, err
	}
	return pgx.CollectRows(rows, pgx.RowTo[int])
}
