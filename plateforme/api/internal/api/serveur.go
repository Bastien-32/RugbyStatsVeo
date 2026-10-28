// Package api expose le journal et ses calculs en HTTP.
package api

import (
	"encoding/json"
	"errors"
	"log/slog"
	"net/http"
	"strconv"

	"github.com/jackc/pgx/v5"

	"rugbystats/api/internal/bd"
	"rugbystats/api/internal/modele"
)

// Un seul club pour l'instant. Il viendra de la session quand
// l'authentification existera ; d'ici là, le découpage des
// requêtes par club est déjà en place.
const clubParDefaut = 1

type Serveur struct {
	base *bd.Base
	log  *slog.Logger
}

func Nouveau(base *bd.Base, log *slog.Logger) *Serveur {
	return &Serveur{base: base, log: log}
}

// Routes : le routeur standard de Go suffit depuis la 1.22,
// les motifs acceptant verbe et variables de chemin.
func (s *Serveur) Routes() http.Handler {
	mux := http.NewServeMux()

	mux.HandleFunc("GET /sante", s.sante)
	mux.HandleFunc("GET /referentiels", s.referentiels)

	mux.HandleFunc("GET /joueurs", s.joueurs)
	mux.HandleFunc("POST /joueurs", s.creerJoueur)

	mux.HandleFunc("GET /matchs", s.matchs)
	mux.HandleFunc("POST /matchs", s.creerMatch)
	mux.HandleFunc("GET /matchs/{id}", s.match)

	mux.HandleFunc("GET /matchs/{id}/feuille", s.feuille)
	mux.HandleFunc("PUT /matchs/{id}/feuille", s.poserFeuille)

	mux.HandleFunc("GET /matchs/{id}/actions", s.actions)
	mux.HandleFunc("POST /matchs/{id}/actions", s.ajouterAction)
	mux.HandleFunc("DELETE /actions/{id}", s.supprimerAction)

	mux.HandleFunc("GET /matchs/{id}/remplacements", s.remplacements)
	mux.HandleFunc("POST /matchs/{id}/remplacements", s.ajouterRemplacement)

	mux.HandleFunc("GET /matchs/{id}/statistiques", s.statistiques)

	return journalisation(s.log, cors(mux))
}

// ---------------------------------------------------------
// Santé et référentiels
// ---------------------------------------------------------

func (s *Serveur) sante(w http.ResponseWriter, r *http.Request) {
	if err := s.base.Verifier(r.Context()); err != nil {
		erreur(w, http.StatusServiceUnavailable, "la base ne répond pas")
		return
	}
	rendre(w, http.StatusOK, map[string]string{"etat": "ok"})
}

func (s *Serveur) referentiels(w http.ResponseWriter, r *http.Request) {
	ref, err := s.base.Referentiels(r.Context())
	if err != nil {
		s.panne(w, err)
		return
	}
	rendre(w, http.StatusOK, ref)
}

// ---------------------------------------------------------
// Joueurs
// ---------------------------------------------------------

func (s *Serveur) joueurs(w http.ResponseWriter, r *http.Request) {
	liste, err := s.base.Joueurs(r.Context(), clubParDefaut)
	if err != nil {
		s.panne(w, err)
		return
	}
	rendre(w, http.StatusOK, liste)
}

func (s *Serveur) creerJoueur(w http.ResponseWriter, r *http.Request) {
	var j modele.Joueur
	if !lire(w, r, &j) {
		return
	}
	if j.Nom == "" {
		erreur(w, http.StatusBadRequest, "le nom est obligatoire")
		return
	}
	if j.Ligne != nil && !j.Ligne.Valide() {
		erreur(w, http.StatusBadRequest, "ligne inconnue : attendu Avant ou Trois-quarts")
		return
	}
	cree, err := s.base.CreerJoueur(r.Context(), clubParDefaut, j)
	if err != nil {
		s.panne(w, err)
		return
	}
	rendre(w, http.StatusCreated, cree)
}

// ---------------------------------------------------------
// Matchs
// ---------------------------------------------------------

func (s *Serveur) matchs(w http.ResponseWriter, r *http.Request) {
	liste, err := s.base.Matchs(r.Context(), clubParDefaut)
	if err != nil {
		s.panne(w, err)
		return
	}
	rendre(w, http.StatusOK, liste)
}

func (s *Serveur) match(w http.ResponseWriter, r *http.Request) {
	id, ok := identifiant(w, r)
	if !ok {
		return
	}
	m, err := s.base.Match(r.Context(), id)
	if errors.Is(err, pgx.ErrNoRows) {
		erreur(w, http.StatusNotFound, "match introuvable")
		return
	}
	if err != nil {
		s.panne(w, err)
		return
	}
	rendre(w, http.StatusOK, m)
}

func (s *Serveur) creerMatch(w http.ResponseWriter, r *http.Request) {
	var m modele.Match
	if !lire(w, r, &m) {
		return
	}
	if m.Adversaire == "" || m.DateMatch == "" || m.Saison == "" {
		erreur(w, http.StatusBadRequest,
			"saison, date et adversaire sont obligatoires")
		return
	}
	if m.Lieu == "" {
		m.Lieu = "Domicile"
	}
	if m.SensMT1 == "" {
		m.SensMT1 = "droite"
	}
	cree, err := s.base.CreerMatch(r.Context(), clubParDefaut, m)
	if err != nil {
		s.panne(w, err)
		return
	}
	rendre(w, http.StatusCreated, cree)
}

// ---------------------------------------------------------
// Feuille de match
// ---------------------------------------------------------

func (s *Serveur) feuille(w http.ResponseWriter, r *http.Request) {
	id, ok := identifiant(w, r)
	if !ok {
		return
	}
	lignes, err := s.base.Feuille(r.Context(), id)
	if err != nil {
		s.panne(w, err)
		return
	}
	rendre(w, http.StatusOK, lignes)
}

func (s *Serveur) poserFeuille(w http.ResponseWriter, r *http.Request) {
	id, ok := identifiant(w, r)
	if !ok {
		return
	}
	var lignes []modele.LigneFeuille
	if !lire(w, r, &lignes) {
		return
	}
	for _, l := range lignes {
		if !l.Ligne.Valide() {
			erreur(w, http.StatusBadRequest,
				"numéro "+strconv.Itoa(l.Numero)+" : ligne de jeu inconnue")
			return
		}
	}
	if err := s.base.PoserFeuille(r.Context(), id, lignes); err != nil {
		s.panne(w, err)
		return
	}
	rendre(w, http.StatusOK, map[string]int{"lignes": len(lignes)})
}

// ---------------------------------------------------------
// Journal
// ---------------------------------------------------------

func (s *Serveur) actions(w http.ResponseWriter, r *http.Request) {
	id, ok := identifiant(w, r)
	if !ok {
		return
	}
	liste, err := s.base.Actions(r.Context(), id)
	if err != nil {
		s.panne(w, err)
		return
	}
	if liste == nil {
		liste = []modele.Action{}
	}
	rendre(w, http.StatusOK, liste)
}

func (s *Serveur) ajouterAction(w http.ResponseWriter, r *http.Request) {
	id, ok := identifiant(w, r)
	if !ok {
		return
	}
	var a modele.Action
	if !lire(w, r, &a) {
		return
	}
	a.MatchID = id

	if !a.Possession.Valide() {
		erreur(w, http.StatusBadRequest, "possession inconnue : attendu Nous ou Adv")
		return
	}
	if !a.MT.Valide() {
		erreur(w, http.StatusBadRequest, "mi-temps inconnue : attendu MT1 ou MT2")
		return
	}
	if a.Temps < 0 {
		erreur(w, http.StatusBadRequest, "le temps vidéo ne peut pas être négatif")
		return
	}
	// Un joueur nommé ou un groupe, jamais les deux : la base
	// le refuserait, autant le dire clairement ici.
	if a.Numero != nil && a.Groupe != nil {
		erreur(w, http.StatusBadRequest,
			"une action porte un numéro ou un groupe, pas les deux")
		return
	}
	if a.Groupe != nil && !modele.GroupeValide(*a.Groupe) {
		erreur(w, http.StatusBadRequest,
			"groupe inconnu : attendu Avant, Trois-quarts, Collectif ou ?")
		return
	}
	if a.Detail != nil && !a.Detail.Issue.Valide() {
		erreur(w, http.StatusBadRequest, "issue inconnue : attendu G, P ou ND")
		return
	}

	cree, err := s.base.AjouterAction(r.Context(), a)
	if err != nil {
		s.panne(w, err)
		return
	}
	rendre(w, http.StatusCreated, cree)
}

func (s *Serveur) supprimerAction(w http.ResponseWriter, r *http.Request) {
	id, err := strconv.ParseInt(r.PathValue("id"), 10, 64)
	if err != nil {
		erreur(w, http.StatusBadRequest, "identifiant invalide")
		return
	}
	err = s.base.SupprimerAction(r.Context(), id)
	if errors.Is(err, pgx.ErrNoRows) {
		erreur(w, http.StatusNotFound, "action introuvable")
		return
	}
	if err != nil {
		s.panne(w, err)
		return
	}
	w.WriteHeader(http.StatusNoContent)
}

// ---------------------------------------------------------
// Remplacements
// ---------------------------------------------------------

func (s *Serveur) remplacements(w http.ResponseWriter, r *http.Request) {
	id, ok := identifiant(w, r)
	if !ok {
		return
	}
	liste, err := s.base.Remplacements(r.Context(), id)
	if err != nil {
		s.panne(w, err)
		return
	}
	if liste == nil {
		liste = []modele.Remplacement{}
	}
	rendre(w, http.StatusOK, liste)
}

func (s *Serveur) ajouterRemplacement(w http.ResponseWriter, r *http.Request) {
	id, ok := identifiant(w, r)
	if !ok {
		return
	}
	var rp modele.Remplacement
	if !lire(w, r, &rp) {
		return
	}
	if rp.Sortant == rp.Entrant {
		erreur(w, http.StatusBadRequest,
			"un joueur ne peut pas se remplacer lui-même")
		return
	}
	if !rp.MT.Valide() {
		erreur(w, http.StatusBadRequest, "mi-temps inconnue")
		return
	}
	cree, err := s.base.AjouterRemplacement(r.Context(), id, rp)
	if err != nil {
		s.panne(w, err)
		return
	}
	rendre(w, http.StatusCreated, cree)
}

// ---------------------------------------------------------
// Statistiques
// ---------------------------------------------------------

type statistiques struct {
	Possession  modele.StatsPossession `json:"possession"`
	TempsDeJeu  []tempsJoueur          `json:"temps_de_jeu"`
	FinJournal  float64                `json:"fin_journal"`
}

type tempsJoueur struct {
	Numero     int     `json:"numero"`
	NomComplet string  `json:"nom_complet"`
	Ligne      string  `json:"ligne"`
	Secondes   float64 `json:"secondes"`
}

func (s *Serveur) statistiques(w http.ResponseWriter, r *http.Request) {
	id, ok := identifiant(w, r)
	if !ok {
		return
	}
	ctx := r.Context()

	types, err := s.base.TypesParCode(ctx)
	if err != nil {
		s.panne(w, err)
		return
	}
	actions, err := s.base.Actions(ctx, id)
	if err != nil {
		s.panne(w, err)
		return
	}
	remps, err := s.base.Remplacements(ctx, id)
	if err != nil {
		s.panne(w, err)
		return
	}
	titulaires, err := s.base.Titulaires(ctx, id)
	if err != nil {
		s.panne(w, err)
		return
	}
	feuille, err := s.base.Feuille(ctx, id)
	if err != nil {
		s.panne(w, err)
		return
	}

	// La fin du journal sert de borne : un joueur jamais
	// remplacé a joué jusqu'à la dernière action saisie.
	fin := 0.0
	for _, a := range actions {
		if a.Temps > fin {
			fin = a.Temps
		}
	}

	cumuls := modele.TempsParNumero(titulaires, remps, fin)
	temps := make([]tempsJoueur, 0, len(cumuls))
	for _, l := range feuille {
		if sec, joue := cumuls[l.Numero]; joue && sec > 0 {
			temps = append(temps, tempsJoueur{
				Numero: l.Numero, NomComplet: l.NomComplet,
				Ligne: string(l.Ligne), Secondes: sec,
			})
		}
	}

	rendre(w, http.StatusOK, statistiques{
		Possession: modele.CalculerPossession(actions, types),
		TempsDeJeu: temps,
		FinJournal: fin,
	})
}

// ---------------------------------------------------------
// Utilitaires
// ---------------------------------------------------------

func identifiant(w http.ResponseWriter, r *http.Request) (int, bool) {
	id, err := strconv.Atoi(r.PathValue("id"))
	if err != nil || id <= 0 {
		erreur(w, http.StatusBadRequest, "identifiant invalide")
		return 0, false
	}
	return id, true
}

func lire(w http.ResponseWriter, r *http.Request, cible any) bool {
	dec := json.NewDecoder(r.Body)
	dec.DisallowUnknownFields()
	if err := dec.Decode(cible); err != nil {
		erreur(w, http.StatusBadRequest, "corps de requête illisible : "+err.Error())
		return false
	}
	return true
}

func rendre(w http.ResponseWriter, code int, corps any) {
	w.Header().Set("Content-Type", "application/json; charset=utf-8")
	w.WriteHeader(code)
	_ = json.NewEncoder(w).Encode(corps)
}

func erreur(w http.ResponseWriter, code int, message string) {
	rendre(w, code, map[string]string{"erreur": message})
}

// panne journalise le détail et n'en dit rien au client : un
// message de PostgreSQL n'a pas à sortir de la machine.
func (s *Serveur) panne(w http.ResponseWriter, err error) {
	s.log.Error("erreur serveur", "err", err)
	erreur(w, http.StatusInternalServerError, "erreur interne")
}

func cors(suivant http.Handler) http.Handler {
	return http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		// Le front est servi par un autre conteneur, donc une
		// autre origine en développement.
		w.Header().Set("Access-Control-Allow-Origin", "*")
		w.Header().Set("Access-Control-Allow-Methods",
			"GET, POST, PUT, DELETE, OPTIONS")
		w.Header().Set("Access-Control-Allow-Headers", "Content-Type")
		if r.Method == http.MethodOptions {
			w.WriteHeader(http.StatusNoContent)
			return
		}
		suivant.ServeHTTP(w, r)
	})
}

func journalisation(log *slog.Logger, suivant http.Handler) http.Handler {
	return http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		log.Debug("requête", "methode", r.Method, "chemin", r.URL.Path)
		suivant.ServeHTTP(w, r)
	})
}
