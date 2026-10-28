// Package modele décrit le vocabulaire du jeu.
//
// Les types y sont fermés : une possession ne peut valoir que
// « Nous » ou « Adv », une issue de touche que G, P ou ND. Ce
// que le classeur garantissait par convention, le compilateur
// le garantit ici.
package modele

// Possession : qui tient le ballon au moment de l'action.
type Possession string

const (
	Nous Possession = "Nous"
	Adv  Possession = "Adv"
)

// Valide dit si la valeur fait partie du vocabulaire.
func (p Possession) Valide() bool {
	return p == Nous || p == Adv
}

// Inverse rend l'autre camp. Une action qui fait basculer la
// possession l'appelle plutôt que de recopier le test.
func (p Possession) Inverse() Possession {
	if p == Nous {
		return Adv
	}
	return Nous
}

// MiTemps : les deux périodes du match.
type MiTemps string

const (
	MT1 MiTemps = "MT1"
	MT2 MiTemps = "MT2"
)

func (m MiTemps) Valide() bool { return m == MT1 || m == MT2 }

// Ligne : la ligne de jeu d'un joueur sur une feuille de match.
// Elle se redéfinit match par match, un avant pouvant dépanner
// derrière.
type Ligne string

const (
	Avant       Ligne = "Avant"
	TroisQuarts Ligne = "Trois-quarts"
)

func (l Ligne) Valide() bool { return l == Avant || l == TroisQuarts }

// Issue d'une touche ou d'une mêlée.
//
//	G  le ballon revient à l'équipe qui lance ;
//	P  il reste à l'autre ;
//	ND la touche est rejouée par l'adversaire.
type Issue string

const (
	Gagnee   Issue = "G"
	Perdue   Issue = "P"
	PasDroit Issue = "ND"
)

func (i Issue) Valide() bool {
	return i == Gagnee || i == Perdue || i == PasDroit
}

// PossessionApres dit à qui revient le ballon après la phase,
// connaissant l'équipe qui la jouait. C'est la règle que la
// palette applique depuis une saison.
func (i Issue) PossessionApres(lanceur Possession) Possession {
	switch i {
	case Gagnee:
		return lanceur
	case Perdue:
		return lanceur.Inverse()
	default: // ND : rejouée par l'autre équipe
		return lanceur.Inverse()
	}
}

// Groupe : ce qui porte une action quand aucun joueur n'est
// nommé. Le classeur les écrit dans la colonne Joueur.
const (
	GroupeAvant       = "Avant"
	GroupeTroisQuarts = "Trois-quarts"
	GroupeCollectif   = "Collectif"
	GroupeInconnu     = "?"
)

// GroupeValide écarte tout ce qui n'est pas un de ces quatre
// libellés : un nom de joueur n'a rien à faire là.
func GroupeValide(g string) bool {
	switch g {
	case GroupeAvant, GroupeTroisQuarts, GroupeCollectif, GroupeInconnu:
		return true
	}
	return false
}

// ---------------------------------------------------------
// Entités
// ---------------------------------------------------------

type Joueur struct {
	ID     int    `json:"id"`
	Nom    string `json:"nom"`
	Prenom string `json:"prenom"`
	Ligne  *Ligne `json:"ligne,omitempty"`
	Actif  bool   `json:"actif"`
}

// NomComplet : ce qui s'affiche sur les boutons et au journal.
func (j Joueur) NomComplet() string {
	if j.Prenom == "" {
		return j.Nom
	}
	return j.Nom + " " + j.Prenom
}

type Match struct {
	ID         int    `json:"id"`
	Saison     string `json:"saison"`
	DateMatch  string `json:"date_match"`
	Adversaire string `json:"adversaire"`
	Categorie  string `json:"categorie"`
	Lieu       string `json:"lieu"`
	Phase      string `json:"phase"`
	Journee    *int   `json:"journee,omitempty"`
	SensMT1    string `json:"sens_mt1"`
	VideoRef   string `json:"video_ref"`
}

// LigneFeuille : un numéro de maillot porté par un joueur, dans
// une ligne donnée, pour ce match-là.
type LigneFeuille struct {
	Numero      int    `json:"numero"`
	JoueurID    int    `json:"joueur_id"`
	NomComplet  string `json:"nom_complet"`
	Ligne       Ligne  `json:"ligne"`
	Titulaire   bool   `json:"titulaire"`
}

// Action : une ligne du journal.
//
// Le temps est celui de la vidéo, en secondes. C'est lui qui
// permet de revenir à l'image, et c'est sur lui que se calcule
// l'état du jeu à un instant donné.
type Action struct {
	ID          int64      `json:"id"`
	MatchID     int        `json:"match_id"`
	Temps       float64    `json:"temps"`
	MT          MiTemps    `json:"mt"`
	TypeCode    string     `json:"type_code"`
	Libelle     string     `json:"libelle,omitempty"`
	Possession  Possession `json:"possession"`
	Numero      *int       `json:"numero,omitempty"`
	Groupe      *string    `json:"groupe,omitempty"`
	MotifCode   *string    `json:"motif_code,omitempty"`
	MotifLib    string     `json:"motif_libelle,omitempty"`
	Commentaire string     `json:"commentaire,omitempty"`
	Detail      *Detail    `json:"detail,omitempty"`
}

// Detail accompagne une touche ou une mêlée.
type Detail struct {
	Issue         Issue  `json:"issue"`
	SauteurNumero *int   `json:"sauteur_numero,omitempty"`
	ZoneSaut      string `json:"zone_saut,omitempty"`
	Alignement    string `json:"alignement,omitempty"`
	Ballon        string `json:"ballon,omitempty"`
	ZoneLongueur  string `json:"zone_longueur,omitempty"`
	ZoneLargeur   string `json:"zone_largeur,omitempty"`
	Utilisation   string `json:"utilisation,omitempty"`
	Observation   string `json:"observation,omitempty"`
}

type Remplacement struct {
	ID         int64   `json:"id"`
	Temps      float64 `json:"temps"`
	MT         MiTemps `json:"mt"`
	Sortant    int     `json:"sortant"`
	Entrant    int     `json:"entrant"`
	Temporaire bool    `json:"temporaire"`
}

// ---------------------------------------------------------
// Référentiels
// ---------------------------------------------------------

type TypeAction struct {
	Code         string `json:"code"`
	Famille      string `json:"famille"`
	Libelle      string `json:"libelle"`
	Rang         int    `json:"rang"`
	OuvreJeu     bool   `json:"ouvre_jeu"`
	FermeJeu     bool   `json:"ferme_jeu"`
	ArretAuto    bool   `json:"arret_auto"`
	Bascule      bool   `json:"bascule"`
	AttendJoueur bool   `json:"attend_joueur"`
}

type FamilleAction struct {
	Code    string `json:"code"`
	Libelle string `json:"libelle"`
	Rang    int    `json:"rang"`
}

type MotifFaute struct {
	Code       string `json:"code"`
	Libelle    string `json:"libelle"`
	Famille    string `json:"famille"`
	Penalite   bool   `json:"penalite"`
	CoupFranc  bool   `json:"coup_franc"`
	Rang       int    `json:"rang"`
}

// Referentiels : tout ce dont l'interface a besoin pour se
// construire, en un seul appel.
type Referentiels struct {
	Familles []FamilleAction `json:"familles"`
	Types    []TypeAction    `json:"types"`
	Motifs   []MotifFaute    `json:"motifs"`
}
