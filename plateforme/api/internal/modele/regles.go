package modele

import "sort"

// =========================================================
// Les règles du journal
//
// Elles vivent ici plutôt que dans l'interface : une
// statistique calculée par le serveur et une palette qui
// grise ses boutons doivent dire la même chose.
// =========================================================

// EtatJeu dit si le jeu est ouvert à un instant donné de la
// vidéo, pour la mi-temps considérée.
//
// La borne compte : on ne regarde que ce qui précède la
// position courante. Revenir en arrière pour compléter une
// phase doit rendre l'état du jeu de cet instant-là, pas
// celui de la fin du match.
func EtatJeu(actions []Action, types map[string]TypeAction,
	mt MiTemps, temps float64) bool {

	const marge = 0.5 // deux lectures du lecteur ne coïncident jamais

	meilleur := -1.0
	ouvert := false

	for _, a := range actions {
		if a.MT != mt || a.Temps > temps+marge {
			continue
		}
		t, connu := types[a.TypeCode]
		if !connu || (!t.OuvreJeu && !t.FermeJeu) {
			continue
		}
		if a.Temps > meilleur {
			meilleur = a.Temps
			ouvert = t.OuvreJeu
		}
	}
	return ouvert
}

// Sequence : un temps de possession continu, d'une reprise à
// l'arrêt ou au changement de camp.
type Sequence struct {
	Equipe Possession `json:"equipe"`
	MT     MiTemps    `json:"mt"`
	Debut  float64    `json:"debut"`
	Fin    float64    `json:"fin"`
}

func (s Sequence) Duree() float64 { return s.Fin - s.Debut }

// SequencesPossession découpe le match en temps de possession.
//
// Une séquence s'ouvre à la reprise, se ferme à l'arrêt, et se
// coupe en deux quand le ballon change de camp en jeu courant.
// Celle qui reste ouverte en fin de journal n'est pas comptée :
// elle peut n'être qu'une saisie en cours.
func SequencesPossession(actions []Action,
	types map[string]TypeAction) []Sequence {

	var seqs []Sequence

	for _, mt := range []MiTemps{MT1, MT2} {
		lignes := make([]Action, 0, len(actions))
		for _, a := range actions {
			if a.MT == mt {
				lignes = append(lignes, a)
			}
		}
		sort.Slice(lignes, func(i, j int) bool {
			return lignes[i].Temps < lignes[j].Temps
		})

		ouverte := false
		arrete := true
		equipe := Possession("")
		debut := 0.0

		for _, a := range lignes {
			t := types[a.TypeCode]

			switch {
			case t.FermeJeu:
				if ouverte {
					seqs = append(seqs, Sequence{equipe, mt, debut, a.Temps})
				}
				ouverte, arrete = false, true

			case t.OuvreJeu:
				if ouverte {
					seqs = append(seqs, Sequence{equipe, mt, debut, a.Temps})
				}
				ouverte, arrete = true, false
				equipe, debut = a.Possession, a.Temps

			case !arrete:
				if !ouverte {
					// Une reprise manque au journal : on
					// rattrape la séquence plutôt que de la
					// perdre.
					ouverte = true
					equipe, debut = a.Possession, a.Temps
				} else if a.Possession != equipe {
					seqs = append(seqs, Sequence{equipe, mt, debut, a.Temps})
					equipe, debut = a.Possession, a.Temps
				}
			}
		}
	}
	return seqs
}

// Tranches de durée, celles du classeur.
var BornesTranches = []float64{20, 40, 60, 90}

var NomsTranches = []string{"0-20 s", "20-40 s", "40-60 s", "60-90 s", "> 90 s"}

func Tranche(duree float64) int {
	for i, b := range BornesTranches {
		if duree < b {
			return i
		}
	}
	return len(BornesTranches)
}

// StatsPossession : le tableau des durées, par camp.
type StatsPossession struct {
	Tranches   []string   `json:"tranches"`
	Nous       []int      `json:"nous"`
	Adv        []int      `json:"adv"`
	CumulNous  float64    `json:"cumul_nous"`
	CumulAdv   float64    `json:"cumul_adv"`
	Sequences  []Sequence `json:"sequences"`
}

func CalculerPossession(actions []Action,
	types map[string]TypeAction) StatsPossession {

	seqs := SequencesPossession(actions, types)
	s := StatsPossession{
		Tranches:  NomsTranches,
		Nous:      make([]int, len(NomsTranches)),
		Adv:       make([]int, len(NomsTranches)),
		Sequences: seqs,
	}
	for _, q := range seqs {
		r := Tranche(q.Duree())
		if q.Equipe == Nous {
			s.Nous[r]++
			s.CumulNous += q.Duree()
		} else {
			s.Adv[r]++
			s.CumulAdv += q.Duree()
		}
	}
	return s
}

// ---------------------------------------------------------
// Temps de jeu
// ---------------------------------------------------------

// Passage : un séjour sur le terrain. Un joueur peut en avoir
// plusieurs — remplacement temporaire en première ligne,
// retour après soin.
type Passage struct {
	Entre float64
	Sorti *float64
}

// TempsParNumero reconstitue les passages depuis la feuille et
// les remplacements, puis en tire un cumul par numéro.
//
// Le temps n'est jamais stocké : il resterait à recalculer à
// la moindre correction d'un remplacement.
func TempsParNumero(titulaires []int, remps []Remplacement,
	fin float64) map[int]float64 {

	passages := map[int][]Passage{}
	for _, n := range titulaires {
		passages[n] = []Passage{{Entre: 0}}
	}

	tri := make([]Remplacement, len(remps))
	copy(tri, remps)
	sort.Slice(tri, func(i, j int) bool { return tri[i].Temps < tri[j].Temps })

	for _, r := range tri {
		// Le sortant ferme son passage en cours.
		liste := passages[r.Sortant]
		for i := len(liste) - 1; i >= 0; i-- {
			if liste[i].Sorti == nil {
				t := r.Temps
				liste[i].Sorti = &t
				break
			}
		}
		passages[r.Sortant] = liste
		passages[r.Entrant] = append(passages[r.Entrant],
			Passage{Entre: r.Temps})
	}

	total := map[int]float64{}
	for n, liste := range passages {
		somme := 0.0
		for _, p := range liste {
			sortie := fin
			if p.Sorti != nil {
				sortie = *p.Sorti
			}
			if sortie > p.Entre {
				somme += sortie - p.Entre
			}
		}
		total[n] = somme
	}
	return total
}
