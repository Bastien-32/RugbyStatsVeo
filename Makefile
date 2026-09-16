# Commandes du projet. "make" seul affiche la liste.

CLASSEUR = Createur de match.xlsm

PYTHON = python3

# oletools n'est installe que dans l'environnement du projet.
PYTHON_VENV = .venv/bin/python

# Ce qui passe de club a main : ni Matchs, ni .gitignore.
ELEMENTS_REPORT = "Createur de match.xlsm" vba vba-recap outils \
	VeoVideoControl INSTALLATION.md INSTALLATION.txt \
	"logiciels à installer" Makefile

MESSAGE = chore: report des évolutions depuis club


.PHONY: help export verifier controler paquet report


help: ## affiche cette liste
	@grep -E '^[a-z]+:.*## ' Makefile \
		| sed 's/:.*## /\t/' \
		| awk -F'\t' '{printf "  make %-11s %s\n", $$1, $$2}'


# make export
export: ## exporte le code VBA dans vba/
	@$(PYTHON_VENV) outils/export_vba.py "$(CLASSEUR)"


# make verifier
verifier: ## controle que vba/ correspond au classeur
	@$(PYTHON_VENV) outils/export_vba.py "$(CLASSEUR)" --verifier


# make controler CLASSEUR="Createur de match copie.xlsm"
controler: ## controle qu'un classeur est vide de donnees
	@$(PYTHON) outils/preparer_distribution.py --controler \
		--classeur "$(CLASSEUR)"


# make paquet CLASSEUR="Createur de match copie.xlsm"
paquet: ## fabrique le dossier et le zip a distribuer
	@$(PYTHON) outils/preparer_distribution.py --zip \
		--classeur "$(CLASSEUR)"


# make report MESSAGE="feat: nouvelle palette"
report: ## reporte les evolutions de club vers main
	@set -e; \
	if [ "$$(git branch --show-current)" != "club" ]; then \
		echo "A lancer depuis la branche club."; \
		exit 1; \
	fi; \
	if [ -n "$$(git status --porcelain)" ]; then \
		echo "Des modifications ne sont pas validees : committer d'abord."; \
		exit 1; \
	fi; \
	git checkout main; \
	git checkout club -- $(ELEMENTS_REPORT); \
	if git diff --cached --quiet; then \
		echo "main est deja a jour."; \
	else \
		git commit -m "$(MESSAGE)"; \
	fi; \
	git checkout club; \
	echo "Report termine, retour sur club."
