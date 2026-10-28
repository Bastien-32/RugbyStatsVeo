// Commande api : le serveur de RugbyStats.
//
// Toute la configuration passe par l'environnement, comme le
// veut un conteneur : rien à monter, rien à éditer dans
// l'image.
package main

import (
	"context"
	"errors"
	"log/slog"
	"net/http"
	"os"
	"os/signal"
	"syscall"
	"time"

	"rugbystats/api/internal/api"
	"rugbystats/api/internal/bd"
)

func main() {
	log := slog.New(slog.NewTextHandler(os.Stdout, &slog.HandlerOptions{
		Level: niveau(),
	}))

	url := variable("BASE_URL", "")
	if url == "" {
		log.Error("BASE_URL manquante")
		os.Exit(1)
	}

	ctx, arret := signal.NotifyContext(context.Background(),
		syscall.SIGINT, syscall.SIGTERM)
	defer arret()

	base, err := attendreLaBase(ctx, url, log)
	if err != nil {
		log.Error("connexion impossible", "err", err)
		os.Exit(1)
	}
	defer base.Fermer()

	serveur := &http.Server{
		Addr:              variable("ECOUTE", ":8080"),
		Handler:           api.Nouveau(base, log).Routes(),
		ReadHeaderTimeout: 5 * time.Second,
	}

	go func() {
		log.Info("écoute", "adresse", serveur.Addr)
		if err := serveur.ListenAndServe(); err != nil &&
			!errors.Is(err, http.ErrServerClosed) {
			log.Error("serveur arrêté", "err", err)
			os.Exit(1)
		}
	}()

	<-ctx.Done()
	log.Info("arrêt demandé, fermeture des connexions")

	fin, annule := context.WithTimeout(context.Background(), 10*time.Second)
	defer annule()
	_ = serveur.Shutdown(fin)
}

// attendreLaBase laisse à PostgreSQL le temps de se lever.
// Compose garantit déjà l'ordre, mais un redémarrage de la
// base ne doit pas tuer l'API.
func attendreLaBase(ctx context.Context, url string,
	log *slog.Logger) (*bd.Base, error) {

	var derniere error
	for essai := 1; essai <= 10; essai++ {
		base, err := bd.Ouvrir(ctx, url)
		if err == nil {
			return base, nil
		}
		derniere = err
		log.Warn("la base ne répond pas encore", "essai", essai)
		select {
		case <-ctx.Done():
			return nil, ctx.Err()
		case <-time.After(time.Second):
		}
	}
	return nil, derniere
}

func variable(nom, defaut string) string {
	if v := os.Getenv(nom); v != "" {
		return v
	}
	return defaut
}

func niveau() slog.Level {
	if os.Getenv("DEBUG") != "" {
		return slog.LevelDebug
	}
	return slog.LevelInfo
}
