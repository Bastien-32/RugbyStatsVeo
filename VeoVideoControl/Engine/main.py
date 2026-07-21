from collections import deque
from typing import Literal
from threading import Lock, Thread
from time import monotonic
import subprocess
import time


from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import uvicorn

from vlc_controller import VLCController


APP_NAME = "VeoVideoControl Engine"
APP_VERSION = "0.4.0"

VideoTarget = Literal["browser", "vlc"]

VLC_EXECUTABLE = (
    "/Applications/VLC.app/Contents/MacOS/VLC"
)

dernier_heartbeat = 0.0
heartbeat_recu = False

app = FastAPI(
    title=APP_NAME,
    version=APP_VERSION,
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

browser_video_state = {
    "source": "browser",
    "connected": False,
    "playerFound": False,
    "currentTime": 0.0,
    "paused": True,
    "url": "",
    "matchId": "",
}

browser_last_seen = 0.0

active_target: VideoTarget | None = None

command_queues = {
    "browser": deque(),
}

command_lock = Lock()

# Dernier signal reçu d’un classeur Stats Rugby.
dernier_heartbeat = 0.0

# Empêche l’arrêt automatique avant qu’un premier classeur
# ait réellement contacté le moteur.
heartbeat_recu = False

vlc = VLCController()


class BrowserVideoState(BaseModel):
    currentTime: float
    paused: bool
    url: str = ""
    matchId: str = ""
    source: VideoTarget = "browser"


@app.get("/")
def root():
    return {
        "application": APP_NAME,
        "version": APP_VERSION,
        "status": "running",
    }


# =========================================================
# ÉTAT DU LECTEUR
# =========================================================

@app.get("/video/state/active")
def get_active_video_state():

    if active_target is None:
        return {
            "source": "",
            "connected": False,
            "playerFound": False,
            "currentTime": 0.0,
            "paused": True,
            "url": "",
            "matchId": "",
        }

    return get_video_state(active_target)

@app.get("/video/state/{target}")
def get_video_state(target: VideoTarget):

    if target == "browser":
        return browser_video_state

    try:
        state = vlc.get_state()

        return {
            "source": "vlc",
            "connected": True,
            "playerFound": True,
            "currentTime": state.current_time,
            "paused": state.paused,
            "url": "",
            "matchId": "",
        }

    except (ConnectionError, OSError):
        return {
            "source": "vlc",
            "connected": False,
            "playerFound": False,
            "currentTime": 0.0,
            "paused": True,
            "url": "",
            "matchId": "",
        }

    except RuntimeError as error:
        raise HTTPException(
            status_code=503,
            detail=str(error),
        ) from error


# Ancienne route conservée pour les tests déjà effectués.
@app.get("/video/state")
def get_browser_video_state_legacy():
    return browser_video_state


@app.post("/video/state")
def set_browser_video_state(
    state: BrowserVideoState,
):
    global browser_video_state
    global browser_last_seen

    browser_video_state = {
        "source": "browser",
        "connected": True,
        "playerFound": True,
        "currentTime": state.currentTime,
        "paused": state.paused,
        "url": state.url,
        "matchId": state.matchId,
    }

    browser_last_seen = monotonic()

    return {"ok": True}

@app.post("/heartbeat")
def recevoir_heartbeat():
    global dernier_heartbeat
    global heartbeat_recu

    dernier_heartbeat = monotonic()
    heartbeat_recu = True

    return {"ok": True}

# =========================================================
# DÉTECTION ET CONNEXION AUTOMATIQUE
# =========================================================

def browser_is_available() -> bool:

    if not browser_video_state["connected"]:
        return False

    return monotonic() - browser_last_seen < 3

def vlc_is_available() -> bool:

    return vlc.is_reachable()


def launch_vlc() -> bool:

    try:
        subprocess.Popen(
            [
                VLC_EXECUTABLE,
                "--extraintf",
                "rc",
                "--rc-host",
                "127.0.0.1:4212",
            ],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            start_new_session=True,
        )

        # Attend au maximum 5 secondes que VLC réponde.
        for _ in range(20):
            time.sleep(0.25)

            if vlc_is_available():
                return True

        return False

    except (
        OSError,
        subprocess.SubprocessError,
    ):
        return False


@app.post("/connect-video")
def connect_video():

    global active_target

    # 1. Priorité à une vidéo Veo ouverte dans le navigateur.

    # Au démarrage du moteur, on laisse quelques secondes

    # à l’extension Chrome pour envoyer son premier état.

    if browser_is_available():

        active_target = "browser"

        return {
            "connected": True,
            "player": "Google Chrome",
            "target": "browser",
        }

    # 2. VLC répond déjà sur le port RC.
    if vlc_is_available():

        active_target = "vlc"

        return {
            "connected": True,
            "player": "VLC",
            "target": "vlc",
        }

    # 3. Rien de compatible n'est ouvert :
    #    lancement automatique de VLC en mode RC.
    if launch_vlc():

        active_target = "vlc"

        return {
            "connected": True,
            "player": "VLC",
            "target": "vlc",
            "launched": True,
        }

    active_target = None

    return {
        "connected": False,
        "player": "",
        "target": "",
        "reason": "no_player_found",
    }


@app.get("/connection-state")
def connection_state():

    if (
        active_target == "browser"
        and browser_is_available()
    ):
        return {
            "connected": True,
            "player": "Google Chrome",
            "target": "browser",
        }

    if (
        active_target == "vlc"
        and vlc_is_available()
    ):
        return {
            "connected": True,
            "player": "VLC",
            "target": "vlc",
        }

    return {
        "connected": False,
        "player": "",
        "target": "",
    }


# =========================================================
# COMMANDES VIDÉO
# =========================================================
@app.post("/command/active/{command}")
def send_active_command(command: str):

    if active_target is None:
        raise HTTPException(
            status_code=503,
            detail="Aucun lecteur vidéo connecté.",
        )

    return send_command(active_target, command)

@app.post("/command/{target}/{command}")
def send_command(
    target: VideoTarget,
    command: str,
):

    commandes_autorisees = {
        "playpause",
        "seek_minus_5",
        "seek_plus_5",
        "reset",
        "previous_action",
        "next_action",
    }

    if command not in commandes_autorisees:
        raise HTTPException(
            status_code=400,
            detail=f"Commande inconnue : {command}",
        )

    if target == "vlc":
        try:
            if command == "playpause":
                vlc.playpause()

            elif command == "seek_minus_5":
                vlc.seek_relative(-5)

            elif command == "seek_plus_5":
                vlc.seek_relative(5)

            elif command == "reset":
                vlc.reset()

        except (ConnectionError, OSError) as error:
            raise HTTPException(
                status_code=503,
                detail=(
                    "VLC n'est pas accessible sur "
                    "127.0.0.1:4212."
                ),
            ) from error

        except RuntimeError as error:
            raise HTTPException(
                status_code=503,
                detail=str(error),
            ) from error

        return {
            "ok": True,
            "target": target,
            "command": command,
        }

    with command_lock:
        command_queues["browser"].append(command)

    return {
        "ok": True,
        "target": target,
        "command": command,
    }


@app.get("/command/{target}")
def get_command(target: VideoTarget):

    # VLC reçoit ses commandes immédiatement ;
    # seule l'extension navigateur consomme une file.
    if target == "vlc":
        return {
            "target": "vlc",
            "command": None,
        }

    with command_lock:
        command = (
            command_queues["browser"].popleft()
            if command_queues["browser"]
            else None
        )

    return {
        "target": "browser",
        "command": command,
    }

# =========================================================
# SURVEILLANCE DES CLASSEURS STATS RUGBY
# =========================================================

def surveiller_heartbeat(server: uvicorn.Server) -> None:
    while not server.should_exit:
        time.sleep(0.5)

        # Le moteur ne s’arrête pas tant qu’il n’a jamais reçu
        # de signal d’un classeur Stats Rugby.
        if not heartbeat_recu:
            continue

        duree_sans_heartbeat = (
            monotonic() - dernier_heartbeat
        )

        if duree_sans_heartbeat >= 5:
            print(
                "Aucun heartbeat reçu depuis 5 secondes : "
                "arrêt de VeoVideoControl.",
                flush=True,
            )

            server.should_exit = True
            return

if __name__ == "__main__":
    configuration = uvicorn.Config(
        app,
        host="127.0.0.1",
        port=48652,
    )

    server = uvicorn.Server(configuration)

    Thread(
        target=surveiller_heartbeat,
        args=(server,),
        daemon=True,
    ).start()

    server.run()