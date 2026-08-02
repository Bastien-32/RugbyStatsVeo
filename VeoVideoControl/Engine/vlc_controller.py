from __future__ import annotations

import base64
import json
from dataclasses import dataclass
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen


@dataclass
class VLCState:
    current_time: float
    paused: bool


class VLCController:
    """Pilote VLC avec son interface HTTP locale.

    L'interface HTTP renvoie du JSON et évite les différences de langue
    et de comportement rencontrées avec l'interface texte RC.
    """

    def __init__(
        self,
        host: str = "127.0.0.1",
        port: int = 4212,
        password: str = "statsrugby-local",
        timeout: float = 0.6,
    ) -> None:
        self.host = host
        self.port = port
        self.password = password
        self.timeout = timeout
        self.base_url = f"http://{self.host}:{self.port}"

        credentials = f":{self.password}".encode("utf-8")
        encoded_credentials = base64.b64encode(credentials).decode("ascii")
        self.authorization_header = f"Basic {encoded_credentials}"

    def _request_status(
        self,
        command: str | None = None,
        value: str | None = None,
    ) -> dict[str, Any]:
        parameters: dict[str, str] = {}

        if command is not None:
            parameters["command"] = command

        if value is not None:
            parameters["val"] = value

        url = f"{self.base_url}/requests/status.json"

        if parameters:
            url = f"{url}?{urlencode(parameters)}"

        request = Request(
            url,
            headers={
                "Authorization": self.authorization_header,
                "Accept": "application/json",
                "Cache-Control": "no-cache",
            },
            method="GET",
        )

        try:
            with urlopen(request, timeout=self.timeout) as response:
                payload = response.read().decode("utf-8", errors="replace")

        except HTTPError as error:
            if error.code == 401:
                raise RuntimeError(
                    "Mot de passe de l'interface HTTP VLC incorrect."
                ) from error

            raise ConnectionError(
                f"Erreur HTTP VLC : {error.code}"
            ) from error

        except (URLError, TimeoutError, OSError) as error:
            raise ConnectionError(
                f"VLC n'est pas accessible sur {self.host}:{self.port}."
            ) from error

        try:
            result = json.loads(payload)
        except json.JSONDecodeError as error:
            raise RuntimeError(
                "Réponse JSON VLC invalide : "
                f"{payload!r}"
            ) from error

        if not isinstance(result, dict):
            raise RuntimeError(
                "Réponse VLC inattendue : "
                f"{result!r}"
            )

        return result

    def is_reachable(self) -> bool:
        try:
            self._request_status()
            return True
        except (ConnectionError, RuntimeError):
            return False

    def playpause(self) -> None:
        self._request_status(command="pl_pause")

    def seek_relative(self, seconds: int) -> None:
        sign = "+" if seconds >= 0 else "-"
        value = f"{sign}{abs(seconds)}S"
        self._request_status(command="seek", value=value)

    def reset(self) -> None:
        state = self._request_status(command="seek", value="0")

        if str(state.get("state", "")).lower() == "playing":
            self._request_status(command="pl_pause")

    def get_time(self) -> float:
        state = self._request_status()

        try:
            return float(state.get("time", 0.0))
        except (TypeError, ValueError) as error:
            raise RuntimeError(
                "Temps VLC introuvable dans la réponse : "
                f"{state!r}"
            ) from error

    def is_playing(self) -> bool:
        state = self._request_status()
        return str(state.get("state", "")).lower() == "playing"

    def get_state(self) -> VLCState:
        state = self._request_status()

        try:
            current_time = float(state.get("time", 0.0))
        except (TypeError, ValueError) as error:
            raise RuntimeError(
                "Temps VLC introuvable dans la réponse : "
                f"{state!r}"
            ) from error

        playback_state = str(state.get("state", "")).lower()

        return VLCState(
            current_time=current_time,
            paused=playback_state != "playing",
        )
