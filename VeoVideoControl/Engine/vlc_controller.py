from __future__ import annotations

import socket
import time
from dataclasses import dataclass


@dataclass
class VLCState:
    current_time: float
    paused: bool


class VLCController:
    def __init__(
        self,
        host: str = "127.0.0.1",
        port: int = 4212,
        timeout: float = 1.0,
    ) -> None:
        self.host = host
        self.port = port
        self.timeout = timeout

    def is_reachable(self) -> bool:
        try:
            with socket.create_connection(
                (self.host, self.port),
                timeout=self.timeout,
            ):
                return True

        except (ConnectionError, OSError):
            return False

    def _read_until_prompt(self, sock: socket.socket) -> str:
        chunks: list[bytes] = []

        while True:
            try:
                data = sock.recv(4096)
            except socket.timeout:
                break

            if not data:
                break

            chunks.append(data)

            contenu = b"".join(chunks)

            if contenu.rstrip().endswith(b">"):
                break

        return b"".join(chunks).decode(
            "utf-8",
            errors="replace",
        )

    def _send_command(self, command: str) -> str:
        with socket.create_connection(
            (self.host, self.port),
            timeout=self.timeout,
        ) as sock:
            sock.settimeout(self.timeout)

            # Consomme entièrement le message d’accueil et son prompt.
            self._read_until_prompt(sock)

            sock.sendall(
                (command + "\n").encode("utf-8")
            )

            # Lit ensuite uniquement la réponse à la commande.
            return self._read_until_prompt(sock)

    @staticmethod
    def _extract_value(response: str) -> str:
        lignes = [
            ligne.strip()
            for ligne in response.splitlines()
            if ligne.strip() and ligne.strip() != ">"
        ]

        if not lignes:
            raise RuntimeError(
                f"Réponse VLC vide ou invalide : {response!r}"
            )

        return lignes[0]

    def playpause(self) -> None:
        if self.is_playing():
            self._send_command("pause")
        else:
            self._send_command("play")

    def seek_relative(self, seconds: int) -> None:
        signe = "+" if seconds >= 0 else ""
        self._send_command(
            f"seek {signe}{seconds}"
        )

    def reset(self) -> None:
        self._send_command("seek 0")

        time.sleep(0.2)

        if self.get_playback_state() == "playing":
            self._send_command("pause")

    def get_time(self) -> float:
        response = self._send_command("get_time")
        valeur = self._extract_value(response)

        try:
            return float(valeur)
        except ValueError as error:
            raise RuntimeError(
                "Temps VLC introuvable dans la réponse : "
                f"{response!r}"
            ) from error

    def is_playing(self) -> bool:
        return self.get_playback_state() == "playing"

    def get_state(self) -> VLCState:
        return VLCState(
            current_time=self.get_time(),
            paused=not self.is_playing(),
        )

    def get_playback_state(self) -> str:
        response = self._send_command("status")

        for line in response.splitlines():
            line = line.strip().lower()

            if line == "( state playing )":
                return "playing"

            if line == "( state paused )":
                return "paused"

            if line == "( state stopped )":
                return "stopped"

        raise RuntimeError(
            "État de lecture VLC introuvable dans la réponse : "
            f"{response!r}"
        )