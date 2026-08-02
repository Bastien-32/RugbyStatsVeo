from __future__ import annotations

import re
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
        timeout: float = 0.5,
        response_quiet_timeout: float = 0.03,
    ) -> None:
        self.host = host
        self.port = port
        self.timeout = timeout
        self.response_quiet_timeout = response_quiet_timeout

        # État conservé localement afin d’éviter une deuxième requête
        # VLC à chaque rafraîchissement du chrono.
        self._paused: bool = True
        self._state_known: bool = False

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
        first_block_received = False

        sock.settimeout(self.timeout)

        while True:
            try:
                data = sock.recv(4096)
            except socket.timeout:
                break

            if not data:
                break

            chunks.append(data)
            first_block_received = True

            content = b"".join(chunks)

            if content.rstrip().endswith(b">"):
                break

            if first_block_received:
                sock.settimeout(self.response_quiet_timeout)

        return b"".join(chunks).decode(
            "utf-8",
            errors="replace",
        )

    def _send_command(self, command: str) -> str:
        with socket.create_connection(
            (self.host, self.port),
            timeout=self.timeout,
        ) as sock:
            self._read_until_prompt(sock)

            sock.sendall(
                (command + "\n").encode("utf-8")
            )

            return self._read_until_prompt(sock)

    @staticmethod
    def _extract_numeric_value(response: str) -> float:
        """
        Recherche une ligne constituée uniquement d’un nombre.

        Cela évite de prendre un message, un écho de commande ou une
        confirmation VLC pour la valeur du temps.
        """
        for raw_line in response.splitlines():
            line = raw_line.strip()

            if not line or line == ">":
                continue

            if re.fullmatch(r"-?\d+(?:[.,]\d+)?", line):
                return float(line.replace(",", "."))

        raise RuntimeError(
            "Valeur numérique VLC introuvable dans la réponse : "
            f"{response!r}"
        )

    def playpause(self) -> None:
        # Dans l’interface RC de VLC, pause agit comme une bascule.
        self._send_command("pause")

        if self._state_known:
            self._paused = not self._paused
        else:
            # La prochaine lecture explicite déterminera l’état réel.
            self._state_known = False

    def seek_relative(self, seconds: int) -> None:
        current_time = self.get_time()
        target_time = max(0, int(round(current_time + seconds)))

        self._send_command(f"seek {target_time}")

    def reset(self) -> None:
        self._send_command("seek 0")
        time.sleep(0.1)

        state = self.get_playback_state()

        if state == "playing":
            self._send_command("pause")

        self._paused = True
        self._state_known = True

    def get_time(self) -> float:
        response = self._send_command("get_time")
        return self._extract_numeric_value(response)

    def is_playing(self) -> bool:
        return self.get_playback_state() == "playing"

    def get_state(self) -> VLCState:
        """
        Une seule requête VLC par rafraîchissement courant.

        Le temps est lu depuis VLC. L’état pause/lecture est conservé en
        mémoire et n’est redemandé que lorsqu’il n’est pas encore connu.
        """
        current_time = self.get_time()

        if not self._state_known:
            state = self.get_playback_state()
            self._paused = state != "playing"
            self._state_known = True

        return VLCState(
            current_time=current_time,
            paused=self._paused,
        )

    def get_playback_state(self) -> str:
        response = self._send_command("status")
        response_lower = response.lower()

        # VLC Windows peut répondre dans la langue du système.
        if (
            "pause" in response_lower
            and (
                "continue" in response_lower
                or "reprendre" in response_lower
            )
        ):
            self._paused = True
            self._state_known = True
            return "paused"

        states_numeric = {
            3: "playing",
            4: "paused",
            5: "stopped",
        }

        for raw_line in response.splitlines():
            line = raw_line.strip().lower()

            if line == "( state playing )":
                self._paused = False
                self._state_known = True
                return "playing"

            if line == "( state paused )":
                self._paused = True
                self._state_known = True
                return "paused"

            if line == "( state stopped )":
                self._paused = True
                self._state_known = True
                return "stopped"

            match = re.search(r"state:\s*(\d+)", line)

            if match:
                code = int(match.group(1))
                state = states_numeric.get(code, "stopped")

                self._paused = state != "playing"
                self._state_known = True

                return state

        raise RuntimeError(
            "Etat de lecture VLC introuvable dans la reponse : "
            f"{response!r}"
        )