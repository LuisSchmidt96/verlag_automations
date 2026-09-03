"""Startpunkt für den Buchdurchgang (auch der Einstieg der .exe)."""

from buchdurchgang.app import App


def main() -> None:
    App().mainloop()


if __name__ == "__main__":
    main()
