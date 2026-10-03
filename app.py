"""Avvio di Formazioni PZZ. Il codice sta nel pacchetto formazioni/."""

import multiprocessing

from formazioni.ui.app import main

if __name__ == "__main__":
    # Nell'exe i processi del batch parallelo rilanciano l'eseguibile stesso:
    # freeze_support li intercetta prima che aprano un'altra finestra.
    multiprocessing.freeze_support()
    main()
