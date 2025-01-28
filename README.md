# Digital Reports

## 🇮🇹 Versione Italiana  
**[🇬🇧 English version available!](#english-version)**

## 🔍 Overview

**Digital Reports** è un'applicazione desktop progettata per aiutare il corpo di Polizia Locale a tenere traccia dei report in maniera digitale. Gestire turni di lavoro, attività svolte, mezzi utilizzati, persone identificate, infrazioni accertate e annotazioni non è mai stato così facile!

### Funzionalità principali

- **Scrittura dei report**: Gestione di turni, attività, mezzi, persone, infrazioni e annotazioni.
- **Calcolo delle ore lavorate**: Automatizza il conteggio delle ore.
- **Modifica dei dati**: Aggiorna e salva direttamente nel database SQLite.

## 🛠️ Setup e librerie

Non preoccuparti, il programma configura automaticamente il database SQLite al primo avvio! Ti servono solo le seguenti librerie Python (se non le hai, installale con `pip install <nome_libreria>`):

```python
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from PIL import Image, ImageTk
import sqlite3
from datetime import datetime, timedelta
import os
import ast
import openpyxl
import locale
import webbrowser
```

## ⚠️ REDACTED: cosa manca?

Per motivi di sicurezza, alcuni file e campi sensibili sono stati esclusi dal progetto. Questi includono risorse come immagini, file Excel per i report settimanali o giornalieri e altre configurazioni. Cercando la parola **`REDACTED`** nel codice, troverai tutti i punti in cui mancano dati o file, insieme a una descrizione di cosa dovrebbero contenere.

Esempi di dati REDACTED:

- **Campi degli agenti**: Informazioni sensibili sui nomi degli agenti.
- **Infrazioni**: Dettagli specifici sugli articoli violati.
- **Report Excel e PDF**: File necessari per generare report giornalieri, mensili, ecc.

> Nota: Non includo questi file per garantire la sicurezza e la riservatezza dei dati sensibili di un corpo di polizia.

## 🔧 Come creare un eseguibile (.exe)

Se vuoi distribuire il programma senza dover installare Python, puoi usare **PyInstaller** per creare un file `.exe`:

1. Assicurati di avere PyInstaller installato:

   ```bash
   pip install pyinstaller
   ```

2. Apri un terminale nella cartella del progetto e lancia il comando:

   ```bash
   pyinstaller --onefile --icon=icon.ico main.py
   ```

   - `--onefile`: Crea un unico file eseguibile.
   - `--icon=icon.ico`: Aggiunge un'icona personalizzata all'eseguibile (se presente).

3. Troverai il file `.exe` nella cartella `dist/`.

> Nota: Ricorda che l'eseguibile includerà tutto il codice, quindi valuta accorgimenti per proteggere informazioni sensibili.

## 🚫 Disclaimer

Il codice potrebbe contenere errori, sia a livello grafico, sia di funzionalità, sia di sicurezza. Non mi assumo alcuna responsabilità per eventuali problemi derivanti dal suo utilizzo. Tuttavia, qualsiasi modifica, ottimizzazione o miglioramento è ben accetto! ❤️

## 🔬 Curiosità tecniche

La tecnologia che alimenta **Digital Reports**:

- **Python**: Il cervello dietro il progetto.
- **Tkinter**: Per l'interfaccia grafica.
- **SQLite**: Per la gestione del database locale.
- **openpyxl**: Per l'integrazione con i file Excel.
- **Moooolte ore e mooolta caffeina**.

Sì, tutto qui.

## 📊 Roadmap

Se vuoi contribuire o migliorare il progetto, ecco alcune idee:

- **Migliorare l'interfaccia utente**: Renderla più moderna e intuitiva.
- **Aggiungere crittografia**: Per proteggere ulteriormente i dati sensibili.
- **Internazionalizzazione**: Supportare più lingue.

## 👨‍💼 Licenza

Il progetto è distribuito sotto la [Licenza MIT](https://github.com/riccardomurachelli/digital_reports/blob/main/LICENSE).

---

Grazie per aver dato un'occhiata al progetto! Se hai domande o suggerimenti, non esitare a creare una **issue** o un **pull request**. Buona programmazione! 🚀

---

# 🇬🇧 English Version

## 🔍 Overview

**Digital Reports** is a desktop application designed to help Local Police forces digitally track their reports. Managing work shifts, activities performed, vehicles used, identified individuals, verified violations, and notes has never been easier!

### Key Features

- **Report Writing**: Manage shifts, activities, vehicles, individuals, violations, and notes.
- **Work Hours Calculation**: Automates hour counting.
- **Data Editing**: Update and save directly to the SQLite database.

## 🛠️ Setup and Libraries

Don't worry, the program automatically sets up the SQLite database on the first run! You'll only need the following Python libraries (if you don't have them, install them using `pip install <library_name>`):

```python
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from PIL import Image, ImageTk
import sqlite3
from datetime import datetime, timedelta
import os
import ast
import openpyxl
import locale
import webbrowser
```

## ⚠️ REDACTED: What's Missing?

For security reasons, certain files and sensitive fields have been excluded from the project. These include resources like images, Excel files for weekly or daily reports, and other configurations. By searching for the word **`REDACTED`** in the code, you'll find all the points where data or files are missing, along with a description of what they should contain.

Examples of REDACTED data:

- **Agent Fields**: Sensitive information about agent names.
- **Violations**: Specific details about violated articles.
- **Excel and PDF Reports**: Files needed to generate daily, monthly, etc., reports.

> Note: These files are not included to ensure the safety and confidentiality of sensitive police data.

## 🔧 How to Create an Executable (.exe)

If you want to distribute the program without requiring Python installation, you can use **PyInstaller** to create an `.exe` file:

1. Make sure you have PyInstaller installed:

   ```bash
   pip install pyinstaller
   ```

2. Open a terminal in the project folder and run:

   ```bash
   pyinstaller --onefile --icon=icon.ico main.py
   ```

   - `--onefile`: Creates a single executable file.
   - `--icon=icon.ico`: Adds a custom icon to the executable (if available).

3. You will find the `.exe` file in the `dist/` folder.

> Note: The executable will include all the code, so consider measures to protect sensitive information.

## 🚫 Disclaimer

The code may contain errors, whether graphical, functional, or security-related. I do not assume any responsibility for potential issues arising from its use. However, any modifications, optimizations, or improvements are more than welcome! ❤️

## 🔬 Technical Highlights

The technology powering **Digital Reports**:

- **Python**: The brain behind the project.
- **Tkinter**: For the graphical interface.
- **SQLite**: For local database management.
- **openpyxl**: For Excel file integration.
- **Loooots of hours and looooots of caffeine**.

Yep, that's it.

## 📊 Roadmap

If you want to contribute or improve the project, here are some ideas:

- **Improve the User Interface**: Make it more modern and intuitive.
- **Add Encryption**: To further secure sensitive data.
- **Internationalization**: Support for multiple languages.

## 👨‍💼 License

The project is distributed under the [MIT License](https://github.com/riccardomurachelli/digital_reports/blob/main/LICENSE).

---

Thanks for checking out the project! If you have any questions or suggestions, feel free to create an **issue** or a **pull request**. Happy coding! 🚀

