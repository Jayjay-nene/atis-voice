# ATIS Voice — Dictée universelle

Appuyez sur une touche, parlez, le texte apparaît au curseur — avec ponctuation et majuscules automatiques. Fonctionne dans toutes les applications. Transcription via l'API Groq (Whisper large-v3-turbo, < 1 seconde).

## Prérequis

- Python 3.10+ ([python.org](https://www.python.org/downloads/) — cocher "Add to PATH")
- pip (inclus avec Python)

## Installation

### Windows

1. Télécharger et dézipper l'archive
2. Ouvrir un terminal dans le dossier
3. `pip install -r requirements.txt`
4. `python voice_command.py` (premier lancement = configuration guidée)
5. Lancer en tant qu'administrateur pour les hotkeys globaux

### macOS

Le script utilise des APIs Windows (`winsound`, `keyboard`, `ctypes.windll`) et n'est pas compatible macOS en l'état. Un portage est envisagé. En attendant : voir [Whisper Dictation](https://github.com/foges/whisper-dictation) pour macOS.

### Linux

Même limitation que macOS — les dépendances Windows ne sont pas disponibles. Un portage est envisagé.

## Configuration

Au premier lancement, un assistant interactif crée le fichier `.env` automatiquement.

Pour reconfigurer à tout moment : `python voice_command.py --setup`

**Clé API Groq (gratuite) :** [console.groq.com/keys](https://console.groq.com/keys) — créer un compte, générer une clé, la coller quand l'assistant la demande. Aucune carte bancaire requise.

| Variable | Description |
|----------|-------------|
| `GROQ_API_KEY` | Clé API Groq (`gsk_...`) |
| `HOTKEY_KEY` | Touche de déclenchement (défaut : `CapsLock`) |
| `NOTES_PATH` | Dossier pour les notes vocales (défaut : `~/Documents/ATIS-Voice-Notes`) |

## Groq — pourquoi ce choix

Groq (sans k) n'a rien à voir avec Grok, l'IA d'Elon Musk — ce sont deux choses sans lien. Groq est une entreprise de hardware qui a conçu les LPU (Language Processing Units), des puces dédiées à l'inférence rapide, contrairement aux GPU classiques. Résultat : la transcription arrive en moins d'une seconde. L'API est gratuite sans CB, avec un quota mensuel largement suffisant pour un usage personnel. Côté empreinte : mutualiser la transcription sur des serveurs optimisés consomme moins qu'un modèle local qui tourne en permanence sur votre CPU.

## Utilisation

| Geste | Action |
|-------|--------|
| `{HOTKEY_KEY}` maintenu > 0.3s | Push-to-talk — relâcher = texte collé au curseur |
| `{HOTKEY_KEY}` tap | Bascule ON/OFF — texte collé au curseur |
| `Shift+{HOTKEY_KEY}` | Bascule ON/OFF — texte dans le clipboard (collage manuel) |
| `Alt+{HOTKEY_KEY}` | Note vocale sauvegardée en fichier `.md` |

`{HOTKEY_KEY}` = la touche configurée dans `.env` (défaut : `CapsLock`).

**Mode offline (sans internet) :**

```
pip install faster-whisper
python voice_command.py --offline
```

Transcription 100% locale avec faster-whisper medium int8. Plus lent (~4-6s CPU) mais fonctionne sans connexion.

## Licence

MIT — Jules Neny / Transformations résilientes
