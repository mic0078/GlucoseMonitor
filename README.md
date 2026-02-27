# 💉 Glucose Monitor

> Monitor poziomu glukozy we krwi w czasie rzeczywistym — pasek systemowy Windows, integracja z FreeStyle Libre / LibreLinkUp

![PowerShell](https://img.shields.io/badge/PowerShell-5.1-blue?logo=powershell)
![Windows](https://img.shields.io/badge/Windows-10%2F11-0078D6?logo=windows)
![LibreLinkUp](https://img.shields.io/badge/API-LibreLinkUp%20v5-red)
![License](https://img.shields.io/badge/license-MIT-green)

---

## 📋 Opis

**Glucose Monitor** to lekka aplikacja desktopowa dla systemu Windows napisana w PowerShell + WPF,
która pobiera odczyty poziomu cukru z sensora **FreeStyle Libre** (poprzez API LibreLinkUp)
i wyświetla je dyskretnie w zasobniku systemowym (tray) oraz w małym oknie na pulpicie.

Aplikacja odświeża dane co minutę bez żadnej interakcji użytkownika.

---

## ✨ Funkcje

| Funkcja | Opis |
|--------|------|
| 🩸 **Odczyt glukozy na żywo** | Pobiera dane z API LibreLinkUp co 60 sekund |
| 📈 **Wykres historyczny** | Graf ostatnich 12 godzin w oknie głównym |
| 🔔 **Ikona tray** | Czytelna ikona kropelki krwi w zasobniku systemowym |
| 🪟 **Tryb kompaktowy** | Małe, przezroczyste okno z wartością glukozy |
| 📌 **Zawsze na wierzchu** | Opcjonalne przypięcie kompaktowego okna nad inne |
| 🔒 **Szyfrowanie hasła** | Hasło przechowywane szyfrowane przez Windows DPAPI |
| 🔄 **Autostart** | Skrypt startuje automatycznie przy logowaniu (opcja instalatora) |
| 📝 **Logi** | Debugowanie przez plik `glucose_debug.log` |

---

## 📦 Wymagania

- **Windows 10 / 11**
- **PowerShell 5.1** (wbudowany w Windows)
- Konto w aplikacji **LibreLink** lub **LibreLinkUp**
- Sensor **FreeStyle Libre** sparowany z kontem

---

## 🚀 Instalacja

### Metoda 1 — Instalator (zalecana)

1. Pobierz lub sklonuj repozytorium
2. Kliknij prawym przyciskiem na `Install_GlucoseMonitor.bat` → **Uruchom jako administrator**
3. Instalator:
   - Skopiuje pliki do `C:\Glucose\`
   - Utworzy skrót na pulpicie z ikoną kropelki krwi
   - Opcjonalnie ustawi autostart przy logowaniu

### Metoda 2 — Ręczna

```powershell
# Uruchom bezpośrednio (bez instalacji)
powershell -NoProfile -ExecutionPolicy Bypass -STA -File GlucoseMonitor.ps1
```

> ⚠️ Wymagany tryb **STA** (`-STA`), bez niego WPF nie zadziała.

---

## ⚙️ Konfiguracja

Przy pierwszym uruchomieniu pojawi się okno konfiguracji. Wprowadź:

| Pole | Wartość |
|------|---------|
| **Email** | Adres e-mail konta LibreLink / LibreLinkUp |
| **Hasło** | Hasło do konta (zostanie zaszyfrowane przez DPAPI) |

Konfiguracja zapisywana jest do pliku `config.ini` w folderze instalacji.

### Plik `config.ini`

```ini
Email=twoj@email.com
EncryptedPassword=<zaszyfrowane przez Windows DPAPI>
```

> 🔒 Hasło jest szyfrowane przy użyciu **Windows DPAPI** — plik `config.ini` jest bezpieczny,
> ale przypisany do konta Windows, na którym został utworzony.

---

## 🖥️ Obsługa

### Okno główne

```
┌─────────────────────────────────────────┐
│  🩸 Glucose Monitor          ⊟  —  ✕  │
├─────────────────────────────────────────┤
│                                         │
│         5.8 mmol/L  →                  │
│                                         │
│   [wykres ostatnich 12h]                │
│                                         │
│   Ostatni odczyt: 14:32:05             │
│   Odświeżenie za: 45 s                  │
│                                [↺]      │
└─────────────────────────────────────────┘
```

| Przycisk | Akcja |
|----------|-------|
| `⊟` | Przełącz tryb kompaktowy |
| `—` | Minimalizuj do tray |
| `✕` | Zamknij aplikację |
| `↺` | Odśwież natychmiast |

### Tryb kompaktowy

Małe, ruchome okno wyświetla tylko aktualną wartość glukozy.
Kliknij na wartość, aby przełączyć tryb **zawsze na wierzchu** (opacity zmniejszona do 55%).

### Ikona tray

| Akcja | Efekt |
|-------|-------|
| Podwójne kliknięcie | Pokaż / ukryj główne okno |
| Prawy przycisk | Menu kontekstowe (Otwórz / Ustawienia / Wyjdź) |

---

## 🩺 Interpretacja wartości

| Kolor | Zakres (mmol/L) | Znaczenie |
|-------|----------------|-----------|
| 🔴 Czerwony | < 3.9 | **Hipoglikemia** — za niski |
| 🟠 Pomarańczowy | 3.9 – 4.4 | Nisko |
| 🟢 Zielony | 4.4 – 10.0 | **Norma** |
| 🟡 Żółty | 10.0 – 13.9 | Wysoko |
| 🔴 Czerwony | > 13.9 | **Hiperglikemia** — za wysoki |

### Strzałki trendu

| Symbol | Znaczenie |
|--------|-----------|
| `↑↑` | Gwałtowny wzrost |
| `↑` | Wzrost |
| `↗` | Lekki wzrost |
| `→` | Stabilnie |
| `↘` | Lekki spadek |
| `↓` | Spadek |
| `↓↓` | Gwałtowny spadek |

---

## 📁 Struktura plików

```
C:\Glucose\
├── GlucoseMonitor.ps1          # Główny skrypt aplikacji
├── Install_GlucoseMonitor.bat  # Instalator (wymaga admina)
├── Launch_GlucoseMonitor.ps1   # Skrypt startowy (używany przez skrót)
├── config.ini                  # Konfiguracja (e-mail + zaszyfrowane hasło)
└── glucose_debug.log           # Logi debugowania
```

---

## 🐛 Rozwiązywanie problemów

| Problem | Rozwiązanie |
|---------|------------|
| Aplikacja nie pokazuje danych | Sprawdź dane logowania w Ustawieniach |
| Błąd logowania w logu | Sprawdź poprawność e-mail i hasła LibreLinkUp |
| Stara ikona na pulpicie | Uruchom instalator ponownie jako administrator |
| Puste okno / crash | Sprawdź `glucose_debug.log` w folderze instalacji |
| Brak sensora w API | Upewnij się, że sensor jest sparowany w aplikacji LibreLink |

### Logi

```powershell
# Podgląd logów na żywo
Get-Content "C:\Glucose\glucose_debug.log" -Wait -Tail 20
```

---

## 🔧 Wymagania API

Aplikacja korzysta z nieoficjalnego API **LibreLinkUp v5**:

- Endpoint: `https://api-eu.libreview.io`
- Autoryzacja: Bearer token (odświeżany automatycznie)
- Identyfikator konta: SHA256 hash `account-id`

> ℹ️ API nie jest oficjalnie udokumentowane przez Abbott.
> Działa na tych samych zasadach co aplikacja LibreLinkUp na Android/iOS.

---

## 📄 Licencja

MIT License — używaj, modyfikuj i dystrybuuj swobodnie.

---

## 👤 Autor

**Michał Nowakowski** — 2026

---

*Aplikacja stworzona do prywatnego użytku. Nie jest urządzeniem medycznym.
Zawsze konsultuj wyniki z lekarzem.*
