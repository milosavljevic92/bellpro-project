# BellPro: School Bell Automation System

BellPro is a robust, open-source software solution designed to automate school bell ringing. By interfacing a PC with existing school PA systems via a custom RS232 hardware interface, BellPro eliminates manual operation and ensures high-precision timing for classes, breaks, and extracurricular activities.

> Originally conceptualized in 2007 and refined over nearly two decades, BellPro has been deployed in numerous schools across Serbia. This repository provides the complete ecosystem needed to deploy the system, including the VB6 source code, PCB designs for the relay interface, and installation scripts.

---

## Key Features

- **Unlimited Scheduling** — Create and manage an infinite number of named bell schedules.
- **Smart Calendar Logic** — Automated handling of weekends, public holidays, and seasonal school breaks.
- **Extracurricular Support** — Easily switch between schedules for sports events, evening activities, or special occasions.
- **System Resilience** — Built-in auto-start and automatic recovery after power outages or reboots.
- **Password Protection** — Secure the configuration with a 4-digit PIN to prevent unauthorized changes.
- **Data Portability** — Full support for importing and exporting schedules via XML.
- **Multi-Language UI** — Interface available in 10 languages, switchable at runtime without restart.

---

## Multi-Language Support

The UI language can be changed live from **Settings → Language** dropdown. The selection is saved and restored on next launch.

Supported languages:

| Language | File |
|----------|------|
| Serbian *(default)* | `Serbian_language.lng` |
| English | `English_language.lng` |
| German | `German_language.lng` |
| French | `French_language.lng` |
| Italian | `Italian_language.lng` |
| Spanish | `Spanish_language.lng` |
| Slovenian | `Slovenian_language.lng` |
| Hungarian | `Hungarian_language.lng` |
| Greek *(transliterated)* | `Greek_language.lng` |
| Macedonian *(transliterated)* | `Macedonian_language.lng` |

Language files are plain INI-style `.lng` files located in the `language\` folder next to the executable. Adding a new language is as simple as copying an existing file, renaming it `<LanguageName>_language.lng`, and translating the values — no recompilation required.

### Adding a Custom Language

1. Copy `language\English_language.lng` to `language\MyLanguage_language.lng`.
2. Translate all values (text after `=`). Do not change the keys (text before `=`).
3. Launch BellPro — the new language will appear automatically in the dropdown.

> **Note:** Language files must be saved in **Windows-1250 encoding with CRLF line endings**. Languages with non-Latin scripts (Cyrillic, Greek, Arabic, etc.) should use Latin transliteration for compatibility with the VB6 runtime.

---

## Technical Architecture

- **Core:** Visual Basic 6 (VB6)
- **Database:** SQLite via ODBC connector
- **Communication:** RS232 serial interface
  - **Triggering:** Pin 7 (RTS) drives the relay output
  - **Detection:** Pin 4 (DTR) → Pin 8 (CTS) loopback detects interface presence
- **Localization:** INI-style `.lng` files, loaded at runtime via `MdlLanguage.bas`

---

## Hardware Setup

The system relies on a physical relay interface connected to the school's sound system.

- **Serial Port:** Native COM ports on the motherboard are strongly recommended for timing stability.
- **USB-to-Serial:** Supported, but performance depends on the chipset. We recommend adapters based on the **FTDI FT232RNL** chip (e.g., [Digitus DA-70156](https://de.assmann.shop/en/Cables-and-Peripherals/Computer-Accessories/IO-Cards/USB-2-0-serial-adapter.html)).
- **Relay Interface:** Optocoupler-isolated RS232 relay board. PCB design files are included in the `Hardware\` folder.

> **Caution:** Some motherboards toggle RS232 signals during BIOS POST or restart, which may cause a brief unintended relay click.

---

## Installation Guide

1. **SQLite ODBC Driver** — Install `sqliteodbc.exe` from the `Tools\` folder before first launch.
2. **Component Registration** — If you see *"Component not found"* errors, register the required OCX files manually. For the DataGrid control, navigate to `InstallScript\dll\` and register `msdatgrd.ocx`.
3. **Authorization** — Use the included KeyGenerator utility to generate a license key tied to your hardware's volume serial number.
4. **Language Files** — Ensure the `language\` folder is present in the same directory as `BellPro.exe`. The folder ships with all 10 language files.

---

## Project Structure

```
BellPro/
├── Form/               VB6 form files (.frm)
├── Module/             VB6 code modules (.bas), including MdlLanguage.bas
├── Control/            Custom VB6 user controls (.ctl) — serial, slider, button
├── Class/              VB6 class modules (.cls)
├── language/           Localization files (*_language.lng)
├── Img/                Application icons and status images
├── Tools/              SQLite ODBC driver, KeyGenerator
├── Hardware/           RS232 relay interface PCB design (Sprint Layout)
├── base.sqlite         Application database
├── config.ini          Runtime configuration
└── BellPro.vbp         VB6 project file
```

---

## Roadmap & Future Development

While the system remains stable and production-ready, community contributions are welcome. Areas for potential improvement:

- **Microcontroller Migration** — Replace the RS232/PC-dependent design with a standalone embedded solution (AVR, PIC, or ESP32-based).
- **Additional Languages** — Contribute new `.lng` translation files for other languages.
- **Modern Installer** — An NSIS or Inno Setup installer script to simplify first-time deployment.
- **Logging** — Ring event log with timestamps for audit purposes.

---

## Contributing

If you deploy BellPro in a school or organization, please open an issue to document the installation — it helps track the project's reach and impact.

Contributions in the form of code improvements, translation files, documentation updates, or hardware design forks are all encouraged.

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/my-improvement`)
3. Commit your changes (`git commit -m 'Add: description'`)
4. Open a Pull Request

---

## License

BellPro is released under the [MIT License](LICENSE) — free to use, modify, and redistribute, provided the original copyright notice is retained in all copies or substantial portions of the software.

---

*Developed and maintained by [Tecomatic](https://tecomatic.rs) — Novi Sad, Serbia.*  
*Inspired by the school bell projects of the past, specifically the work of Voja Milanović: http://vojo.milanovic.org/zvono.htm*
