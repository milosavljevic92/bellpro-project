# BellPro

BellPro is a software solution for automating school bells. The project enables precise control of ringing schedules via computer, integrating seamlessly with existing school PA systems and other organizational facilities.

## About the Project

This project was created to replace outdated, manual bell systems with a modern, automated solution. Developed since 2007, BellPro has proven to be a reliable system, currently active in over 10 schools across Serbia.

This repository contains the full source code (VB6), documentation, PCB interface schematics, as well as the tools required for installation and maintenance.

## Key Features

* **Flexible Scheduling:** Unlimited number of schedules with custom time entries.
* **Extracurricular Support:** Dedicated schedules for sports halls and other non-academic activities.
* **Smart Planning:** Automatic disabling of bells during weekends, holidays, and school breaks.
* **Reliability:** Auto-start with the system, resilient to power outages and computer reboots.
* **Security:** Password protection to prevent unauthorized changes.
* **Compatibility:** Import/Export functionality for schedules.

## Technical Specifications

* **Programming Language:** Visual Basic 6 (VB6).
* **Database:** SQLite.
* **Interface:** RS232 with relay (triggering via RTS - pin 7, interface detection via DTR - pin 4).

## Installation & Development

To compile and run the program in the VB6 environment, you must:

1.  **SQLite Connector:** Install `sqliteodbc.exe` (located in the `Tools` folder).
2.  **Dependencies:** If you encounter issues loading the `MSDataGridLib.DataGrid` DLL, manually import it from `InstallScript -> dll -> msdatgrd.ocx`.
3.  **Hardware:** The system performs best with native serial ports on the motherboard. If using USB-to-Serial converters, we highly recommend using those based on the **FTDI (e.g., FT232RNL)** chip (e.g., Digitus).

> **Note:** During a computer restart, a brief signal may be sent to the RS232 port, which might cause a momentary ring.

## Future Development

This is an open-source project, and we welcome community contributions. Ideas for future improvements include:

* Migrating the interface to modern microcontrollers (Atmel/PIC).
* Developing an "offline" version using an RTC (Real Time Clock) module and EEPROM.
* Code modernization and migration to newer frameworks.

## License

This software is completely free to use. Please feel free to install, use, and modify it. If you install it in an institution, feel free to let us know so we can track active installations.

---
*Project developed with dedication and a passion for open-source software. Original inspiration drawn from the work of Voja Milanović.*