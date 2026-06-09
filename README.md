# BellPro: School Bell Automation System
BellPro is a robust, open-source software solution designed to automate school bell ringing. By interfacing a PC with existing school PA systems via a custom RS232 hardware interface, BellPro eliminates manual operation and ensures high-precision timing for classes, breaks, and extracurricular activities.
## Overview
Originally conceptualized in 2007 and refined over nearly two decades, BellPro has been a staple in numerous schools across Serbia. This repository provides the complete ecosystem needed to deploy the system, including the VB6 source code, PCB designs for the relay interface, and installation scripts.
## Key Features
* **Unlimited Scheduling:** Create and manage an infinite number of bell schedules.
* **Smart Calendar Logic:** Automated handling of weekends, public holidays, and seasonal school breaks.
* **Extracurricular Support:** Easily switch between different schedules for sports events or evening activities.
* **System Resilience:** Built-in auto-start and ability to recover operations after power outages or reboots.
* **Password Protection:** Secure the configuration to prevent unauthorized modifications.
* **Data Portability:** Full support for importing and exporting schedules.
## Technical Architecture
* **Core:** Developed in Visual Basic 6 (VB6).
* **Database:** SQLite (using ODBC connector for integration).
* **Communication:** RS232 serial interface.
    * **Triggering:** Pin 7 (RTS) drives the relay output.
    * **Detection:** Pin 4 (DTR) → Pin 8 (CTS) loopback detects interface presence.
## Hardware Setup
The system relies on a physical relay interface connected to the school's sound system.
* **Serial Port:** Native serial ports (COM ports) on motherboards are highly recommended for stability.
* **USB-to-Serial:** If using a USB adapter, performance depends heavily on the chipset. We recommend high-quality adapters based on the **FTDI FT232RNL** chip (e.g., Digitus).
https://de.assmann.shop/en/Cables-and-Peripherals/Computer-Accessories/IO-Cards/USB-2-0-serial-adapter.html
* **Caution:** Note that some motherboards may toggle RS232 signals during BIOS post/restart, which could cause a brief relay click.
## Installation Guide
1. **Prerequisites:** Install the SQLite ODBC driver (`sqliteodbc.exe`) provided in the `Tools` folder.
2. **Library Registration:** If you encounter "Component not found" errors, manually register the required OCX files. Specifically, for the DataGrid, navigate to `InstallScript -> dll` and register `msdatgrd.ocx`.
3. **Configuration:** Use the included KeyGenerator to authorize your specific hardware build.
## Roadmap & Future Development
While the system remains stable, we encourage the community to help modernize the codebase. Potential areas for contribution include:
* **Microcontroller Integration:** Replacing the RS232/PC-dependent design with a modern Atmel or PIC microcontroller.
## Contributing
We want to keep this project alive! If you decide to install BellPro in a school or organization, please let us know or open an issue to document the installation. Contributions in the form of code improvements, documentation, or hardware design forks are highly encouraged.
## License
BellPro is released under the [MIT License](LICENSE).
You are free to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the software, provided the original copyright notice and this permission notice
are included in all copies or substantial portions of the software.
---
*Inspired by the legendary school bell projects of the past, specifically the work of Voja Milanović. http://vojo.milanovic.org/zvono.htm*
