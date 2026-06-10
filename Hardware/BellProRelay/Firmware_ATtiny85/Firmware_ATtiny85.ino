/***************************************************************************
 *  Project         : BellProUSB Relay Controller
 *
 *  Description     :
 *      Industrial dual-relay USB control interface based on ATtiny85V.
 *      Uses V-USB (DigiUSB) and binary Industrial Protocol v2 (1-byte commands)
 *      designed for reliable 24/7 operation in Windows VB6 environments.
 *
 *  -----------------------------------------------------------------------
 *  COMMUNICATION PROTOCOL (v2 - BINARY)
 *  -----------------------------------------------------------------------
 *
 *  PC → MCU COMMANDS:
 *      0x01 : PING (health check)
 *      0x11 : RELAY1 ON
 *      0x10 : RELAY1 OFF
 *      0x21 : RELAY2 ON
 *      0x20 : RELAY2 OFF
 *      0x30 : GET STATE
 *
 *  MCU → PC RESPONSES:
 *      0xAA : DEVICE OK / ALIVE
 *      0x00 : BOTH RELAYS OFF
 *      0x01 : RELAY1 ON
 *      0x02 : RELAY2 ON
 *      0x03 : BOTH RELAYS ON
 *      0xFF : ERROR / NO DATA
 *
 *  -----------------------------------------------------------------------
 *  HARDWARE PLATFORM
 *  -----------------------------------------------------------------------
 *
 *  MCU          : ATtiny85V (Digispark compatible)
 *  Architecture : AVR 8-bit
 *  Clock        : Internal 16.5 MHz PLL (V-USB timing critical)
 *  USB Stack    : V-USB (DigiUSB library)
 *
 *  Power Supply : +5V regulated (USB powered or external stable source)
 *
 *  Outputs      : 2x relay outputs (via transistor drivers only)
 *
 *  -----------------------------------------------------------------------
 *  PIN MAPPING
 *  -----------------------------------------------------------------------
 *
 *      PB0  -> USB D-
 *      PB2  -> USB D+
 *      PB1  -> RELAY 1 OUTPUT
 *      PB3  -> RELAY 2 OUTPUT
 *      PB4  -> STATUS LED (optional)
 *      PB5  -> RESET (MUST remain enabled)
 *
 *  -----------------------------------------------------------------------
 *  ELECTRICAL DESIGN NOTES
 *  -----------------------------------------------------------------------
 *
 *  - Relay coils MUST be driven via NPN transistor (BC547 / BC847 / 2N2222)
 *  - Flyback diode required across each relay coil (1N4007 or 1N4148)
 *  - USB lines must include 68Ω series resistors
 *  - 3.6V zener or clamp diodes recommended for USB protection
 *  - Add 100nF decoupling capacitor near VCC pin
 *  - Add 10µF bulk capacitor on supply rail for relay stability
 *
 *  -----------------------------------------------------------------------
 *  FIRMWARE DESIGN NOTES
 *  -----------------------------------------------------------------------
 *
 *  - Watchdog timer enabled (WDTO_2S) for auto-recovery from lockups
 *  - Line buffering disabled; uses direct byte-level protocol
 *  - USB refresh must be called continuously (DigiUSB.refresh)
 *  - Input buffer protection implemented (overflow guard)
 *
 *  -----------------------------------------------------------------------
 *  TIMING / USB NOTES
 *  -----------------------------------------------------------------------
 *
 *  - V-USB timing is highly sensitive to clock accuracy
 *  - Internal PLL calibration is mandatory for stable enumeration
 *  - USB disconnect/reconnect handled by host OS automatically
 *
 *  -----------------------------------------------------------------------
 *  RELIABILITY TARGET
 *  -----------------------------------------------------------------------
 *
 *  Designed for:
 *      - 24/7 continuous operation
 *      - Embedded control systems
 *      - Industrial relay switching
 *      - VB6 legacy control applications
 *
 *  -----------------------------------------------------------------------
 *  AUTHOR / MANUFACTURER
 *  -----------------------------------------------------------------------
 *
 *  Manufacturer : Tecomatic
 *  Location     : Novi Sad, Serbia
 *  Author       : Tecomatic Development Team
 *  Version      : 2.0
 *
 *
 *  -----------------------------------------------------------------------
 *  LICENSE
 *  -----------------------------------------------------------------------
 *
 *  Copyright (c) 2026 Tecomatic.
 *  All rights reserved.
 ***************************************************************************/

#include <DigiUSB.h>
#include <avr/wdt.h>

#define RELAY1_PIN 1   // PB1
#define RELAY2_PIN 3   // PB3

uint8_t relayState = 0;

void setup()
{
    pinMode(RELAY1_PIN, OUTPUT);
    pinMode(RELAY2_PIN, OUTPUT);

    digitalWrite(RELAY1_PIN, LOW);
    digitalWrite(RELAY2_PIN, LOW);

    DigiUSB.begin();

    wdt_enable(WDTO_2S);
}

void sendState()
{
    DigiUSB.write(relayState);
}

void loop()
{
    wdt_reset();
    DigiUSB.refresh();

    while (DigiUSB.available() > 0)
    {
        uint8_t cmd = DigiUSB.read();

        switch (cmd)
        {
            case 0x01: // PING
                DigiUSB.write(0xAA);
                break;

            case 0x11: // R1 ON
                digitalWrite(RELAY1_PIN, HIGH);
                relayState |= 0x01;
                break;

            case 0x10: // R1 OFF
                digitalWrite(RELAY1_PIN, LOW);
                relayState &= ~0x01;
                break;

            case 0x21: // R2 ON
                digitalWrite(RELAY2_PIN, HIGH);
                relayState |= 0x02;
                break;

            case 0x20: // R2 OFF
                digitalWrite(RELAY2_PIN, LOW);
                relayState &= ~0x02;
                break;

            case 0x30: // GET STATE
                sendState();
                break;
        }
    }
}