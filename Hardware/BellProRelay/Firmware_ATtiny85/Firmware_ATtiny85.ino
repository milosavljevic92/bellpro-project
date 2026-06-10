/***************************************************************************
 *  Project     : BellProUSB Relay
 *  Description : Dual relay USB interface (ATtiny85 + V-USB / DigiUSB)
 *
 *  Communication Protocol:
 *      PING   - Device identification
 *      RLY11  - Relay 1 ON
 *      RLY10  - Relay 1 OFF
 *      RLY21  - Relay 2 ON
 *      RLY20  - Relay 2 OFF
 *
 *  Hardware:
 *      MCU      : ATtiny85V (Digispark compatible)
 *      Clock    : Internal 16.5 MHz PLL
 *      USB      : V-USB (DigiUSB library)
 *      Outputs  : 2x Relay (via transistor drivers)
 *
 *  Pin Mapping:
 *      PB0 -> USB D-
 *      PB2 -> USB D+
 *      PB1 -> Relay 1
 *      PB3 -> Relay 2
 *
 *  Manufacturer : Tecomatic
 *  Location     : Novi Sad, Serbia
 *  Project Type : Embedded USB Control Interface
 *
 *  Author       : Tecomatic Development Team
 *  Version      : 1.0
 *
 *  Notes:
 *      - Do NOT disable RESET pin (required for programming)
 *      - Ensure stable 5V supply for USB reliability
 *      - Use proper transistor drivers and flyback diodes on relays
 *
 *  Copyright (c) 2026 Tecomatic.
 *  All rights reserved.
 ***************************************************************************/

#include <DigiUSB.h>

#define RELAY1_PIN 1   // PB1
#define RELAY2_PIN 3   // PB3

String cmd;

void setup()
{
    pinMode(RELAY1_PIN, OUTPUT);
    pinMode(RELAY2_PIN, OUTPUT);

    digitalWrite(RELAY1_PIN, LOW);
    digitalWrite(RELAY2_PIN, LOW);

    DigiUSB.begin();
}

void loop()
{
    DigiUSB.refresh();

    while (DigiUSB.available())
    {
        char c = DigiUSB.read();

        if (c == '\n')
        {
            cmd.trim();

            if (cmd == "PING")
            {
                DigiUSB.println("BELLPRO_OK");
            }
            else if (cmd == "RLY11")
            {
                digitalWrite(RELAY1_PIN, HIGH);
            }
            else if (cmd == "RLY10")
            {
                digitalWrite(RELAY1_PIN, LOW);
            }
            else if (cmd == "RLY21")
            {
                digitalWrite(RELAY2_PIN, HIGH);
            }
            else if (cmd == "RLY20")
            {
                digitalWrite(RELAY2_PIN, LOW);
            }

            cmd = "";
        }
        else
        {
            cmd += c;
        }
    }
}