/***************************************************************************
 *  Project     : BellProRelay
 *  Description : Dual relay USB interface controlled via serial commands.
 *
 *  Commands:
 *      PING   - Device identification
 *      RLY11  - Relay 1 ON
 *      RLY10  - Relay 1 OFF
 *      RLY21  - Relay 2 ON
 *      RLY20  - Relay 2 OFF
 *
 *  Manufacturer : Tecomatic
 *  Location     : Novi Sad, Serbia
 *  Author       : Tecomatic Development Team
 *  Version      : 1.0
 *
 *  Copyright (c) 2025 Tecomatic.
 *  All rights reserved.
 ***************************************************************************/
 
#define RELAY1_PIN 7
#define RELAY2_PIN 8

String cmd;

void setup()
{
    pinMode(RELAY1_PIN, OUTPUT);
    pinMode(RELAY2_PIN, OUTPUT);

    digitalWrite(RELAY1_PIN, LOW);
    digitalWrite(RELAY2_PIN, LOW);

    Serial.begin(9600);
}

void loop()
{
    if (!Serial.available())
        return;

    cmd = Serial.readStringUntil('\n');
    cmd.trim();

    if (cmd == "PING")
    {
        Serial.println("BELLRLY");
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
}