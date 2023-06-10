#define portLed 13

void setup() {
	pinMode(portLed, OUTPUT);

}

void loop() {
	digitalWrite(portLed, HIGH);
	delay(100);
	digitalWrite(portLed, LOW);
	delay(100);

}
