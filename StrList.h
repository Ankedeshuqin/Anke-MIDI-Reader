/* Note list */
LPWSTR alpszNote[] = {
    L"C0", L"C#0", L"D0", L"D#0", L"E0", L"F0", L"F#0", L"G0",
    L"G#0", L"A0", L"A#0", L"B0", L"C1", L"C#1", L"D1", L"D#1",
    L"E1", L"F1", L"F#1", L"G1", L"G#1", L"A1", L"A#1", L"B1",
    L"C2", L"C#2", L"D2", L"D#2", L"E2", L"F2", L"F#2", L"G2",
    L"G#2", L"A2", L"A#2", L"B2", L"C3", L"C#3", L"D3", L"D#3",
    L"E3", L"F3", L"F#3", L"G3", L"G#3", L"A3", L"A#3", L"B3",
    L"C4", L"C#4", L"D4", L"D#4", L"E4", L"F4", L"F#4", L"G4",
    L"G#4", L"A4", L"A#4", L"B4", L"C5", L"C#5", L"D5", L"D#5",
    L"E5", L"F5", L"F#5", L"G5", L"G#5", L"A5", L"A#5", L"B5",
    L"C6", L"C#6", L"D6", L"D#6", L"E6", L"F6", L"F#6", L"G6",
    L"G#6", L"A6", L"A#6", L"B6", L"C7", L"C#7", L"D7", L"D#7",
    L"E7", L"F7", L"F#7", L"G7", L"G#7", L"A7", L"A#7", L"B7",
    L"C8", L"C#8", L"D8", L"D#8", L"E8", L"F8", L"F#8", L"G8",
    L"G#8", L"A8", L"A#8", L"B8", L"C9", L"C#9", L"D9", L"D#9",
    L"E9", L"F9", L"F#9", L"G9", L"G#9", L"A9", L"A#9", L"B9",
    L"C10", L"C#10", L"D10", L"D#10", L"E10", L"F10", L"F#10", L"G10"
};

/* Controller list */
LPWSTR alpszCtl[] = {
    L"Bank select MSB", L"Modulation wheel", L"Breath controller", L"", L"Foot controller", L"Portamento time", L"Data entry MSB", L"Volume",
    L"Balance", L"", L"Pan", L"Expression", L"Effect control 1", L"Effect control 2", L"", L"",
    L"General purpose controller 1", L"General purpose controller 2", L"General purpose controller 3", L"General purpose controller 4", L"", L"", L"", L"",
    L"", L"", L"", L"", L"", L"", L"", L"",
    L"Bank select LSB", L"Modulation wheel LSB", L"Breath controller LSB", L"", L"Foot controller LSB", L"Portamento time LSB", L"Data entry LSB", L"Volume LSB",
    L"Balance LSB", L"", L"Pan LSB", L"Expression LSB", L"Effect control 1 LSB", L"Effect control 2 LSB", L"", L"",
    L"General purpose controller 1 LSB", L"General purpose controller 2 LSB", L"General purpose controller 3 LSB", L"General purpose controller 4 LSB", L"", L"", L"", L"",
    L"", L"", L"", L"", L"", L"", L"", L"",
    L"Pedal (sustain)", L"Portamento", L"Pedal (sostenuto)", L"Pedal (soft)", L"Pedal (legato)", L"Pedal (sustain 2)", L"Sound variation", L"Resonance",
    L"Sound release time", L"Sound attack time", L"Brightness", L"Decay time", L"Vibrato rate", L"Vibrato depth", L"Vibrato delay", L"",
    L"General purpose controller 5", L"General purpose controller 6", L"General purpose controller 7", L"General purpose controller 8", L"", L"", L"", L"",
    L"", L"", L"", L"Reverb depth", L"Tremolo depth", L"Chorus depth", L"Celeste depth", L"Phaser depth",
    L"Data entry increment", L"Data entry decrement", L"NRPN LSB", L"NRPN MSB", L"RPN LSB", L"RPN MSB", L"", L"",
    L"", L"", L"", L"", L"", L"", L"", L"",
    L"", L"", L"", L"", L"", L"", L"", L"",
    L"All sounds off", L"All controllers off", L"Local keyboard", L"All notes off", L"Omni mode off", L"Omni mode on", L"Mono mode", L"Poly mode"
};

/* Program change list */
LPWSTR alpszPrg[] = {
    L"Acoustic Grand Piano", L"Bright Acoustic Piano", L"Electric Grand Piano", L"Honky-tonk Piano", L"Rhodes Piano", L"Chorused Piano", L"Harpsichord", L"Clavinet",
    L"Celesta", L"Glockenspiel", L"Music Box", L"Vibraphone", L"Marimba", L"Xylophone", L"Tubular Bell", L"Dulcimer",
    L"Hammond Organ", L"Percussive Organ", L"Rock Organ", L"Church Organ", L"Reed Organ", L"Accordion", L"Harmonica", L"Tango Accordion",
    L"Acoustic Guitar (nylon)", L"Acoustic Guitar (steel)", L"Electric Guitar (jazz)", L"Electric Guitar (clean)", L"Electric Guitar (muted)", L"Overdriven Guitar", L"Distortion Guitar", L"Guitar Harmonics",
    L"Acoustic Bass", L"Electric Bass (finger)", L"Electric Bass (pick)", L"Fretless Bass", L"Slap Bass 1", L"Slap Bass 2", L"Synth Bass 1", L"Synth Bass 2",
    L"Violin", L"Viola", L"Cello", L"Contrabass", L"Tremolo Strings", L"Pizzicato Strings", L"Orchestral Harp", L"Timpani",
    L"String Ensemble 1", L"String Ensemble 2", L"Synth Strings 1", L"Synth Strings 2", L"Choir Aahs", L"Voice Oohs", L"Synth Voice", L"Orchestra Hit",
    L"Trumpet", L"Trombone", L"Tuba", L"Muted Trumpet", L"French Horn", L"Brass Section", L"Synth Brass 1", L"Synth Brass 2",
    L"Soprano Sax", L"Alto Sax", L"Tenor Sax", L"Baritone Sax", L"Oboe", L"English Horn", L"Bassoon", L"Clarinet",
    L"Piccolo", L"Flute", L"Recorder", L"Pan Flute", L"Bottle Blow", L"Shakuhachi", L"Whistle", L"Ocarina",
    L"Lead 1 (square)", L"Lead 2 (sawtooth)", L"Lead 3 (calliope lead)", L"Lead 4 (chiff lead)", L"Lead 5 (charang)", L"Lead 6 (voice)", L"Lead 7 (fifths)", L"Lead 8 (bass + lead)",
    L"Pad 1 (new age)", L"Pad 2 (warm)", L"Pad 3 (polysynth)", L"Pad 4 (choir)", L"Pad 5 (bowed)", L"Pad 6 (metallic)", L"Pad 7 (halo)", L"Pad 8 (sweep)",
    L"FX 1 (rain)", L"FX 2 (soundtrack)", L"FX 3 (crystal)", L"FX 4 (atmosphere)", L"FX 5 (brightness)", L"FX 6 (goblins)", L"FX 7 (echoes)", L"FX 8 (sci-fi)",
    L"Sitar", L"Banjo", L"Shamisen", L"Koto", L"Kalimba", L"Bagpipe", L"Fiddle", L"Shanai",
    L"Tinkle Bell", L"Agogo", L"Steel Drums", L"Woodblock", L"Taiko Drum", L"Melodic Tom", L"Synth Drum", L"Reverse Cymbal",
    L"Guitar Fret Noise", L"Breath Noise", L"Seashore", L"Bird Tweet", L"Telephone Ring", L"Helicopter", L"Applause", L"Gunshot"
};

/* Meta event list */
LPWSTR alpszMeta[] = {
    L"Sequence number", L"Text", L"Copyright", L"Track name", L"Instrument name", L"Lyric", L"Marker", L"Cue point",
    L"Program name", L"Device name", L"", L"", L"", L"", L"", L"",
    L"", L"", L"", L"", L"", L"", L"", L"",
    L"", L"", L"", L"", L"", L"", L"", L"",
    L"Channel prefix", L"MIDI port", L"", L"", L"", L"", L"", L"",
    L"", L"", L"", L"", L"", L"", L"", L"End of track",
    L"", L"", L"", L"", L"", L"", L"", L"",
    L"", L"", L"", L"", L"", L"", L"", L"",
    L"", L"", L"", L"", L"", L"", L"", L"",
    L"", L"", L"", L"", L"", L"", L"", L"",
    L"", L"Tempo", L"", L"", L"SMPTE offset", L"", L"", L"",
    L"Time signature", L"Key signature", L"", L"", L"", L"", L"", L"",
    L"", L"", L"", L"", L"", L"", L"", L"",
    L"", L"", L"", L"", L"", L"", L"", L"",
    L"", L"", L"", L"", L"", L"", L"", L"",
    L"", L"", L"", L"", L"", L"", L"", L"Sequencer specific"
};