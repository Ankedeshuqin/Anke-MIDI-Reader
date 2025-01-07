# Anke-MIDI-Reader
A MIDI reader, player and a simple tonality analyzer, using Win32.
Made by Ankedeshuqin.

/* Features */
Read and play a MIDI file.
--- Support for MIDI files containing non-standard chunks. RMI files are also supported.
View the event list and tempo list of a MIDI file.
--- The initial tempo, average tempo and duration will be indicated on the Tempo list page.
Filter by track, channel and event type on the event list.
Transpose and change the tempo ratio for playing.
Mute tracks or channels within a track for playing.
Analyze the tonality (accuracy is another matter).
--- When opening a file, the number of its each note will be counted for tonality analysis (but notes of non-chromatic instruments will be excluded). You can also fill in your custom dataset of note count and then analyze the tonality. The most probable tonality will be indicated and the probability of each tonality will represented in column chart.

The Microsoft Visual Studio project file (AnkeMidi.vcxproj) is also included.
