# Anke-MIDI-Reader
Uploaded 2025/6/29:<br>
Feature added: Save event list or tempo list as CSV (by right click on the list -> Save as CSV).

----------------------------------------------------------------------

A MIDI reader, player and a simple tonality analyzer, using Win32.<br>
Made by Ankedeshuqin.

/* Features */<br>
Read and play a MIDI file.<br>
--- Support for MIDI files containing non-standard chunks. RMI files are also supported.<br>
View the event list and tempo list of a MIDI file.<br>
--- The initial tempo, average tempo and duration will be indicated on the Tempo list page.<br>
Filter by track, channel and event type on the event list.<br>
Transpose and change the tempo ratio for playing.<br>
Mute tracks or channels within a track for playing.<br>
Analyze the tonality (accuracy is another matter).<br>
--- When opening a file, the number of its each note will be counted for tonality analysis (but notes of non-chromatic instruments will be excluded). You can also fill in your custom dataset of note count and then analyze the tonality. The most probable tonality will be indicated and the probability of each tonality will represented in column chart.

The Microsoft Visual Studio project file (AnkeMidi.vcxproj) is also included.

![Result](https://github.com/user-attachments/assets/956ed6d6-8aba-46ef-a71b-e5a43dee0a31)
