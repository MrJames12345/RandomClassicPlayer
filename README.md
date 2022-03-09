Play a random classic song from a chosen genre.



## Setup

Add the root of your 'RandomClassicPlayer' folder to your 'Path' in environment system variables.

Open 'Edit the system envorinment variables
> Environment Variables...
> Under 'System Variables', select the variable named 'Path'
> Click 'Edit'
> Click 'Browse'
> Select your 'RandomClassicPlayer' folder
> Click 'Ok' > 'Ok' > 'Ok



## Play Random Song

Run the command 'PlayRandomClassic' manually or add this command to another script.
A random song will play from the chosen genre folder.



## Genres

Set a genre by running the 'SetGenre <num>' command.

When running 'PlayRandomClassic' when the genre is set to '0',
a random genre will be selected before then playing a random song from that genre.

If the entered genre is not an option (is negative or is higher than the number of available genres),
the genre will be set to '0'.

A genre is registered as any folder in the root directory that only has '.mp3' files.
Any folder not meeting this criteria is completely ignored from genres.